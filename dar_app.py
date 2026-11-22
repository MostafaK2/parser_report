#!/usr/bin/env python3
"""
DAR review app.

    python dar_app.py                 # then File > Open report
    python dar_app.py report.xlsx     # open a report straight away

Opens a shift report (read only - the Excel file is never changed), lists every
tour / lock / escort that needs an officer, and builds the DAR as you pick.
Rows with only one person on the site are filled in for you; rows with two or
more people start blank and must be picked before the DAR can be copied or saved.
"""
from __future__ import annotations

import json
import os
import sys

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:
    _v = f"{sys.version_info.major}.{sys.version_info.minor}"
    sys.exit(
        "\nThis Python doesn't include tkinter, which the app window needs.\n"
        "tkinter is part of Python, so it can't be installed with pip. Fix:\n"
        f"  Mac (Homebrew Python):  brew install python-tk@{_v}\n"
        "  Ubuntu/Debian Linux:    sudo apt install python3-tk\n"
        "  Or install Python from python.org, which includes it.\n"
        f"\n(Python in use: {sys.executable}, version {_v})\n"
        "The command-line version still works without it:  python3 shift_report_parser.py report.xlsx\n"
    )

_SRC = os.path.dirname(os.path.abspath(__file__))
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from tkinter import font as tkfont  # noqa: E402

from dar_session import SEPARATOR, DarSession  # noqa: E402
from shift_report_parser import BLANK, DEFAULT_CONFIG, ReportError, app_dir, default_settings, load_config  # noqa: E402

HERE = app_dir()

APP_TITLE = "Shift Report → DAR"
SETTINGS_FILE = os.path.join(HERE, "dar_settings.json")
SETTINGS_FALLBACK = os.path.join(os.path.expanduser("~"), ".dar_settings.json")

MODE_TEXT = {"ask": "Ask me", "rotate": "Rotate", "random": "Random", "auto": "Auto"}
MODE_HINT = "Rotate: 1st tour officer, 2nd Lt.   Auto: officer first.\nSettings never change hand picks."

COLORS = {
    "single": "#6b7280", "rule": "#1d4ed8", "report": "#1d4ed8", "manual": "#7c3aed",
    "prefill": "#0f766e", "missing": "#b91c1c", "ok": "#15803d",
}


class ScrollFrame(ttk.Frame):
    """A frame with a vertical scrollbar; put children in .inner."""

    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")
        for w in (self.canvas, self.inner):
            w.bind("<Enter>", self._bind_wheel)
            w.bind("<Leave>", self._unbind_wheel)

    def _bind_wheel(self, _e=None):
        self.bind_all("<MouseWheel>", self.on_wheel)
        self.bind_all("<Button-4>", self.on_wheel)
        self.bind_all("<Button-5>", self.on_wheel)

    def _unbind_wheel(self, _e=None):
        self.unbind_all("<MouseWheel>")
        self.unbind_all("<Button-4>")
        self.unbind_all("<Button-5>")

    def on_wheel(self, e):
        if self.canvas.yview() == (0.0, 1.0):
            return "break"
        if getattr(e, "num", None) == 4:
            step = -1
        elif getattr(e, "num", None) == 5:
            step = 1
        else:
            d = e.delta
            step = -1 * (d // 120 if abs(d) >= 120 else (1 if d > 0 else -1))
        self.canvas.yview_scroll(step, "units")
        return "break"

    def scroll_top(self):
        self.canvas.yview_moveto(0)


class DarApp(tk.Tk):
    def __init__(self, report_path=None, config_path=DEFAULT_CONFIG):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1180x760")
        self.minsize(900, 600)
        self.config_path = config_path
        try:
            self.cfg = load_config(config_path)
        except (OSError, ValueError) as e:
            messagebox.showerror(APP_TITLE, f"Couldn't read the config file:\n{config_path}\n\n{e}")
            self.destroy()
            raise SystemExit(1)
        self.settings = self._load_settings()
        self.session: DarSession | None = None
        self.rows = {}               # task key -> dict of widgets
        self.only_choose = tk.BooleanVar(value=False)
        self.mode_vars = {}
        self.shuffle_buttons = {}
        self._save_warned = False

        self._style()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.bind("<Control-o>", lambda e: self.open_dialog())
        self.bind("<F5>", lambda e: self.reload())
        if report_path:
            self.after(50, lambda: self.open_report(report_path))
        else:
            self._render_empty("Open a shift report to start  (File ▸ Open report, or Ctrl+O).")

    # ------------------------------------------------------------- layout
    def _style(self):
        st = ttk.Style(self)
        if "clam" in st.theme_names():
            st.theme_use("clam")
        base = tkfont.nametofont("TkDefaultFont")
        self.f_head = base.copy()
        self.f_head.configure(weight="bold")
        self.f_title = base.copy()
        self.f_title.configure(weight="bold", size=base.actual("size") + 2)
        mono = tkfont.nametofont("TkFixedFont")
        self.f_mono = mono.copy()
        self.f_mono.configure(size=max(10, base.actual("size")))
        self.f_mono_bold = self.f_mono.copy()
        self.f_mono_bold.configure(weight="bold")
        st.configure("Head.TLabel", font=self.f_head)
        st.configure("Title.TLabel", font=self.f_title)
        st.map("Missing.TCombobox", fieldbackground=[("readonly", "#fde2e2")])
        st.map("Manual.TCombobox", fieldbackground=[("readonly", "#efe7ff")])
        st.configure("Small.TButton", padding=(4, 0))

    def _build(self):
        menubar = tk.Menu(self)
        filem = tk.Menu(menubar, tearoff=False)
        filem.add_command(label="Open report…", accelerator="Ctrl+O", command=self.open_dialog)
        filem.add_command(label="Reload report", accelerator="F5", command=self.reload)
        filem.add_separator()
        filem.add_command(label="Clear all hand picks", command=self.clear_manual)
        filem.add_separator()
        filem.add_command(label="Quit", command=self.on_close)
        menubar.add_cascade(label="File", menu=filem)
        self.config(menu=menubar)

        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill="x")
        ttk.Button(top, text="Open report…", command=self.open_dialog).pack(side="left")
        self.reload_btn = ttk.Button(top, text="Reload", command=self.reload, state="disabled")
        self.reload_btn.pack(side="left", padx=(6, 0))
        self.file_lbl = ttk.Label(top, text="No report open", style="Title.TLabel")
        self.file_lbl.pack(side="left", padx=14)
        self.shift_lbl = ttk.Label(top, text="")
        self.shift_lbl.pack(side="right")

        info = ttk.Frame(self, padding=(10, 0))
        info.pack(fill="x")
        sbox = ttk.LabelFrame(info, text="Settings (only for sites with 2+ people)", padding=8)
        sbox.pack(side="right", fill="y", padx=(10, 0))
        self.duty_box = ttk.LabelFrame(info, text="On duty", padding=8)
        self.duty_box.pack(side="left", fill="both", expand=True)
        for r, (cat, s) in enumerate(self.cfg.get("settings", {}).items()):
            ttk.Label(sbox, text=s.get("label", cat) + ":").grid(row=r, column=0, sticky="w", padx=(0, 6), pady=1)
            var = tk.StringVar(value=self.settings.get(cat, s.get("default", s["options"][0])))
            self.mode_vars[cat] = var
            for c, opt in enumerate(s["options"]):
                ttk.Radiobutton(sbox, text=MODE_TEXT.get(opt, opt.title()), value=opt, variable=var,
                                command=lambda cat=cat: self.on_mode(cat)).grid(row=r, column=c + 1, sticky="w", padx=3)
            if "random" in s["options"]:
                b = ttk.Button(sbox, text="Shuffle", style="Small.TButton", command=lambda cat=cat: self.on_shuffle(cat))
                b.grid(row=r, column=len(s["options"]) + 1, padx=(6, 0))
                self.shuffle_buttons[cat] = b
        ttk.Label(sbox, text=MODE_HINT, foreground=COLORS["single"], wraplength=420) \
            .grid(row=len(self.cfg.get("settings", {})), column=0, columnspan=6, sticky="w", pady=(6, 0))
        self._sync_shuffle()

        self._build_bottom()
        panes = ttk.PanedWindow(self, orient="vertical")
        panes.pack(fill="both", expand=True, padx=10, pady=8)

        tbox = ttk.LabelFrame(panes, text="Tours, locks and escorts", padding=(6, 4))
        bar = ttk.Frame(tbox)
        bar.pack(fill="x")
        ttk.Checkbutton(bar, text="Show only rows that need a pick", variable=self.only_choose,
                        command=self.apply_filter).pack(side="left")
        self.count_lbl = ttk.Label(bar, text="")
        self.count_lbl.pack(side="right")
        self.list = ScrollFrame(tbox)
        self.list.pack(fill="both", expand=True, pady=(4, 0))
        panes.add(tbox, weight=3)

        pbox = ttk.LabelFrame(panes, text="DAR preview", padding=(6, 4))
        self.preview = tk.Text(pbox, wrap="word", height=8, relief="flat", padx=8, pady=6,
                               font=self.f_mono)
        psb = ttk.Scrollbar(pbox, orient="vertical", command=self.preview.yview)
        self.preview.configure(yscrollcommand=psb.set, state="disabled")
        self.preview.tag_configure("blank", foreground=COLORS["missing"], font=self.f_mono_bold)
        self.preview.tag_configure("head", font=self.f_mono_bold)
        self.preview.pack(side="left", fill="both", expand=True)
        psb.pack(side="right", fill="y")
        panes.add(pbox, weight=2)

    def _build_bottom(self):
        bottom = ttk.Frame(self, padding=(10, 0, 10, 10))
        bottom.pack(side="bottom", fill="x")
        self.status_lbl = tk.Label(bottom, text="", anchor="w", font=self.f_head)
        self.status_lbl.pack(side="left")
        self.msg_lbl = tk.Label(bottom, text="", anchor="w", fg="#92400e")
        self.msg_lbl.pack(side="left", padx=12)
        self.save_btn = ttk.Button(bottom, text="Save .txt…", command=self.save_txt, state="disabled")
        self.save_btn.pack(side="right")
        self.copy_btn = ttk.Button(bottom, text="Copy DAR", command=self.copy_dar, state="disabled")
        self.copy_btn.pack(side="right", padx=6)

    # ------------------------------------------------------------ settings file
    def _load_settings(self):
        defaults = default_settings(self.cfg)
        for path in (SETTINGS_FILE, SETTINGS_FALLBACK):
            try:
                with open(path, encoding="utf-8") as fh:
                    saved = json.load(fh)
                if isinstance(saved, dict):
                    for cat, mode in saved.items():
                        if cat in defaults and mode in self.cfg["settings"][cat]["options"]:
                            defaults[cat] = mode
                    break
            except (OSError, ValueError):
                continue
        return defaults

    def _save_settings(self):
        for path in (SETTINGS_FILE, SETTINGS_FALLBACK):
            try:
                with open(path, "w", encoding="utf-8") as fh:
                    json.dump(self.settings, fh, indent=2)
                return
            except OSError:
                continue

    # ------------------------------------------------------------- open/reload
    def open_dialog(self):
        initial = os.path.dirname(self.session.report.path) if self.session else os.getcwd()
        path = filedialog.askopenfilename(title="Open shift report", initialdir=initial,
                                          filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if path:
            self.open_report(path)

    def open_report(self, path):
        if os.path.basename(path).startswith("~$"):
            messagebox.showerror(APP_TITLE, "That's Excel's temporary lock file. Open the report itself.")
            return False
        try:
            session = DarSession(path, self.config_path, settings=self.settings)
        except ReportError as e:
            messagebox.showerror(APP_TITLE, str(e))
            return False
        except Exception as e:  # anything unexpected in an odd workbook: show it, don't crash
            messagebox.showerror(APP_TITLE, f"Couldn't read this report:\n{e.__class__.__name__}: {e}")
            return False
        warn = session.load_picks()
        self.session = session
        self._save_warned = False
        self.file_lbl.configure(text=os.path.basename(path))
        self.shift_lbl.configure(text=f"{session.report.date}   {session.report.shift}")
        self.reload_btn.configure(state="normal")
        self.msg_lbl.configure(text=warn)
        self._build_duty()
        self._build_rows()
        self.refresh()
        self.list.scroll_top()
        return True

    def reload(self):
        if self.session:
            self._autosave()
            self.open_report(self.session.report.path)

    def clear_manual(self):
        if self.session and self.session.manual:
            if messagebox.askyesno(APP_TITLE, f"Clear {len(self.session.manual)} hand pick(s)?"):
                self.session.manual.clear()
                self.changed()

    # ------------------------------------------------------------ on duty panel
    def _build_duty(self):
        for w in self.duty_box.winfo_children():
            w.destroy()
        s = self.session
        staff = s.report.staff
        r = 0
        for site in staff.sites.values():
            counted = staff.counted(site)
            disp = staff.manual_only(site)
            names = ", ".join(p.name for p in counted) or "— nobody"
            ttk.Label(self.duty_box, text=site.name, style="Head.TLabel").grid(row=r, column=0, sticky="w", padx=(0, 10))
            ttk.Label(self.duty_box, text=names).grid(row=r, column=1, sticky="w")
            n = len(counted)
            tag = "pick needed" if n >= 2 else ("automatic" if n == 1 else "covered by fallback")
            tk.Label(self.duty_box, text=f"({n}) {tag}", fg=COLORS["missing"] if n >= 2 else COLORS["single"]) \
                .grid(row=r, column=2, sticky="w", padx=8)
            r += 1
            if disp:
                ttk.Label(self.duty_box, text="+ " + ", ".join(p.name for p in disp) + "  (dispatch: pick by hand only)",
                          foreground=COLORS["single"]).grid(row=r, column=1, columnspan=2, sticky="w")
                r += 1
        others = [p for p in staff.people if p.line not in staff.sites]
        if others:
            ttk.Label(self.duty_box, text="Other", style="Head.TLabel").grid(row=r, column=0, sticky="w")
            ttk.Label(self.duty_box, text=", ".join(f"{p.name} ({p.line})" for p in others)) \
                .grid(row=r, column=1, columnspan=3, sticky="w")

    # ------------------------------------------------------------ task rows
    def _render_empty(self, text):
        for w in self.list.inner.winfo_children():
            w.destroy()
        self.rows = {}
        ttk.Label(self.list.inner, text=text, padding=20, foreground=COLORS["single"]).grid(row=0, column=0, sticky="w")

    def _site_text(self, t):
        if t.status == "rule":
            return f"{t.site} staff"
        if "→" in t.site_note:
            return f"{t.site} (covering {t.site_note.split(' →')[0]})"
        if t.site_note:
            return "— nobody on duty"
        return t.site or "—"

    def _build_rows(self):
        for w in self.list.inner.winfo_children():
            w.destroy()
        self.rows = {}
        tasks = self.session.assignable
        if not tasks:
            self._render_empty("Nothing on this report needs an officer.")
            return
        inner = self.list.inner
        heads = ["Time", "Type", "Location", "Site", "Officer", "", ""]
        for c, h in enumerate(heads):
            ttk.Label(inner, text=h, style="Head.TLabel").grid(row=0, column=c, sticky="w", padx=4, pady=(0, 4))
        inner.columnconfigure(7, weight=1)  # spare space goes to the right, columns stay together
        for i, t in enumerate(tasks, start=1):
            w = {}
            w["time"] = ttk.Label(inner, text=f"{t.time[:2]}:{t.time[2:]}")
            w["type"] = ttk.Label(inner, text=t.label)
            w["loc"] = ttk.Label(inner, text=t.place or t.location or "", wraplength=300)
            w["site"] = ttk.Label(inner, text=self._site_text(t), foreground=COLORS["single"])
            cb = ttk.Combobox(inner, state="readonly", width=30, values=self.session.dropdown(t))
            cb.bind("<<ComboboxSelected>>", lambda e, k=t.key: self.on_pick(k))
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):  # wheel scrolls the list, not the value
                cb.bind(seq, self.list.on_wheel)
            w["cb"] = cb
            w["state"] = ttk.Label(inner, text="", anchor="w", width=20)
            w["reset"] = ttk.Button(inner, text="Reset", style="Small.TButton", width=6,
                                    command=lambda k=t.key: self.on_reset(k))
            for c, name in enumerate(["time", "type", "loc", "site", "cb", "state", "reset"]):
                w[name].grid(row=i, column=c, sticky="w", padx=4, pady=2)
            w["task"] = t
            self.rows[t.key] = w
        self.apply_filter()

    def apply_filter(self):
        only = self.only_choose.get()
        for key, w in self.rows.items():
            t = w["task"]
            show = (not only) or t.status == "choose" or (self.session and key in self.session.manual)
            for name in ("time", "type", "loc", "site", "cb", "state", "reset"):
                if show:
                    w[name].grid()
                else:
                    w[name].grid_remove()
            if show:
                self._refresh_row(key)

    def _refresh_row(self, key):
        w = self.rows.get(key)
        if not w:
            return
        s, t = self.session, w["task"]
        cur = s.officer(t) or ""
        values = s.dropdown(t)
        w["cb"].configure(values=values)
        w["cb"].set(cur)
        st = s.state(t)
        if st == "manual":
            text, color, style = "set by you", COLORS["manual"], "Manual.TCombobox"
        elif st == "single":
            text, color, style = "only one on site", COLORS["single"], "TCombobox"
        elif st == "rule":
            text, color, style = t.rule_name or "location rule", COLORS["rule"], "TCombobox"
        elif st == "report":
            text, color, style = "from report", COLORS["report"], "TCombobox"
        elif st.startswith("prefill:"):
            text, color, style = MODE_TEXT.get(st.split(":")[1], "setting").split(" (")[0].lower(), \
                COLORS["prefill"], "TCombobox"
        else:
            text, color, style = "⚠ choose", COLORS["missing"], "Missing.TCombobox"
        w["state"].configure(text=text, foreground=color)
        w["cb"].configure(style=style)
        if st == "manual":
            if w["time"].winfo_manager():
                w["reset"].grid()
        else:
            w["reset"].grid_remove()

    # ------------------------------------------------------------ events
    def on_pick(self, key):
        w = self.rows[key]
        value = w["cb"].get()
        if value == SEPARATOR:
            self._refresh_row(key)
            return
        self.session.set_officer(key, value)
        self.changed()

    def on_reset(self, key):
        self.session.reset(key)
        self.changed()

    def on_mode(self, cat):
        mode = self.mode_vars[cat].get()
        self.settings[cat] = mode
        self._save_settings()
        self._sync_shuffle()
        if self.session:
            self.session.set_setting(cat, mode)
            self.changed()

    def on_shuffle(self, cat):
        if self.session:
            self.session.reroll(cat)
            self.changed()

    def _sync_shuffle(self):
        for cat, b in self.shuffle_buttons.items():
            b.configure(state="normal" if self.mode_vars[cat].get() == "random" else "disabled")

    def changed(self):
        self._autosave()
        self.refresh()

    def refresh(self):
        if not self.session:
            return
        for key in self.rows:
            self._refresh_row(key)
        if self.only_choose.get():
            self.apply_filter()
        self._refresh_preview()
        missing = self.session.missing()
        total = len(self.session.assignable)
        need = sum(1 for t in self.session.assignable if t.status == "choose")
        self.count_lbl.configure(text=f"{total} rows · {need} at sites with 2+ people")
        if missing:
            self.status_lbl.configure(text=f"⚠ {len(missing)} item(s) still need an officer", fg=COLORS["missing"])
            self.copy_btn.configure(state="disabled")
            self.save_btn.configure(state="disabled")
        else:
            self.status_lbl.configure(text="✓ Ready", fg=COLORS["ok"])
            self.copy_btn.configure(state="normal")
            self.save_btn.configure(state="normal")

    def _refresh_preview(self):
        text = self.session.render()
        p = self.preview
        top = p.yview()[0]
        p.configure(state="normal")
        p.delete("1.0", "end")
        p.insert("1.0", text)
        n_head = len(self.session.report.header_lines())
        p.tag_add("head", "1.0", f"{n_head}.end")
        start = "1.0"
        while True:
            pos = p.search(BLANK, start, stopindex="end")
            if not pos:
                break
            end = f"{pos}+{len(BLANK)}c"
            p.tag_add("blank", pos, end)
            start = end
        p.configure(state="disabled")
        p.yview_moveto(top)

    def _autosave(self):
        if not self.session:
            return
        try:
            self.session.save_picks()
        except OSError:
            if not self._save_warned:
                self._save_warned = True
                self.msg_lbl.configure(text="Picks can't be saved next to this report (folder is read-only).")

    # ------------------------------------------------------------ output
    def dar_text(self):
        return self.session.render() if self.session else ""

    def copy_dar(self):
        if not self.session or self.session.missing():
            return
        self.clipboard_clear()
        self.clipboard_append(self.dar_text())
        self.update()
        self.msg_lbl.configure(text="DAR copied to the clipboard.")

    def save_txt(self):
        if not self.session or self.session.missing():
            return
        rp = self.session.report.path
        stem = os.path.splitext(os.path.basename(rp))[0]
        path = filedialog.asksaveasfilename(title="Save DAR", initialdir=os.path.dirname(rp),
                                            initialfile=f"{stem}_DAR.txt", defaultextension=".txt",
                                            filetypes=[("Text file", "*.txt"), ("All files", "*.*")])
        if path:
            self.write_txt(path)

    def write_txt(self, path):
        try:
            with open(path, "w", encoding="utf-8", newline="") as fh:
                fh.write(self.dar_text().replace("\n", os.linesep) + os.linesep)
            self.msg_lbl.configure(text=f"Saved {os.path.basename(path)}")
            return True
        except OSError as e:
            messagebox.showerror(APP_TITLE, f"Couldn't save the file:\n{e}")
            return False

    def report_callback_exception(self, exc, val, tb):
        """Unexpected error inside the window: log it and show it instead of failing silently
        (a packaged .exe has no console)."""
        import traceback
        text = "".join(traceback.format_exception(exc, val, tb))
        log = os.path.join(HERE, "dar_app_error.log")
        try:
            with open(log, "a", encoding="utf-8") as fh:
                fh.write(text + "\n")
        except OSError:
            log = "(couldn't write log)"
        messagebox.showerror(APP_TITLE, f"Something went wrong:\n{val}\n\nDetails saved to:\n{log}")

    def on_close(self):
        self._autosave()
        self._save_settings()
        self.destroy()


def _windows_dpi():
    """Sharp text on high-DPI Windows screens (no-op elsewhere)."""
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass


def main():
    _windows_dpi()
    path = sys.argv[1] if len(sys.argv) > 1 else None
    app = DarApp(path)
    app.mainloop()


if __name__ == "__main__":
    main()
