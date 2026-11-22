"""Drives the real Tk window (needs a display; run under xvfb-run on a server)."""
import os
import shutil
import stat
import subprocess
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [os.path.dirname(HERE), HERE]

tk = pytest.importorskip("tkinter")
import fixtures as F  # noqa: E402
import dar_app  # noqa: E402
from dar_session import SEPARATOR  # noqa: E402
from shift_report_parser import BLANK  # noqa: E402

SHOTS = os.path.join(HERE, "out", "screens")
LT, PSO, DISP = "Lt. Fahid Afzal", "PSO. Tim Green", "Disp. Mostafa Kamal"


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.setattr(dar_app, "SETTINGS_FILE", str(tmp_path / "settings.json"))
    monkeypatch.setattr(dar_app, "SETTINGS_FALLBACK", str(tmp_path / "settings2.json"))
    errors = []
    monkeypatch.setattr(dar_app.messagebox, "showerror", lambda *a, **k: errors.append(a))
    monkeypatch.setattr(dar_app.messagebox, "askyesno", lambda *a, **k: True)
    a = dar_app.DarApp()
    a.errors = errors
    a.update()
    yield a
    try:
        a.destroy()
    except tk.TclError:
        pass


def pump(a):
    for _ in range(3):
        a.update_idletasks()
        a.update()


def pick(a, key, value):
    cb = a.rows[key]["cb"]
    cb.set(value)
    cb.event_generate("<<ComboboxSelected>>")
    pump(a)


def visible_rows(a):
    return [k for k, w in a.rows.items() if w["cb"].winfo_manager()]


def shot(a, name):
    os.makedirs(SHOTS, exist_ok=True)
    pump(a)
    a.lift()
    pump(a)
    path = os.path.join(SHOTS, name)
    if shutil.which("import"):
        subprocess.run(["import", "-window", str(a.winfo_id()) if False else "root", path], check=False)
    return path


def test_open_and_required_picks(app, tmp_path):
    p = F.base_copy("gui_base.xlsx")
    assert app.open_report(p)
    pump(app)
    s = app.session
    assert len(app.rows) == 11
    assert app.status_lbl.cget("text").startswith("⚠ 4")
    assert str(app.copy_btn.cget("state")) == "disabled" and str(app.save_btn.cget("state")) == "disabled"
    preview = app.preview.get("1.0", "end")
    assert preview.count(BLANK) == 4 and preview.startswith("11/21/2026 0000-0800")
    doble = [t for t in s.assignable if t.site == "Doble"]
    for t in doble:
        w = app.rows[t.key]
        assert w["cb"].get() == "" and "choose" in w["state"].cget("text")
        assert w["cb"].cget("style") == "Missing.TCombobox"
    south = next(t for t in s.assignable if t.site == "South")
    assert app.rows[south.key]["cb"].get() == "PSO. Amare Dugoma"
    assert app.rows[south.key]["state"].cget("text") == "only one on site"
    assert not app.rows[south.key]["reset"].winfo_manager()
    shot(app, "1_opened.png")

    # separator does nothing
    pick(app, doble[0].key, SEPARATOR)
    assert app.rows[doble[0].key]["cb"].get() == "" and doble[0].key not in s.manual

    # pick all four, including the dispatcher by hand
    pick(app, doble[0].key, PSO)
    pick(app, doble[1].key, DISP)
    pick(app, doble[2].key, LT)
    assert app.status_lbl.cget("text").startswith("⚠ 1")
    pick(app, doble[3].key, LT)
    assert app.status_lbl.cget("text") == "✓ Ready"
    assert str(app.copy_btn.cget("state")) == "normal"
    assert app.rows[doble[1].key]["state"].cget("text") == "set by you"
    assert app.rows[doble[1].key]["reset"].winfo_manager() == "grid"
    assert BLANK not in app.preview.get("1.0", "end") and DISP in app.preview.get("1.0", "end")

    # hand pick on a single row + reset
    pick(app, south.key, DISP)
    assert app.rows[south.key]["state"].cget("text") == "set by you"
    app.rows[south.key]["reset"].invoke()
    pump(app)
    assert app.rows[south.key]["cb"].get() == "PSO. Amare Dugoma"
    assert not app.rows[south.key]["reset"].winfo_manager()

    # copy + save
    app.copy_btn.invoke()
    pump(app)
    assert app.clipboard_get() == s.render()
    out = tmp_path / "dar.txt"
    assert app.write_txt(str(out))
    assert out.read_text().strip().splitlines() == s.render().splitlines()
    shot(app, "2_all_picked.png")

    # picks were autosaved: reload restores them
    assert os.path.exists(s.picks_path)
    app.reload()
    pump(app)
    assert app.status_lbl.cget("text") == "✓ Ready"
    assert app.session.officer(app.session.task(doble[1].key)) == DISP
    os.remove(app.session.picks_path)


def test_settings_filter_shuffle(app):
    p = F.make("gui_grown.xlsx", start=F.grown)
    assert app.open_report(p)
    pump(app)
    s = app.session
    n_choose = sum(1 for t in s.assignable if t.status == "choose")
    assert n_choose > 0
    assert str(app.shuffle_buttons["pda"].cget("state")) == "disabled"

    app.only_choose.set(True)
    app.apply_filter()
    pump(app)
    assert len(visible_rows(app)) == n_choose
    shot(app, "3_filter.png")
    app.only_choose.set(False)
    app.apply_filter()
    pump(app)
    assert len(visible_rows(app)) == len(app.rows)

    for cat, mode in (("pda", "rotate"), ("visual", "random"), ("lock", "auto")):
        app.mode_vars[cat].set(mode)
        app.on_mode(cat)
    pump(app)
    assert app.status_lbl.cget("text") == "✓ Ready"
    assert str(app.shuffle_buttons["visual"].cget("state")) == "normal"
    states = {app.rows[t.key]["state"].cget("text") for t in s.assignable if t.status == "choose"}
    assert states <= {"rotate", "random", "auto"}
    for _ in range(10):
        app.shuffle_buttons["visual"].invoke()
    pump(app)
    assert app.status_lbl.cget("text") == "✓ Ready"
    shot(app, "4_settings.png")

    app.mode_vars["pda"].set("ask")
    app.on_mode("pda")
    pump(app)
    assert app.status_lbl.cget("text").startswith("⚠")
    assert "34R Prentiss rule" in {w["state"].cget("text") for w in app.rows.values()}
    assert any("covering Lunder" in w["site"].cget("text") for w in app.rows.values())

    app.on_close()
    with open(dar_app.SETTINGS_FILE) as fh:
        assert '"pda": "ask"' in fh.read()
    os.remove(s.picks_path)


def test_wheel_on_dropdown_does_not_change_value(app):
    assert app.open_report(F.base_copy("gui_wheel.xlsx"))
    pump(app)
    key = next(t.key for t in app.session.assignable if t.status == "single")
    cb = app.rows[key]["cb"]
    before = cb.get()
    for seq in ("<Button-4>", "<Button-5>"):
        cb.event_generate(seq)
    cb.event_generate("<MouseWheel>", delta=-120)
    pump(app)
    assert cb.get() == before and key not in app.session.manual
    os.path.exists(app.session.picks_path) and os.remove(app.session.picks_path)


def test_bad_files_show_error_and_keep_current(app, tmp_path):
    assert app.open_report(F.base_copy("gui_keep.xlsx"))
    current = app.session
    bad = tmp_path / "bad.xlsx"
    bad.write_text("nope")
    assert not app.open_report(str(bad))
    assert not app.open_report(str(tmp_path / "missing.xlsx"))
    lock = tmp_path / "~$report.xlsx"
    lock.write_text("x")
    assert not app.open_report(str(lock))
    assert len(app.errors) == 3 and app.session is current
    os.path.exists(current.picks_path) and os.remove(current.picks_path)


def test_report_with_nothing_to_assign(app):
    def edit(ws):
        for title in ("PDA Tours", "Visual Tours"):
            top = F.find_row(ws, title)
            for rr in range(top + 2, top + 6):
                for c in range(2, 9):
                    cell = ws.cell(rr, c)
                    if type(cell).__name__ != "MergedCell" and (c > 2 or title == "Visual Tours"):
                        cell.value = None
    assert app.open_report(F.make("gui_empty.xlsx", edit))
    pump(app)
    assert app.rows == {} and app.status_lbl.cget("text") == "✓ Ready"
    assert str(app.copy_btn.cget("state")) == "normal"


def test_readonly_folder_warns_once(app, tmp_path):
    if os.geteuid() == 0:
        pytest.skip("root ignores folder permissions")
    ro = tmp_path / "ro"
    ro.mkdir()
    shutil.copy(F.BASE, ro / "r.xlsx")
    os.chmod(ro, stat.S_IRUSR | stat.S_IXUSR)
    try:
        assert app.open_report(str(ro / "r.xlsx"))
        t = app.session.missing()[0]
        pick(app, t.key, LT)
        assert "read-only" in app.msg_lbl.cget("text")
    finally:
        os.chmod(ro, stat.S_IRWXU)


def test_command_line_path_opens(tmp_path, monkeypatch):
    monkeypatch.setattr(dar_app, "SETTINGS_FILE", str(tmp_path / "s.json"))
    a = dar_app.DarApp(F.base_copy("gui_cli.xlsx"))
    try:
        a.after(200, a.quit)
        a.mainloop()
        assert a.session is not None and len(a.rows) == 11
    finally:
        a.destroy()


def test_save_failure_warns_once(app, monkeypatch):
    assert app.open_report(F.base_copy("gui_fail.xlsx"))
    calls = []

    def boom(*a, **k):
        calls.append(1)
        raise PermissionError("read-only")
    monkeypatch.setattr(app.session, "save_picks", boom)
    for t in list(app.session.missing())[:2]:
        pick(app, t.key, LT)
    assert len(calls) == 2 and "read-only" in app.msg_lbl.cget("text")
    assert app.status_lbl.cget("text").startswith("⚠ 2")
