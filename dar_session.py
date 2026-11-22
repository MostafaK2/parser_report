"""
DarSession - everything the app does, without any window code (so it can be tested).

Who ends up on a DAR line (highest first):
  1. manual  - picked by hand in the app (any row, any on-duty person)
  2. auto    - report / rule / single rows: the automatic officer
  3. prefill - choose rows only: filled in by a setting (rotate / random / auto)
  4. blank   - choose rows with nothing above: shown as ________ and counted as missing

Picks are saved next to the report as <report name>.dar.json and restored on reopen.
"""
from __future__ import annotations

import json
import os
import random
from typing import Dict, List, Optional

from shift_report_parser import DEFAULT_CONFIG, Report, Task, default_settings, extract, load_config, render

PICKS_VERSION = 1
SEPARATOR = "──────────"


def picks_path_for(report_path: str) -> str:
    stem, _ = os.path.splitext(report_path)
    return stem + ".dar.json"


class DarSession:
    def __init__(self, report_path: str, config_path: str = DEFAULT_CONFIG, settings: Optional[Dict] = None,
                 seed: Optional[int] = None):
        self.config = load_config(config_path)
        self.report: Report = extract(report_path, self.config)
        self.settings = default_settings(self.config)
        if settings:
            for cat, mode in settings.items():
                self.set_setting(cat, mode, recompute=False)
        self.manual: Dict[str, str] = {}
        self.prefill: Dict[str, str] = {}
        self.rng = random.Random(seed)
        self.recompute_prefill()

    # ------------------------------------------------------------------ tasks
    @property
    def tasks(self) -> List[Task]:
        return self.report.tasks

    @property
    def assignable(self) -> List[Task]:
        return [t for t in self.report.tasks if t.needs_officer]

    def task(self, key) -> Optional[Task]:
        return next((t for t in self.report.tasks if t.key == key), None)

    # --------------------------------------------------------------- settings
    def options_for(self, category) -> List[str]:
        return self.config.get("settings", {}).get(category, {}).get("options", ["ask"])

    def set_setting(self, category, mode, recompute=True):
        if category not in self.config.get("settings", {}) or mode not in self.options_for(category):
            return False
        self.settings[category] = mode
        if recompute:
            self.recompute_prefill(category)
        return True

    def recompute_prefill(self, category=None):
        """Refill choose rows from the settings. Hand picks are never touched."""
        for t in self.assignable:
            if t.status != "choose" or (category and t.category != category):
                continue
            self.prefill.pop(t.key, None)
            mode = self.settings.get(t.category, "ask")
            if mode in ("rotate", "auto") and t.rotate_pick:
                self.prefill[t.key] = t.rotate_pick
            elif mode == "random" and t.candidates:
                self.prefill[t.key] = self.rng.choice(t.candidates)

    def reroll(self, category=None):
        self.recompute_prefill(category)

    # ------------------------------------------------------------------ picks
    def officer(self, t: Task) -> Optional[str]:
        if not t.needs_officer:
            return None
        if t.key in self.manual:
            return self.manual[t.key]
        if t.status != "choose":
            return t.auto
        return self.prefill.get(t.key)

    def state(self, t: Task) -> str:
        """manual / report / rule / single / prefill:<mode> / missing / none"""
        if not t.needs_officer:
            return "none"
        if t.key in self.manual:
            return "manual"
        if t.status != "choose":
            return t.status if t.auto else "missing"
        if t.key in self.prefill:
            return "prefill:" + self.settings.get(t.category, "")
        return "missing"

    def set_officer(self, key, name: Optional[str]):
        """Pick by hand. Picking the automatic person (or nothing) clears the hand pick."""
        t = self.task(key)
        if t is None or not t.needs_officer:
            return
        name = (name or "").strip()
        if not name or name == SEPARATOR:
            self.manual.pop(key, None)
            return
        default = t.auto if t.status != "choose" else self.prefill.get(key)
        if t.status != "choose" and name == default:
            self.manual.pop(key, None)
        else:
            self.manual[key] = name

    def reset(self, key):
        self.manual.pop(key, None)

    def dropdown(self, t: Task) -> List[str]:
        """Dropdown values: people who count, then the site dispatcher, then everyone else on duty."""
        top = list(t.candidates)
        mid = [n for n in t.extras if n not in top]
        rest = [n for n in self.report.all_staff if n not in top and n not in mid]
        values = list(top)
        for group in (mid, rest):
            if group:
                if values:
                    values.append(SEPARATOR)
                values.extend(group)
        cur = self.officer(t)
        if cur and cur not in values:  # e.g. a name written on the report that isn't on duty
            values = [cur] + ([SEPARATOR] if values else []) + values
        return values

    def missing(self) -> List[Task]:
        return [t for t in self.assignable if not self.officer(t)]

    def render(self) -> str:
        return render(self.report, self.officer)

    def render_lines(self):
        """[(text, task or None)] - for the preview; header lines have task None."""
        out = [(line, None) for line in self.report.header_lines()] + [("", None)]
        text = self.render().split("\n")
        for t, line in zip(self.report.tasks, text[len(out):]):
            out.append((line, t))
        return out

    # ------------------------------------------------------------ save / load
    @property
    def picks_path(self) -> str:
        return picks_path_for(self.report.path)

    def save_picks(self, path=None):
        path = path or self.picks_path
        data = {
            "version": PICKS_VERSION,
            "report": os.path.basename(self.report.path),
            "settings": self.settings,
            "manual": self.manual,
            "prefill": {k: v for k, v in self.prefill.items() if self.settings.get(self.task(k).category) == "random"},
        }
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2, ensure_ascii=False)
        os.replace(tmp, path)

    def load_picks(self, path=None) -> str:
        """Restore saved picks. -> '' if fine/nothing saved, else a short warning."""
        path = path or self.picks_path
        if not os.path.exists(path):
            return ""
        try:
            with open(path, encoding="utf-8") as fh:
                data = json.load(fh)
            if not isinstance(data, dict) or data.get("version") != PICKS_VERSION:
                return "Saved picks were in an unknown format and were ignored."
        except (OSError, ValueError):
            return "Saved picks couldn't be read and were ignored."
        keys = {t.key: t for t in self.assignable}
        dropped = 0
        for k, v in (data.get("manual") or {}).items():
            if k in keys and isinstance(v, str) and v.strip():
                self.manual[k] = v
            else:
                dropped += 1
        saved_settings = data.get("settings") or {}
        for k, v in (data.get("prefill") or {}).items():
            t = keys.get(k)
            if t and t.status == "choose" and v in t.candidates and \
                    self.settings.get(t.category) == "random" == saved_settings.get(t.category):
                self.prefill[k] = v
        return f"{dropped} saved pick(s) no longer match the report and were skipped." if dropped else ""
