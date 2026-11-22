#!/usr/bin/env python3
"""
Shift Report Parser - engine.

All report-specific data (campuses, buildings, sites, section names, column names,
who does tours, DAR wording) lives in report_config.json. This file only knows how
to read the sheet layout.

Command line:
    python shift_report_parser.py <report.xlsx>                  # print the DAR
    python shift_report_parser.py <report.xlsx> --sections       # show the section map
    python shift_report_parser.py <report.xlsx> --use-saved      # apply picks saved by the app
    python shift_report_parser.py <report.xlsx> --config other.json

Python use:
    report = extract("report.xlsx")              # -> Report with .tasks, .staff, .sites ...
    text   = render(report, chooser)             # chooser(task) -> officer name or None

Pipeline
--------
1. find_sections()  finds every section from the sheet's merged title rows.
2. Staff            reads Personnel on Duty into people, staff lines and sites.
3. handlers         turn each DAR section into raw items (time, location, fields).
4. _build_tasks()   gives every item a site, the people who can be picked, and a status:
                      none    - no officer in the line (incident reports)
                      report  - officer written on the report (a "Done By" column / "01:00 Lt. X")
                      rule    - a location rule picked the officer (e.g. 34R Prentiss)
                      single  - only one person counts at the site, so it's them
                      choose  - two or more people count, someone has to pick
5. render()         sorts by time since shift start and fills in the templates.
"""
from __future__ import annotations

import json
import os
import re
import sys
from collections import OrderedDict, defaultdict
from dataclasses import dataclass, field
from datetime import datetime, time
from typing import Callable, Dict, List, Optional

import openpyxl

def app_dir() -> str:
    """Folder the user sees: next to the .exe when packaged, else next to this file."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _default_config() -> str:
    """report_config.json next to the app wins (so it can be edited); else the bundled copy."""
    for d in (app_dir(), getattr(sys, "_MEIPASS", None)):
        if d and os.path.isfile(os.path.join(d, "report_config.json")):
            return os.path.join(d, "report_config.json")
    return os.path.join(app_dir(), "report_config.json")


HERE = app_dir()
DEFAULT_CONFIG = _default_config()
BLANK = "________"


class ReportError(Exception):
    """The file couldn't be read as a shift report."""


# ==========================================================================
# Value helpers
# ==========================================================================
def clean(val):
    """Cell value -> tidy string/number, or None for empty / N/A."""
    if val is None:
        return None
    if isinstance(val, str):
        s = " ".join(val.split())
        return None if s.lower() in ("", "n/a", "na", "-", "none") else s
    return val


def norm(text) -> str:
    return " ".join(re.sub(r"[^a-z0-9/#.' ]", " ", str(text).lower()).split())


def has_words(text, keywords) -> bool:
    """True if any keyword appears in text at the start of a word ('desc' matches 'Description')."""
    if text is None:
        return False
    t = norm(text)
    return any(re.search(r"(?<![a-z0-9])" + re.escape(k.lower()), t) for k in keywords)


_TIME_RE = re.compile(r"^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?(?:\s*[-–(]?\s*(.*?)\)?)?$")
_HHMM_RE = re.compile(r"^(\d{1,2})(\d{2})$")


def parse_time(val):
    """-> (HHMM, trailing text or None), or None if val isn't a time.
    Accepts time/datetime, Excel day fractions, '4:00', '04:00 PM', '0400', '01:00 Lt. Afzal'."""
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (time, datetime)):
        return val.strftime("%H%M"), None
    if isinstance(val, (int, float)):
        if 0 <= val < 1:
            mins = round(val * 1440) % 1440
            return f"{mins // 60:02d}{mins % 60:02d}", None
        return None
    s = str(val).strip()
    m = _TIME_RE.match(s)
    if m:
        h, mi, ampm, rest = int(m.group(1)), int(m.group(2)), m.group(3), clean(m.group(4))
    else:
        m = _HHMM_RE.match(s)
        if not m:
            return None
        h, mi, ampm, rest = int(m.group(1)), int(m.group(2)), None, None
    if ampm:
        if h < 1 or h > 12:
            return None
        h = h % 12 + (12 if ampm.lower() == "pm" else 0)
    if h > 23 or mi > 59:
        return None
    return f"{h:02d}{mi:02d}", rest


def render_template(segments, fields) -> str:
    """Each template segment is kept only if all of its {fields} have values."""
    out = []
    for seg in segments:
        names = re.findall(r"{(\w+)}", seg)
        if all(fields.get(n) not in (None, "") for n in names):
            out.append(seg.format(**{n: fields[n] for n in names}))
    return "".join(out)


def minutes(hhmm: str) -> int:
    return int(hhmm[:2]) * 60 + int(hhmm[2:])


# ==========================================================================
# Sheet layout: sections, tables, label/value blocks
# ==========================================================================
class Sheet:
    def __init__(self, ws):
        self.ws = ws
        self.merge_at = {(r.min_row, r.min_col): r for r in ws.merged_cells.ranges}
        self.min_row, self.max_row, self.max_col = ws.min_row, ws.max_row, ws.max_column

    def value(self, r, c):
        return self.ws.cell(r, c).value

    def span(self, r, c):
        rng = self.merge_at.get((r, c))
        return (rng.min_col, rng.max_col, rng.max_row - rng.min_row + 1) if rng else (c, c, 1)

    def row_cells(self, r, c1, c2):
        return [(c, v) for c in range(c1, c2 + 1)
                if (v := self.value(r, c)) is not None and str(v).strip() != ""]


@dataclass
class Section:
    name: str
    title_row: int
    start_row: int
    end_row: int
    sheet: Sheet = field(repr=False)
    left: int = field(repr=False)
    right: int = field(repr=False)

    @property
    def rows(self):
        return range(self.start_row, self.end_row + 1)

    def table(self, full_width=False):
        """First row = column headers. full_width=True also reads cells right of the form."""
        sh = self.sheet
        if self.start_row > self.end_row:
            return [], []
        right = max(sh.max_col, self.right) if full_width else self.right
        headers = [(c, str(v).strip()) for c, v in sh.row_cells(self.start_row, self.left, right)]
        records = []
        for r in range(self.start_row + 1, self.end_row + 1):
            cells = sh.row_cells(r, self.left, right)
            if any(clean(v) is not None for _, v in cells):
                records.append({"row": r, "cells": cells, "by_col": dict(cells)})
        return headers, records

    def key_value_blocks(self):
        """Label/value layouts. -> OrderedDict {block title: [(label, value), ...]}"""
        sh = self.sheet
        groups = None
        for r in self.rows:
            cells = sh.row_cells(r, self.left, self.right)
            spans = [sh.span(r, c) for c, _ in cells]
            if cells and all(n == 1 and c2 > c1 for c1, c2, n in spans):
                groups = [(c1, c2) for c1, c2, _ in spans]
                break
        blocks = OrderedDict()
        if not groups:
            return blocks
        current = {g: "General" for g in groups}
        for r in self.rows:
            for g1, g2 in groups:
                v = sh.value(r, g1)
                if v is None:
                    continue
                c1, c2, _ = sh.span(r, g1)
                if c2 >= g2:
                    if (c1, c2) == (g1, g2):
                        current[(g1, g2)] = str(v).strip()
                        blocks.setdefault(current[(g1, g2)], [])
                    continue
                value = next((val for _, val in sh.row_cells(r, c2 + 1, g2)), None)
                blocks.setdefault(current[(g1, g2)], []).append((str(v).strip(), value))
        return blocks


def find_sections(ws):
    """Section titles = full-width single-row merged cells. -> OrderedDict {name: Section}"""
    sh = Sheet(ws)
    widths = defaultdict(int)
    for (r, c), rng in sh.merge_at.items():
        if rng.min_row == rng.max_row and sh.value(r, c) is not None:
            widths[(rng.min_col, rng.max_col)] += 1
    if not widths:
        raise ReportError("No section titles found (expected merged title rows like 'PDA Tours').")
    left, right = max(widths, key=lambda k: (k[1] - k[0], widths[k]))

    titles = []
    for r in range(sh.min_row, sh.max_row + 1):
        rng = sh.merge_at.get((r, left))
        if not rng or rng.min_row != rng.max_row or (rng.min_col, rng.max_col) != (left, right):
            continue
        text = clean(sh.value(r, left))
        if isinstance(text, str) and not (titles and titles[-1][0] == r - 1):
            titles.append((r, text))

    def last_used(r1, r2):
        return next((r for r in range(r2, r1 - 1, -1) if sh.row_cells(r, left, right)), r1 - 1)

    sections = OrderedDict()
    first = titles[0][0] if titles else sh.max_row + 1
    if first > sh.min_row:
        sections["Header"] = Section("Header", 0, sh.min_row, last_used(sh.min_row, first - 1), sh, left, right)
    for i, (r, name) in enumerate(titles):
        nxt = titles[i + 1][0] if i + 1 < len(titles) else sh.max_row + 1
        key, n = name, 2
        while key in sections:
            key, n = f"{name} ({n})", n + 1
        sections[key] = Section(name, r, r + 1, last_used(r + 1, nxt - 1), sh, left, right)
    return sections


def matches(name, keywords) -> bool:
    return all(k.lower() in norm(name) for k in keywords)


def column_for(headers, keywords, exclude=()):
    for c, label in headers:
        if c not in exclude and has_words(label, keywords):
            return c
    return None


# ==========================================================================
# Campuses, staff, sites
# ==========================================================================
class Campuses:
    def __init__(self, cfg):
        self.cfg = cfg["campuses"]

    def campus_name(self, text):
        """Campus if text IS a campus name ('Doble', 'Porter Sq', 'Uhall'), else None."""
        if not text:
            return None
        t = norm(text)
        return next((c for c, d in self.cfg.items() if t in (n.lower() for n in d.get("names", []))), None)

    def resolve(self, text):
        if not text:
            return None
        if (c := self.campus_name(text)):
            return c
        for key in ("names", "nicknames", "keywords"):
            for campus, d in self.cfg.items():
                if has_words(text, d.get(key, [])):
                    return campus
        return None

    def place(self, location):
        c = self.campus_name(location)
        return f"{c} Campus" if c else location


@dataclass
class Person:
    name: str
    label: str            # post on the form, e.g. "Uhall Officer"
    role: Optional[str]   # dispatcher / supervisor / officer / None
    block: str            # block title on the form, e.g. "Porter Square Campus"
    line: Optional[str]   # staff line in the DAR, e.g. "Uhall"


@dataclass
class Site:
    name: str
    campus: Optional[str]
    match: List[str]
    fallback: Optional[str]
    people: List[Person]


class Staff:
    def __init__(self, blocks, cfg):
        pcfg = cfg["personnel"]
        self.pcfg = pcfg
        self.roles = pcfg["roles"]
        self.rules = pcfg.get("staff_lines", [])
        self.people: List[Person] = []
        for block, pairs in blocks.items():
            for label, value in pairs:
                name = clean(value)
                if name is None:
                    continue
                line = self._line_for(label, block)
                if line is None and pcfg.get("include_other_blocks", True):
                    line = block
                self.people.append(Person(str(name), label, self.role_of(label), block, line))

        self.sites: "OrderedDict[str, Site]" = OrderedDict()
        for rule in self.rules:
            s = rule.get("site")
            if s is not None:
                self.sites[rule["label"]] = Site(rule["label"], s.get("campus"), s.get("match", []),
                                                 s.get("fallback"), [p for p in self.people if p.line == rule["label"]])

    def role_of(self, label):
        return next((r for r, kws in self.roles.items() if has_words(label, kws)), None)

    def _line_for(self, label, block):
        def fits(rule):
            if rule.get("labels") and not has_words(label, rule["labels"]):
                return False
            if rule.get("blocks") and not has_words(block, rule["blocks"]):
                return False
            return bool(rule.get("labels") or rule.get("blocks"))
        ordered = [r for r in self.rules if r.get("labels")] + [r for r in self.rules if not r.get("labels")]
        rule = next((r for r in ordered if fits(r)), None)
        return rule["label"] if rule else None

    # ---- DAR staff lines ----
    def staff_lines(self):
        order = self.pcfg["staff_line_order"]
        lines = OrderedDict((r["label"], []) for r in self.rules)
        for p in self.people:
            if p.line is not None:
                lines.setdefault(p.line, []).append(p)
        out = []
        for label, people in lines.items():
            if people:
                people = sorted(people, key=lambda p: order.index(p.role) if p.role in order else len(order))
                out.append(f"{label}: {', '.join(p.name for p in people)}")
        return out

    # ---- sites ----
    def counted(self, site: Optional[Site], roles=None) -> List[Person]:
        """People at the site who count toward the one-person rule, in role order."""
        if site is None:
            return []
        roles = roles or self.pcfg.get("site_roles", ["officer", "supervisor"])
        return sorted([p for p in site.people if p.role in roles], key=lambda p: roles.index(p.role))

    def manual_only(self, site: Optional[Site]) -> List[Person]:
        roles = self.pcfg.get("manual_only_roles", ["dispatcher"])
        return [p for p in site.people if p.role in roles] if site else []

    def site_for(self, location, campuses: Campuses) -> Optional[Site]:
        """The site a location belongs to (before falling back)."""
        if location:
            for s in self.sites.values():
                if s.match and has_words(location, s.match):
                    return s
            campus = campuses.resolve(location)
            for s in self.sites.values():
                if campus and s.campus == campus:
                    return s
        return self.sites.get(self.pcfg.get("fallback_site"))

    def covering(self, site: Optional[Site]):
        """Follow fallbacks until a site has someone who counts. -> (site, note)"""
        seen, start = set(), site
        while site is not None and site.name not in seen:
            if self.counted(site):
                note = "" if site is start else f"{start.name} → {site.name} (nobody at {start.name})"
                return site, note
            seen.add(site.name)
            nxt = site.fallback or self.pcfg.get("fallback_site")
            site = self.sites.get(nxt) if nxt else None
        return start, (f"nobody on duty at {start.name}" if start else "no site")

    def match_name(self, text, site: Optional[Site]):
        """A written officer ('Lt', 'PSO', 'Afzal', 'Lt. Fahid Afzal') -> a person's name if it can be matched."""
        t = norm(text)
        if not t:
            return None
        for p in self.people:  # full or partial name match
            pn = norm(p.name)
            if t == pn or (len(t) > 3 and t in pn):
                return p.name
        role = self.role_of(text)
        if role and len(t.split()) == 1:
            pool = [p for p in (site.people if site else []) if p.role == role] or \
                   [p for p in self.people if p.role == role]
            if pool:
                return pool[0].name
        return text.strip()


# ==========================================================================
# Tasks
# ==========================================================================
@dataclass
class Task:
    key: str                 # stable id, used to save picks
    section: str             # section title on the sheet
    label: str               # short type, e.g. "PDA", "Visual", "Special lock"
    category: Optional[str]  # settings group: "pda" / "visual" / "lock"
    time: str                # HHMM
    row: int
    location: Optional[str]
    place: Optional[str]
    fields: Dict[str, str]
    template: List[str]
    needs_officer: bool
    status: str = "none"     # none / report / rule / single / choose
    site: Optional[str] = None
    site_note: str = ""
    rule_name: str = ""
    candidates: List[str] = field(default_factory=list)   # people who count (top of dropdown)
    extras: List[str] = field(default_factory=list)       # site dispatcher(s): manual only
    auto: Optional[str] = None       # officer for report/rule/single
    rotate_pick: Optional[str] = None  # officer the "rotate"/"auto" setting would pick
    sort_key: int = 0
    order: int = 0           # section order on the sheet


@dataclass
class Report:
    path: str
    config: dict
    date: str
    shift: str
    shift_start: int
    staff: Staff
    campuses: Campuses
    tasks: List[Task]
    sections: "OrderedDict[str, Section]"

    @property
    def all_staff(self) -> List[str]:
        return [p.name for p in self.staff.people]

    def header_lines(self) -> List[str]:
        return [f"{self.date} {self.shift}".strip()] + self.staff.staff_lines()


def since_start(hhmm, shift_start):
    return (minutes(hhmm) - shift_start) % 1440


# ---- section handlers -> raw items: dict(time, location, fields, done_by, row) ----
def handle_rows(sec, scfg, campuses):
    headers, recs = sec.table()
    cols = {k: column_for(headers, kws) for k, kws in scfg.get("columns", {}).items()}
    items = []
    for rec in recs:
        f = {k: (clean(rec["by_col"].get(c)) if c else None) for k, c in cols.items()}
        t = parse_time(f.get("time"))
        f["time"] = t[0] if t else None
        for k in scfg.get("lowercase", []):
            if isinstance(f.get(k), str):
                f[k] = f[k].lower()
        for k in scfg.get("nonzero", []):
            nums = re.findall(r"\d+", str(f.get(k) or ""))
            if not nums or int(nums[0]) == 0:
                f[k] = None
        f = {k: (str(v) if v is not None else None) for k, v in f.items()}
        if not all(f.get(k) for k in scfg.get("required", [])):
            continue
        done_by = f.pop("done_by", None)
        items.append({"time": f["time"], "location": f.get("location"), "fields": f,
                      "done_by": done_by, "row": rec["row"]})
    return items


def handle_tours_by_row(sec, scfg, campuses):
    headers, recs = sec.table(full_width=True)
    loc_col = column_for(headers, scfg["columns"]["location"])
    by_cols = {c for c, lbl in headers if c != loc_col and has_words(lbl, scfg["columns"].get("done_by", []))}
    items = []
    for rec in recs:
        location = clean(rec["by_col"].get(loc_col)) if loc_col else None
        if not isinstance(location, str):
            continue
        last = None
        for c, v in rec["cells"]:
            if c == loc_col:
                continue
            if c in by_cols:
                if last is not None and clean(v):
                    items[last]["done_by"] = str(clean(v))
                continue
            for tok in (re.split(r"[/,;\n]", v) if isinstance(v, str) else [v]):
                pt = parse_time(tok.strip() if isinstance(tok, str) else tok)
                if pt:
                    items.append({"time": pt[0], "location": location, "fields": {"time": pt[0]},
                                  "done_by": pt[1], "row": rec["row"]})
                    last = len(items) - 1
    return items


def handle_tours_in_cells(sec, scfg, campuses):
    headers, recs = sec.table(full_width=True)
    items = []
    for rec in recs:
        for _, v in rec["cells"]:
            if not isinstance(v, str):
                continue
            times, place_parts = [], []
            for tok in v.split("/"):
                pt = parse_time(tok.strip())
                (times.append(pt) if pt else place_parts.append(tok.strip()))
            location = clean("/".join(p for p in place_parts if p))
            if location and times:
                for t, name in times:
                    items.append({"time": t, "location": location, "fields": {"time": t},
                                  "done_by": name, "row": rec["row"]})
    return items


HANDLERS = {"rows": handle_rows, "tours_by_row": handle_tours_by_row, "tours_in_cells": handle_tours_in_cells}


def section_config(name, cfg):
    return next((s for s in cfg["sections"] if matches(name, s["match"])), None)


def _build_tasks(report_sections, cfg, staff: Staff, campuses: Campuses, shift_start: int) -> List[Task]:
    tasks: List[Task] = []
    seen_keys = defaultdict(int)
    for order, (name, sec) in enumerate(report_sections.items()):
        scfg = section_config(name, cfg)
        if not scfg:
            continue
        needs = any("{officer}" in seg for seg in scfg["template"])
        rotation = scfg.get("rotation", ["officer", "supervisor"])
        for it in HANDLERS[scfg["type"]](sec, scfg, campuses):
            loc = it["location"]
            fields = dict(it["fields"])
            fields.setdefault("location", loc)
            fields["place"] = campuses.place(loc) if loc else None
            base = f"{name}|{loc or ''}|{it['time']}"
            seen_keys[base] += 1
            t = Task(key=f"{base}|{seen_keys[base]}", section=name, label=scfg.get("label", name),
                     category=scfg.get("category"), time=it["time"], row=it["row"], location=loc,
                     place=fields["place"], fields=fields, template=scfg["template"], needs_officer=needs,
                     sort_key=since_start(it["time"], shift_start), order=order)
            if needs:
                _assign(t, it.get("done_by"), rotation, cfg, staff, campuses)
            tasks.append(t)

    # rotation within each (section, location), in time order
    groups = defaultdict(list)
    for t in tasks:
        if t.status == "choose":
            groups[(t.section, t.location)].append(t)
    for items in groups.values():
        items.sort(key=lambda t: (t.sort_key, t.row))
        for i, t in enumerate(items):
            if t.candidates:
                t.rotate_pick = t.candidates[i % len(t.candidates)]
    tasks.sort(key=lambda t: (t.sort_key, t.order, t.row))
    return tasks


def _assign(t: Task, done_by, rotation, cfg, staff: Staff, campuses: Campuses):
    loc = t.location
    site = staff.site_for(loc, campuses)

    # location rule (e.g. 34R Prentiss -> Doble Lt, then Disp, then PSO)
    for rule in cfg.get("location_rules", []):
        if loc and has_words(loc, rule["match"]):
            rsite = staff.sites.get(rule.get("site")) or site
            chain = [p for role in rule.get("roles", []) for p in (rsite.people if rsite else []) if p.role == role]
            t.rule_name = rule.get("name", "location rule")
            site = rsite
            if chain:
                t.status, t.auto, t.site = "rule", chain[0].name, rsite.name
                t.candidates = _unique([p.name for p in chain])
                t.extras = [n for n in (p.name for p in staff.manual_only(rsite)) if n not in t.candidates]
                break
    else:
        rsite = None

    if t.status != "rule":
        cover, note = staff.covering(site)
        counted = staff.counted(cover, rotation) or staff.counted(cover)
        t.site, t.site_note = (cover.name if cover else None), note
        t.candidates = _unique([p.name for p in counted])
        t.extras = [p.name for p in staff.manual_only(cover) if p.name not in t.candidates]
        if len(t.candidates) == 1:
            t.status, t.auto = "single", t.candidates[0]
        else:
            t.status = "choose"

    if done_by:  # written on the report wins over everything automatic
        t.status, t.auto = "report", staff.match_name(done_by, staff.sites.get(t.site) if t.site else None)


def _unique(names):
    out = []
    for n in names:
        if n not in out:
            out.append(n)
    return out


# ==========================================================================
# Public API
# ==========================================================================
def load_config(config_path=DEFAULT_CONFIG) -> dict:
    with open(config_path, encoding="utf-8") as fh:
        return json.load(fh)


def _open_sheet(filepath, cfg):
    if not os.path.isfile(filepath):
        raise ReportError(f"File not found: {filepath}")
    try:
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception as e:  # not an xlsx, corrupt, password protected ...
        raise ReportError(f"Couldn't open {os.path.basename(filepath)} as an Excel file ({e.__class__.__name__}).")
    name = cfg.get("sheet_name")
    return wb[name] if name in wb.sheetnames else wb.worksheets[0]


def extract(filepath, config=None) -> Report:
    cfg = config if isinstance(config, dict) else load_config(config or DEFAULT_CONFIG)
    ws = _open_sheet(filepath, cfg)
    sections = find_sections(ws)
    campuses = Campuses(cfg)

    hdr = next((s for n, s in sections.items() if matches(n, cfg["header"]["section_match"])), None)
    info = {}
    if hdr:
        for pairs in hdr.key_value_blocks().values():
            for label, value in pairs:
                for key, kws in cfg["header"]["fields"].items():
                    if key not in info and norm(label).rstrip(":. ") in kws:
                        info[key] = value
    date_val = info.get("date")
    date_str = date_val.strftime("%m/%d/%Y") if isinstance(date_val, datetime) else str(clean(date_val) or "")
    shift = str(clean(info.get("shift")) or "")
    start = parse_time(re.split(r"\s*-\s*", shift)[0]) if shift else None
    shift_start = minutes(start[0]) if start else 0

    pers = next((s for n, s in sections.items() if matches(n, cfg["personnel"]["section_match"])), None)
    staff = Staff(pers.key_value_blocks() if pers else {}, cfg)
    tasks = _build_tasks(sections, cfg, staff, campuses, shift_start)
    return Report(filepath, cfg, date_str, shift, shift_start, staff, campuses, tasks, sections)


def render(report: Report, chooser: Optional[Callable[[Task], Optional[str]]] = None) -> str:
    """chooser(task) -> officer name (or None to leave a blank). Default: automatic officer only."""
    chooser = chooser or (lambda t: t.auto)
    lines = []
    for t in report.tasks:
        f = dict(t.fields)
        if t.needs_officer:
            f["officer"] = chooser(t) or BLANK
        lines.append(render_template(t.template, f))
    return "\n".join(report.header_lines() + [""] + lines)


def default_settings(cfg, key="default"):
    return {cat: s.get(key, s["options"][0]) for cat, s in cfg.get("settings", {}).items()}


def print_sections(filepath, config_path=DEFAULT_CONFIG):
    cfg = load_config(config_path)
    sections = find_sections(_open_sheet(filepath, cfg))
    special = cfg["header"]["section_match"], cfg["personnel"]["section_match"]
    for name, sec in sections.items():
        scfg = section_config(name, cfg)
        if scfg:
            use = f"DAR ({scfg['type']})"
        elif any(matches(name, m) for m in special):
            use = "header/staff"
        elif any(k in norm(name) for k in cfg.get("not_in_dar", [])):
            use = "not in DAR"
        else:
            use = "NEW - not in config"
        title = f"title row {sec.title_row:>3}" if sec.title_row else "no title     "
        span = f"rows {sec.start_row}-{sec.end_row}" if sec.start_row <= sec.end_row else "empty"
        print(f"{name[:45]:<46} {title}  {span:<12} {use}")


def main(argv):
    args = list(argv)
    config = DEFAULT_CONFIG
    if "--config" in args:
        i = args.index("--config")
        if i + 1 >= len(args):
            print("--config needs a file path")
            return 1
        config = args[i + 1]
        del args[i:i + 2]
    files = [a for a in args if not a.startswith("--")]
    if len(files) != 1:
        print("Usage: shift_report_parser.py <report.xlsx> [--config file.json] [--sections] [--use-saved]")
        return 1
    try:
        if "--sections" in args:
            print_sections(files[0], config)
            return 0
        if "--use-saved" in args:
            from dar_session import DarSession
            s = DarSession(files[0], config, settings=default_settings(load_config(config), "cli_default"))
            s.load_picks()
            print(s.render())
            return 0
        report = extract(files[0], config)
        cli = default_settings(report.config, "cli_default")

        def chooser(t):
            if t.status != "choose":
                return t.auto
            return t.rotate_pick if cli.get(t.category) in ("rotate", "auto") else None
        print(render(report, chooser))
        return 0
    except ReportError as e:
        print(f"Error: {e}")
        return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
