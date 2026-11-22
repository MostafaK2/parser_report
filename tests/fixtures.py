"""Builds test workbooks from the real report by editing a copy (rows inserted like Excel does)."""
import datetime as d
import os
import shutil
import warnings

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(HERE, "report0.xlsx")
OUT = os.path.join(HERE, "out")


def _load():
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return openpyxl.load_workbook(BASE)


def find_row(ws, text, col=2):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().lower() == text.lower():
            return r
    raise KeyError(text)


def insert(ws, at, n=1):
    """Insert rows and move merged cells, as Excel does."""
    ranges = list(ws.merged_cells.ranges)
    for r in ranges:
        ws.merged_cells.remove(r)
    ws.insert_rows(at, n)
    for r in ranges:
        if r.min_row >= at:
            r.shift(0, n)
        ws.merged_cells.add(r)


def staff(ws, label, block_title, value):
    """Set a personnel value by its label inside a block (e.g. 'Officer' under 'Doble Campus')."""
    top = find_row(ws, "Public Safety Personnel on Duty")
    for col_label, col_val in ((2, 3), (4, 5)):
        block = None
        for r in range(top + 1, top + 14):
            v = ws.cell(r, col_label).value
            if isinstance(v, str):
                rng = next((m for m in ws.merged_cells.ranges if m.min_row == r and m.min_col == col_label), None)
                if rng and rng.max_col > col_label and (col_label == 2 and rng.max_col == 3 or col_label == 4):
                    block = v.strip()
                    continue
                if block and block.lower().startswith(block_title.lower()) and v.strip().lower() == label.lower():
                    ws.cell(r, col_val).value = value
                    return
    raise KeyError((label, block_title))


def save(wb, name):
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, name)
    wb.save(path)
    return path


def base_copy(name="base.xlsx"):
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, name)
    shutil.copy(BASE, path)
    return path


def grown(name="grown.xlsx"):
    """Incidents, PDA locations (incl. 34R Prentiss), visual cells, locks, escorts all grown."""
    wb = _load()
    ws = wb["Shift Report"]
    inc = find_row(ws, "Incident Report")
    insert(ws, inc + 2, 2)
    ws.cell(inc + 2, 2).value = d.time(3, 15); ws.cell(inc + 2, 3).value = "Fire alarm"
    ws.cell(inc + 2, 4).value = "31 Mellen (Jenckes)"; ws.cell(inc + 2, 5).value = "Cambridge Fire"
    ws.cell(inc + 3, 2).value = "5:45 AM"; ws.cell(inc + 3, 3).value = "Medical"
    ws.cell(inc + 3, 4).value = "3 Phillips (Grad)"; ws.cell(inc + 3, 5).value = "Ambulance/EMT"
    ws.cell(inc + 4, 2).value = d.time(4, 0); ws.cell(inc + 4, 3).value = "Noise complaint"

    pda = find_row(ws, "PDA Tours")
    south = find_row(ws, "South")  # PDA South row (first 'South' in col B after personnel 'South ')
    south = next(r for r in range(pda, pda + 8) if str(ws.cell(r, 2).value).strip() == "South")
    insert(ws, south + 1, 1)
    ws.cell(south + 1, 2).value = "34R Prentiss/34 Prentiss"
    ws.cell(south + 1, 3).value = d.time(0, 45); ws.cell(south + 1, 4).value = d.time(3, 30)
    ws.cell(south + 1, 8).value = d.time(7, 40)            # past the form's last column
    doble = next(r for r in range(pda, pda + 8) if str(ws.cell(r, 2).value).strip() == "Doble")
    ws.cell(doble, 5).value = d.time(7, 50)                 # 3rd Doble tour
    lunder = next(r for r in range(pda, pda + 8) if str(ws.cell(r, 2).value).strip() == "Lunder")
    ws.cell(lunder, 3).value = d.time(2, 15)

    vis = find_row(ws, "Visual Tours")
    ws.cell(vis + 2, 2).value = "Doble/02:00/04:00"         # location first
    ws.cell(vis + 2, 3).value = "04:00/South"
    ws.cell(vis + 2, 4).value = None
    ws.cell(vis + 2, 6).value = "00:30/05:00/34R Prentiss/34 Prentiss"
    insert(ws, vis + 3, 1)
    ws.cell(vis + 3, 2).value = "06:00/Uhall"
    ws.cell(vis + 3, 3).value = "06:30/33 1/2 Mellen (Facilites/Security)"

    rl = find_row(ws, "Routine Locks/Unlocks")
    ws.cell(rl + 2, 2).value = d.time(0, 15); ws.cell(rl + 2, 3).value = "30 Mellen (Doble Hall)"
    ws.cell(rl + 2, 4).value = "Locked"
    sl = find_row(ws, "Special Lock/Unlocks")
    ws.cell(sl + 2, 2).value = "05:30"; ws.cell(sl + 2, 3).value = "3 Phillips (Grad)"; ws.cell(sl + 2, 4).value = "Unlocked"
    ws.cell(sl + 3, 2).value = d.time(6, 10); ws.cell(sl + 3, 3).value = "1801 Mass Ave (Lunder)"; ws.cell(sl + 3, 4).value = "Locked"
    ws.cell(sl + 4, 2).value = d.time(6, 20); ws.cell(sl + 4, 3).value = "34R Prentiss"; ws.cell(sl + 4, 4).value = "Unlocked"
    insert(ws, sl + 5, 1)  # more lock rows than the template had
    ws.cell(sl + 5, 2).value = d.time(6, 40); ws.cell(sl + 5, 3).value = "24 Mellen (Mellen)"; ws.cell(sl + 5, 4).value = "Locked"
    ws.cell(sl + 6, 2).value = d.time(6, 41); ws.cell(sl + 6, 3).value = "24 Mellen (Mellen)"; ws.cell(sl + 6, 4).value = "Unlocked"

    ce = find_row(ws, "Contractor Escorts")
    ws.cell(ce + 2, 2).value = d.time(2, 30); ws.cell(ce + 2, 3).value = "33 Mellen (Classrooms)"
    ws.cell(ce + 2, 4).value = "ABC Electric"; ws.cell(ce + 2, 6).value = "Bill Eve"
    return wb


def make(name, edit=None, start=None):
    wb = start() if start else _load()
    if edit:
        edit(wb["Shift Report"])
    return save(wb, name)
