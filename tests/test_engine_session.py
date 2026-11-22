import datetime as d
import json
import os
import stat
import sys

import openpyxl
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [os.path.dirname(HERE), HERE]

import fixtures as F  # noqa: E402
from dar_session import SEPARATOR, DarSession  # noqa: E402
from shift_report_parser import BLANK, ReportError, extract, main, parse_time, render  # noqa: E402

EXPECTED_CLI = """11/21/2026 0000-0800
Doble: Disp. Mostafa Kamal, Lt. Fahid Afzal, PSO. Tim Green
Uhall: PSO. David Montissol
South: PSO. Amare Dugoma

0001: PSO. Amare Dugoma started a PDA tour of South Campus
0030: PSO. David Montissol started a Visual tour of Porter Campus
0100: PSO. Tim Green started a PDA tour of Doble Campus
0100: PSO. David Montissol started a PDA tour of Porter Campus
0200: PSO. Tim Green started a Visual tour of Doble Campus
0400: Lt. Fahid Afzal started a Visual tour of Doble Campus
0400: PSO. Amare Dugoma started a Visual tour of South Campus
0500: PSO. David Montissol started a Visual tour of Porter Campus
0623: Lt. Fahid Afzal started a PDA tour of Doble Campus
0700: PSO. Amare Dugoma started a PDA tour of South Campus
0712: PSO. David Montissol started a PDA tour of Porter Campus"""

LT, PSO, DISP = "Lt. Fahid Afzal", "PSO. Tim Green", "Disp. Mostafa Kamal"
UHALL, SOUTH = "PSO. David Montissol", "PSO. Amare Dugoma"


def by(tasks, **kw):
    return [t for t in tasks if all(getattr(t, k) == v for k, v in kw.items())]


@pytest.fixture
def base():
    return F.base_copy()


@pytest.fixture
def grown():
    return F.make("grown.xlsx", start=F.grown)


# ----------------------------------------------------------------- engine
def test_cli_output_unchanged(base, capsys):
    assert main([base]) == 0
    out = capsys.readouterr().out.strip()
    assert out.split("\n")[:5] == EXPECTED_CLI.split("\n")[:5]
    assert sorted(out.split("\n")[5:]) == sorted(EXPECTED_CLI.split("\n")[5:])


def test_one_person_rule_and_dispatcher_not_counted(base):
    r = extract(base)
    doble = by(r.tasks, site="Doble")
    assert doble and all(t.status == "choose" for t in doble)
    assert all(t.candidates == [PSO, LT] for t in doble)       # dispatcher not counted
    assert all(t.extras == [DISP] for t in doble)               # ... but available by hand
    for t in by(r.tasks, site="Uhall") + by(r.tasks, site="South"):
        assert t.status == "single" and t.auto in (UHALL, SOUTH)


def test_rotation_officer_then_lt(base):
    r = extract(base)
    pda = sorted(by(r.tasks, label="PDA", site="Doble"), key=lambda t: t.sort_key)
    assert [t.rotate_pick for t in pda] == [PSO, LT]


@pytest.mark.parametrize("missing,expected", [("Officer", LT), ("Shift Supervisor", PSO)])
def test_doble_single_when_one_missing(missing, expected):
    p = F.make(f"doble_no_{missing}.xlsx", lambda ws: F.staff(ws, missing, "Doble Campus", "N/A"))
    r = extract(p)
    for t in by(r.tasks, site="Doble"):
        assert t.status == "single" and t.auto == expected


def test_prentiss_rule_chain():
    def no_lt(ws):
        F.staff(ws, "Shift Supervisor", "Doble Campus", "N/A")

    def no_lt_disp(ws):
        no_lt(ws)
        F.staff(ws, "Dispatcher", "Doble Campus", "N/A")

    for edit, expected in ((None, LT), (no_lt, DISP), (no_lt_disp, PSO)):
        wb = F.grown()
        if edit:
            edit(wb["Shift Report"])
        r = extract(F.save(wb, "prentiss.xlsx"))
        pr = [t for t in r.tasks if t.location and "34R" in t.location]
        assert len(pr) == 6                                      # 4 PDA (one past col G), 1 visual... and 1 lock
        assert all(t.status == "rule" and t.auto == expected for t in pr)


def test_prentiss_nobody_at_doble_does_not_crash():
    def nobody(ws):
        for lbl in ("Shift Supervisor", "Dispatcher", "Officer"):
            F.staff(ws, lbl, "Doble Campus", "N/A")
    wb = F.grown()
    nobody(wb["Shift Report"])
    path = F.save(wb, "nodoble.xlsx")
    r = extract(path)
    pr = [t for t in r.tasks if t.location and "34R" in t.location]
    assert pr and all(t.status in ("choose", "single") for t in pr)
    s = DarSession(path)
    assert BLANK in s.render() or not s.missing()


def test_lunder_officer_present_vs_absent(grown):
    r = extract(grown)
    lund = [t for t in r.tasks if t.location and "lunder" in t.location.lower()]
    assert lund and all(t.site == "Uhall" and "Lunder →" in t.site_note for t in lund)

    wb = F.grown()
    F.staff(wb["Shift Report"], "Lunder Officer", "Porter Square Campus", "PSO. Jared Brown")
    r = extract(F.save(wb, "lunder.xlsx"))
    lund = [t for t in r.tasks if t.location and "lunder" in t.location.lower()]
    assert all(t.site == "Lunder" and t.status == "single" and t.auto == "PSO. Jared Brown" for t in lund)
    assert "Lunder: PSO. Jared Brown" in r.header_lines()
    uhall_line = next(l for l in r.header_lines() if l.startswith("Uhall"))
    assert "Jared" not in uhall_line


def test_south_empty_falls_back_to_doble():
    p = F.make("nosouth.xlsx", lambda ws: F.staff(ws, "Officer", "South", "N/A"))
    r = extract(p)
    south = [t for t in r.tasks if t.location == "South"]
    assert south and all(t.site == "Doble" and t.status == "choose" for t in south)


def test_uhall_two_people_rotates():
    p = F.make("uhall2.xlsx", lambda ws: F.staff(ws, "Site Supervisor", "Porter Square Campus", "Lt. Rosalyn Garnett"))
    r = extract(p)
    up = sorted(by(r.tasks, label="PDA", site="Uhall"), key=lambda t: t.sort_key)
    assert [t.status for t in up] == ["choose", "choose"]
    assert [t.rotate_pick for t in up] == [UHALL, "Lt. Rosalyn Garnett"]


def test_details_training_lines():
    def edit(ws):
        F.staff(ws, "Supervisor", "Details", "Lt. Chauncey Barnes")
        F.staff(ws, "OJT 1", "Training", "PSO. Derick Willis")
    r = extract(F.make("details.xlsx", edit))
    assert "Details: Lt. Chauncey Barnes" in r.header_lines()
    assert "Training: PSO. Derick Willis" in r.header_lines()
    s = DarSession(r.path)
    t = by(s.assignable, site="Doble")[0]
    assert "Lt. Chauncey Barnes" in s.dropdown(t)       # other on-duty staff selectable


def test_done_by_column_and_inline_names():
    def edit(ws):
        pda = F.find_row(ws, "PDA Tours")
        ws.cell(pda + 1, 8).value = "Done By"
        doble = next(r for r in range(pda, pda + 8) if str(ws.cell(r, 2).value).strip() == "Doble")
        ws.cell(doble, 8).value = "Lt"                        # role word -> that person
        vis = F.find_row(ws, "Visual Tours")
        ws.cell(vis + 2, 3).value = "02:00 Afzal/04:00/Doble"  # partial name -> full name
    r = extract(F.make("doneby.xlsx", edit))
    pda = sorted(by(r.tasks, label="PDA", location="Doble"), key=lambda t: t.sort_key)
    assert pda[-1].status == "report" and pda[-1].auto == LT   # done-by after the 2nd time
    assert pda[0].status == "choose"
    vis = by(r.tasks, label="Visual", location="Doble")
    assert any(t.status == "report" and t.auto == LT and t.time == "0200" for t in vis)


def test_grown_sections_and_ordering(grown):
    r = extract(grown)
    text = render(r, lambda t: t.auto or "X")
    assert "0315: IR: Fire alarm. IR at 31 Mellen (Jenckes). Cambridge Fire contacted." in text
    assert "0400: IR: Noise complaint." in text                 # no location / dept -> pieces dropped
    assert "0545: IR: Medical." in text
    assert "33 1/2 Mellen (Facilites/Security)" in text          # slash in a location kept
    assert "escorted ABC Electric at 33 Mellen (Classrooms) per Bill Eve" in text
    assert "0641: X unlocked 24 Mellen (Mellen)" in text           # inserted lock row
    times = [l.split(":")[0] for l in text.split("\n")[r.header_lines().__len__() + 1:]]
    assert times == sorted(times)


def test_duplicate_time_location_get_distinct_keys():
    def edit(ws):
        vis = F.find_row(ws, "Visual Tours")
        ws.cell(vis + 2, 6).value = "02:00/Doble"             # same as another cell
    r = extract(F.make("dup.xlsx", edit))
    keys = [t.key for t in r.tasks]
    assert len(keys) == len(set(keys))
    assert len(by(r.tasks, label="Visual", location="Doble", time="0200")) == 2


def test_overnight_shift_sort():
    def edit(ws):
        r = F.find_row(ws, "Shift:")
        ws.cell(r, 3).value = "2200-0600"
        pda = F.find_row(ws, "PDA Tours")
        doble = next(x for x in range(pda, pda + 8) if str(ws.cell(x, 2).value).strip() == "Doble")
        ws.cell(doble, 3).value = d.time(23, 10)
        ws.cell(doble, 4).value = d.time(1, 5)
    r = extract(F.make("overnight.xlsx", edit))
    t = [x.time for x in r.tasks]
    assert t[0] == "2310" or t.index("2310") < t.index("0105")


def test_empty_report_has_no_tasks():
    def edit(ws):
        for title in ("PDA Tours", "Visual Tours"):
            top = F.find_row(ws, title)
            for rr in range(top + 2, top + 6):
                for c in range(2, 9):
                    if not type(ws.cell(rr, c)).__name__ == "MergedCell" and ws.cell(rr, 2).value not in ("Ticketing",):
                        if c > 2 or title == "Visual Tours":
                            ws.cell(rr, c).value = None
    s = DarSession(F.make("empty.xlsx", edit))
    assert s.assignable == [] and s.missing() == []
    assert s.render().startswith("11/21/2026 0000-0800")


@pytest.mark.parametrize("val,exp", [("4:00", "0400"), ("04:00 PM", "1600"), ("12:30 AM", "0030"),
                                     ("0400", "0400"), (d.time(7, 5), "0705"), (0.5, "1200"),
                                     ("25:00", None), ("13:00 PM", None), ("Doble", None), ("33 1", None),
                                     (True, None), (7, None)])
def test_parse_time(val, exp):
    got = parse_time(val)
    assert (got[0] if got else None) == exp


def test_bad_files(tmp_path):
    with pytest.raises(ReportError):
        extract(str(tmp_path / "nope.xlsx"))
    bad = tmp_path / "bad.xlsx"
    bad.write_text("not excel")
    with pytest.raises(ReportError):
        extract(str(bad))
    wb = openpyxl.Workbook()
    wb.active["A1"] = "hello"
    p = tmp_path / "plain.xlsx"
    wb.save(p)
    with pytest.raises(ReportError):
        extract(str(p))
    assert main([str(bad)]) == 2


# ----------------------------------------------------------------- session
def test_session_ask_mode_requires_picks(base):
    s = DarSession(base)
    miss = s.missing()
    assert len(miss) == 4 and all(t.site == "Doble" for t in miss)
    assert s.render().count(BLANK) == 4
    for t in miss:
        s.set_officer(t.key, LT)
    assert s.missing() == [] and BLANK not in s.render()


def test_session_rotate_random_auto(grown):
    s = DarSession(grown, settings={"pda": "rotate", "visual": "random", "lock": "auto"}, seed=1)
    for t in s.assignable:
        if t.status == "choose":
            assert s.officer(t) in t.candidates
            assert s.state(t).startswith("prefill:")
    assert s.missing() == []
    before = {t.key: s.officer(t) for t in s.assignable}
    assert s.render() == s.render()                              # random is stable between renders
    s.set_setting("pda", "ask")
    assert all(s.officer(t) is None for t in s.assignable if t.category == "pda" and t.status == "choose")
    assert all(s.officer(t) == before[t.key] for t in s.assignable if t.category == "visual")
    assert not s.set_setting("pda", "nonsense") and not s.set_setting("nope", "ask")


def test_manual_beats_everything_and_reset(grown):
    s = DarSession(grown, settings={"pda": "rotate"})
    single = next(t for t in s.assignable if t.status == "single")
    rule = next(t for t in s.assignable if t.status == "rule")
    choose = next(t for t in s.assignable if t.status == "choose" and t.category == "pda")
    s.set_officer(single.key, DISP)
    s.set_officer(rule.key, PSO)
    s.set_officer(choose.key, DISP)
    assert (s.officer(single), s.officer(rule), s.officer(choose)) == (DISP, PSO, DISP)
    assert s.state(single) == s.state(rule) == s.state(choose) == "manual"
    s.set_setting("pda", "random")                               # settings never touch hand picks
    assert s.officer(choose) == DISP
    s.set_officer(single.key, single.auto)                       # picking the automatic person = no hand pick
    assert single.key not in s.manual and s.state(single) == "single"
    s.reset(rule.key)
    assert s.officer(rule) == rule.auto
    s.set_officer(choose.key, SEPARATOR)                         # separator = clear
    assert choose.key not in s.manual


def test_dropdown_groups(base):
    s = DarSession(base)
    t = by(s.assignable, site="Doble")[0]
    v = s.dropdown(t)
    assert v[:2] == [PSO, LT] and v[2] == SEPARATOR and v[3] == DISP
    assert UHALL in v and SOUTH in v                             # everyone else on duty after
    t = by(s.assignable, site="South")[0]
    assert s.dropdown(t)[0] == SOUTH


def test_picks_save_load_roundtrip(grown):
    s = DarSession(grown, settings={"visual": "random"}, seed=3)
    choose = [t for t in s.assignable if t.status == "choose"]
    s.set_officer(choose[0].key, DISP)
    vis_random = {t.key: s.officer(t) for t in s.assignable if t.category == "visual" and t.status == "choose"}
    s.save_picks()
    s2 = DarSession(grown, settings={"visual": "random"}, seed=99)
    assert s2.load_picks() == ""
    assert s2.officer(s2.task(choose[0].key)) == DISP
    assert {k: s2.officer(s2.task(k)) for k in vis_random} == vis_random   # random picks kept on reopen
    s3 = DarSession(grown, settings={"visual": "ask"})
    s3.load_picks()
    assert all(s3.officer(s3.task(k)) is None for k in vis_random)       # setting changed -> not restored


def test_picks_corrupt_stale_and_readonly(grown, tmp_path):
    s = DarSession(grown)
    with open(s.picks_path, "w") as fh:
        fh.write("{not json")
    assert "couldn't be read" in DarSession(grown).load_picks()
    with open(s.picks_path, "w") as fh:
        json.dump({"version": 1, "manual": {"gone|x|0000|1": "Someone"}}, fh)
    assert "no longer match" in DarSession(grown).load_picks()
    with open(s.picks_path, "w") as fh:
        json.dump({"version": 42}, fh)
    assert "unknown format" in DarSession(grown).load_picks()
    os.remove(s.picks_path)
    ro = tmp_path / "ro"
    ro.mkdir()
    p = ro / "r.xlsx"
    p.write_bytes(open(grown, "rb").read())
    s = DarSession(str(p))
    os.chmod(ro, stat.S_IRUSR | stat.S_IXUSR)
    try:
        if os.geteuid() != 0:   # root ignores permissions
            with pytest.raises(OSError):
                s.save_picks()
    finally:
        os.chmod(ro, stat.S_IRWXU)


def test_cli_use_saved(base, capsys):
    s = DarSession(base)
    for t in s.missing():
        s.set_officer(t.key, DISP)
    s.save_picks()
    main([base, "--use-saved"])
    out = capsys.readouterr().out
    assert out.count(DISP) == 1 + 4 and BLANK not in out
    os.remove(s.picks_path)


def test_packaged_paths(tmp_path, monkeypatch):
    import shift_report_parser as P
    exe_dir = tmp_path / "ShiftReportDAR"
    bundle = exe_dir / "_internal"
    bundle.mkdir(parents=True)
    (bundle / "report_config.json").write_text("{}")
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(exe_dir / "ShiftReportDAR.exe"))
    monkeypatch.setattr(sys, "_MEIPASS", str(bundle), raising=False)
    assert P.app_dir() == str(exe_dir)
    assert P._default_config() == str(bundle / "report_config.json")       # bundled copy
    (exe_dir / "report_config.json").write_text("{}")
    assert P._default_config() == str(exe_dir / "report_config.json")      # edited copy next to .exe wins
