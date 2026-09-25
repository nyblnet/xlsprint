"""Pure tests for the Windows runner and a static lint of the VBA module.

Nothing here launches Excel. The VBA is checked structurally and its JSON
escape table is exercised through a Python port driven by the table parsed
out of the .bas itself.
"""

import json
import re
import sys
from pathlib import Path

import pytest

from xlsprint import runner
from xlsprint.runner import RS, US, ProfileOptions, RunnerError

BAS = Path(runner.__file__).with_name("vba") / "XLSprintTimer.bas"


def opts(**kw):
    base = dict(workbook=Path("book.xlsx"), out_dir=Path("out"))
    base.update(kw)
    return ProfileOptions(**base)


PLAN_INFO = {
    "sheets": [
        {"sheet": "Inputs", "formula_cells": 0, "formula_area": None},
        {"sheet": "Calc", "formula_cells": 500, "formula_area": "B2:K101", "array_areas": []},
        {"sheet": "My Sheet", "formula_cells": 3, "formula_area": "C3:D4"},
    ],
    "names": [
        {"name": "Totals", "scope": "workbook", "sheet": "Calc", "address": "$K$2:$K$101"},
        {"name": "Totals", "scope": "Calc", "sheet": "Calc", "address": "$B$2"},
    ],
    "groups": [
        {"group": "G0001", "sheet": "Calc", "areas": ["B2:B101", "D2:D101"], "areas_total": 2},
        {"group": "G0002", "sheet": "Calc", "areas": ["E2:E3"] * 32, "areas_total": 40},
    ],
}


# --------------------------------------------------------------------------
# Plan building


def test_split_column_blocks_even_and_capped():
    assert runner.split_column_blocks("B2:K101", 4) == [
        "$B$2:$D$101", "$E$2:$G$101", "$H$2:$I$101", "$J$2:$K$101"]
    assert runner.split_column_blocks("C3:D4", 4) == ["$C$3:$C$4", "$D$3:$D$4"]
    assert runner.split_column_blocks("$A$1", 4) == ["$A$1"]
    assert runner.split_column_blocks("A1:Z9", 0) == []
    # reversed corners are normalized
    assert runner.split_column_blocks("D4:C3", 1) == ["$C$3:$D$4"]


def test_split_blocks_cover_area_exactly():
    blocks = runner.split_column_blocks("A1:XFD5", 7)
    cols = []
    for b in blocks:
        r1, c1, r2, c2 = runner.parse_area(b)
        assert (r1, r2) == (1, 5)
        cols.extend(range(c1, c2 + 1))
    assert cols == list(range(1, 16385))


@pytest.mark.parametrize("bad", ["A0", "XFE1", "A1048577", "A1:B", "Sheet1!A1", "A1,B2", ""])
def test_parse_area_rejects(bad):
    with pytest.raises(RunnerError):
        runner.parse_area(bad)


def test_parse_user_range():
    assert runner.parse_user_range("Calc!a1:b2") == ("Calc", "$A$1:$B$2")
    assert runner.parse_user_range("'It''s here'!C3") == ("It's here", "$C$3")
    with pytest.raises(RunnerError):
        runner.parse_user_range("A1:B2")


def test_build_plan_default_order():
    steps = runner.build_plan(PLAN_INFO, opts(ranges=["Calc!A1:A5"], macros=["Module1.Refresh"]))
    assert steps[0] == ("full",) and steps[1] == ("recalc",)
    assert [s for s in steps if s[0] == "sheet"] == [("sheet", "Calc"), ("sheet", "My Sheet")]
    ranges = [s for s in steps if s[0] == "range"]
    assert ranges[0] == ("range", "Calc", "$A$1:$A$5", "user")
    assert ranges[1:5] == [
        ("range", "Calc", "$B$2:$D$101", "auto 1/4"), ("range", "Calc", "$E$2:$G$101", "auto 2/4"),
        ("range", "Calc", "$H$2:$I$101", "auto 3/4"), ("range", "Calc", "$J$2:$K$101", "auto 4/4")]
    assert ranges[5:] == [("range", "My Sheet", "$C$3:$C$4", "auto 1/2"), ("range", "My Sheet", "$D$3:$D$4", "auto 2/2")]
    assert [s for s in steps if s[0] == "name"] == [
        ("name", "Calc", "$K$2:$K$101", "Totals", "workbook"), ("name", "Calc", "$B$2", "Totals", "Calc")]
    assert not [s for s in steps if s[0] == "group"]
    assert steps[-1] == ("macro", "Module1.Refresh")
    kinds = [s[0] for s in steps]
    order = ["full", "recalc", "sheet", "range", "name", "macro"]
    assert kinds == sorted(kinds, key=order.index)


def test_build_plan_gating():
    steps = runner.build_plan(PLAN_INFO, opts(names_timing=False, group_timing=True, full_rebuild=True, blocks=1))
    assert steps[:3] == [("fullrebuild",), ("full",), ("recalc",)]
    assert not [s for s in steps if s[0] == "name"]
    # G0002 exceeds 32 areas and is skipped
    assert [s for s in steps if s[0] == "group"] == [("group", "Calc", "B2:B101,D2:D101", "G0001")]
    assert [s for s in steps if s[0] == "range"] == [
        ("range", "Calc", "$B$2:$K$101", "auto 1/1"), ("range", "My Sheet", "$C$3:$D$4", "auto 1/1")]


def test_extend_with_arrays_disjoint():
    # block B2:D10 touches array C9:C12 and array D1:E2; F1:F5 is untouched
    out = runner.extend_with_arrays(["B2:D10"], ["C9:C12", "D1:E2", "F1:F5"])
    rects = [runner.parse_area(a) for a in out]
    assert rects[0] == (2, 2, 10, 4)
    cells = [(r, c) for r1, c1, r2, c2 in rects for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
    assert len(cells) == len(set(cells))  # no cell twice
    want = {(r, c) for r in range(2, 11) for c in range(2, 5)} | {(11, 3), (12, 3), (1, 4), (1, 5), (2, 5)}
    assert set(cells) == want
    # an array fully inside the block adds nothing; repeated input areas are merged
    assert runner.extend_with_arrays(["A1:C3", "A1:C3"], ["B2:B3"]) == ["$A$1:$C$3"]
    assert sorted(runner.extend_with_arrays(["A1:B2", "B2:C3"], [])) == ["$A$1:$B$2", "$B$3:$C$3", "$C$2"]


def test_compile_plan_keys_names_and_arrays():
    info = dict(PLAN_INFO, sheets=[dict(s) for s in PLAN_INFO["sheets"]])
    info["sheets"][1]["array_areas"] = ["L2:M3", "D100:D102"]
    steps = runner.build_plan(info, opts(ranges=["Calc!A1:A5"], group_timing=True, macros=["M"]))
    ident = lambda s: "h:" + s.lower()  # noqa: E731
    vba, meta = runner.compile_plan(steps, info, ident)
    enc = runner.encode_plan(vba)
    assert all(US.join(s) in enc for s in vba)
    ranges = [s for s in vba if s[0] == "range"]
    assert [s[3] for s in ranges] == ["R%04d" % i for i in range(1, len(ranges) + 1)]
    r1 = ranges[0]
    assert r1[1:3] == ("Calc", "$A$1:$A$5") and meta[r1[3]] == {"name": "h:calc!$A$1:$A$5", "sheet": "Calc", "note": "user"}
    # the auto block B:D reaches row 101 and picks up the array D100:D102
    bd = next(s for s in ranges if s[2].startswith("$B$2:$D$101"))
    assert bd[2] == "$B$2:$D$101,$D$102"
    assert meta[bd[3]]["name"] == "h:calc!$B$2:$D$101"  # trace name is the planned block
    names = [s for s in vba if s[0] == "name"]
    assert [meta[s[3]]["name"] for s in names] == ["h:totals", "h:calc!h:totals"]  # scope keeps them apart
    assert len({s[3] for s in names}) == 2
    grp = next(s for s in vba if s[0] == "group")
    assert grp[3] == "G0001" and meta["G0001"] == {"name": "G0001", "sheet": "Calc", "group": "G0001"}
    assert ("macro", "M") in vba and ("sheet", "Calc") in vba


def test_encode_plan_separators():
    enc = runner.encode_plan([("full",), ("sheet", "A B"), ("range", "S", "$A$1,$B$2", "R0001"), ("macro", "M")])
    assert enc == "full" + RS + "sheet" + US + "A B" + RS + "range" + US + "S" + US + "$A$1,$B$2" + US + "R0001" + RS + "macro" + US + "M"
    assert [s.split(US) for s in enc.split(RS)][2] == ["range", "S", "$A$1,$B$2", "R0001"]


@pytest.mark.parametrize("bad", [
    [("sheet",)], [("sheet", "")], [("sheet", "a" + RS)], [("range", "S", "A1" + US, "R0001")],
    [("bogus",)], [()], [("full", "extra")], [("range", "S", "$A$1", "user")],
    [("name", "S", "$A$1", "N0001", "workbook")],
])
def test_encode_plan_rejects(bad):
    with pytest.raises(RunnerError):
        runner.encode_plan(bad)


def test_plan_file_is_utf16_with_bom(tmp_path):
    enc = runner.encode_plan([("full",), ("sheet", "Mod\u00e8le \u65e5\u672c " + "x" * 400)])
    p = runner.write_plan_file(tmp_path / "plan.txt", enc)
    raw = p.read_bytes()
    assert raw[:2] == b"\xff\xfe"
    assert raw[2:].decode("utf-16-le") == enc
    assert len(enc) > 255  # longer than an Application.Run string argument may carry


def test_adapt_plan_info_from_formulas_module(tmp_path):
    formulas = pytest.importorskip("xlsprint.formulas")
    synthbook = pytest.importorskip("xlsprint.synthbook")
    trace = pytest.importorskip("xlsprint.trace")
    book = synthbook.make_synthetic_workbook(tmp_path / "synth.xlsx", seed=0)
    fobj, plan = formulas.inspect_workbook_and_plan(book, names="hashed", salt="pepper")
    # trace names hashed by the runner match formulas.json's hashed names
    for hashed, clear in plan.get("redaction_map", {}).items():
        assert trace.redact_name(clear, "pepper") == hashed
    info = runner.adapt_plan_info(plan)
    assert info["sheets"] and all(isinstance(s["sheet"], str) for s in info["sheets"])
    assert any(s["formula_cells"] > 0 and s["formula_area"] for s in info["sheets"])
    for n in info["names"]:
        runner.parse_area(n["address"])
    steps = runner.build_plan(info, opts(group_timing=True))
    vba, meta = runner.compile_plan(steps, info, runner.make_ident("hashed", "pepper"))
    runner.encode_plan(vba)  # every field is encodable
    assert {"full", "recalc", "sheet", "range"} <= {s[0] for s in vba}
    for m in meta.values():
        assert trace.check_name(m["name"]) is None


def test_adapt_plan_info_names():
    info = runner.adapt_plan_info({
        "sheets": [{"sheet": "S", "formula_cells": 2, "formula_area": None, "used_range": "B2:D10"},
                   {"sheet": "E", "formula_cells": 0, "formula_area": None, "used_range": None}],
        "names": [
            {"name": "N", "refers_to_range": "'S x'!$A$1:$A$2", "is_range": True, "formula_cells": 2, "has_formulas": True},
            {"name": "K", "is_range": False},
            {"name": "Z", "sheet": "S", "address": "$B$2", "formula_cells": 0, "has_formulas": False},
            {"name": "Col", "scope": "S", "sheet": "S", "address": "$C:$C", "formula_cells": 3, "has_formulas": True},
            {"name": "Rows", "sheet": "S", "address": "$3:$4", "formula_cells": 3, "has_formulas": True},
            {"name": "Far", "sheet": "S", "address": "$Z:$Z", "formula_cells": 1, "has_formulas": True},
            {"name": "NoUsed", "sheet": "E", "address": "$A:$A", "formula_cells": 1, "has_formulas": True},
        ],
    })
    assert info["names"] == [
        {"name": "N", "scope": "workbook", "sheet": "S x", "address": "$A$1:$A$2"},
        {"name": "Col", "scope": "S", "sheet": "S", "address": "$C$2:$C$10"},
        {"name": "Rows", "scope": "workbook", "sheet": "S", "address": "$B$3:$D$4"},
    ]
    assert info["skipped_names"] == {"not_a_range": 1, "no_formulas": 1, "outside_used_range": 2}
    assert info["sheets"][0]["formula_area"] is None  # used_range is not a formula area
    assert info["sheets"][0]["array_areas"] == []
    assert info["groups"] == []


def test_clip_to_used():
    assert runner.clip_to_used("a1:b2", None) == "$A$1:$B$2"
    assert runner.clip_to_used("$A:$C", "B5:Z9") == "$B$5:$C$9"
    assert runner.clip_to_used("$7:$2", "B5:Z9") == "$B$5:$Z$7"
    assert runner.clip_to_used("$A:$A", "B5:Z9") is None
    assert runner.clip_to_used("$A:$A", None) is None


# --------------------------------------------------------------------------
# Paths, hashing, argv, redaction


def test_check_paths(tmp_path):
    wb = tmp_path / "books" / "a.xlsx"
    wb.parent.mkdir()
    wb.write_bytes(b"x")
    out = tmp_path / "out"
    assert runner.check_paths(wb, out) == (wb.resolve(), out.resolve())
    with pytest.raises(RunnerError, match="UNC"):
        runner.check_paths(wb, r"\\server\share\out")
    with pytest.raises(RunnerError, match="UNC"):
        runner.check_paths("//server/share/a.xlsx", out)
    with pytest.raises(RunnerError, match="parent"):
        runner.check_paths(wb, tmp_path / "nope" / "out")
    with pytest.raises(RunnerError, match="own folder"):
        runner.check_paths(wb, wb.parent)
    with pytest.raises(RunnerError, match="not found"):
        runner.check_paths(tmp_path / "missing.xlsx", out)
    with pytest.raises(RunnerError, match="network"):
        runner.check_paths(wb, out, is_network=lambda p: p == out.resolve())
    for ext in (".csv", ".xlsb", ".xls"):
        bad = tmp_path / ("a" + ext)
        bad.write_text("1")
        with pytest.raises(RunnerError, match="unsupported"):
            runner.check_paths(bad, out)


def test_sha256_and_verify(tmp_path):
    p = tmp_path / "f.bin"
    p.write_bytes(b"abc")
    digest = runner.sha256_file(p)
    assert digest == "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    runner.verify_unchanged(p, digest)
    p.write_bytes(b"abd")
    with pytest.raises(RunnerError, match="changed"):
        runner.verify_unchanged(p, digest)


ARGV = [r"C:\Python\Scripts\xlsprint.exe", "profile", r"C:\Users\me\Secret Plans\q3.xlsx",
        "--out", "/home/me/out/", "--out=/tmp/x/y", "--repeats", "5", "--range", "'Pay roll'!A1:B2",
        "--macro", "Module1.Payroll", "--range=Calc!C3", "--macro=Other"]


def test_reduce_argv_clear():
    ident = runner.make_ident("clear", "s")
    assert runner.reduce_argv(ARGV, ident=ident) == [
        "xlsprint.exe", "profile", "q3.xlsx", "--out", "out", "--out=y", "--repeats", "5",
        "--range", "Pay roll!$A$1:$B$2", "--macro", "Module1.Payroll", "--range=Calc!$C$3", "--macro=Other"]


def test_reduce_argv_hashed():
    trace = pytest.importorskip("xlsprint.trace")
    ident = runner.make_ident("hashed", "s")
    out = runner.reduce_argv(ARGV, ident=ident, hash_files=True)
    h = lambda s: trace.redact_name(s, "s")  # noqa: E731
    assert out == [
        "xlsprint.exe", "profile", h("q3") + ".xlsx", "--out", "out", "--out=y", "--repeats", "5",
        "--range", h("Pay roll") + "!$A$1:$B$2", "--macro", h("Module1.Payroll"),
        "--range=" + h("Calc") + "!$C$3", "--macro=" + h("Other")]
    joined = " ".join(out)
    for leak in ("Secret", "q3.", "Pay roll", "Payroll", "Calc", "Other"):
        assert leak not in joined


def test_make_ident_rule10():
    trace = pytest.importorskip("xlsprint.trace")
    clear = runner.make_ident("clear", "s")
    hashed = runner.make_ident("hashed", "s")
    assert clear("Sheet1") == "Sheet1" and clear("Mod\u00e8le") == "Mod\u00e8le"
    assert hashed("Sheet1") == trace.redact_name("Sheet1", "s")
    for bad in ("=Sheet", "x" * 300, "a\tb"):
        assert clear(bad) == trace.redact_name(bad, "s")
    # whatever trace.check_name rejects is hashed even in clear mode
    fake = runner.make_ident("clear", "s", check_name=lambda n: "bad" if "*" in n else None)
    assert fake("a*b") == trace.redact_name("a*b", "s") and fake("ab") == "ab"


def test_scrub_text():
    ident = lambda s: "h:" + str(len(s))  # noqa: E731
    msg = (r"com_error: 'C:\Users\me\Secret Plans\q3.xlsx' could not be found; sheet Payroll 2024 "
           "failed in /home/me/out/work/abc/q3.xlsx (code 1004)")
    out = runner.scrub_text(msg, paths=[r"C:\Users\me\Secret Plans\q3.xlsx"],
                            idents=["Payroll 2024", "q3.xlsx", "ab"], ident=ident)
    assert "Secret" not in out and "/home/me" not in out and "Payroll" not in out
    assert "<path>" in out and "h:12" in out and "(code 1004)" in out
    # clear mode: paths go, identifiers stay
    assert "Payroll 2024" in runner.scrub_text(msg, idents=["Payroll 2024"])
    # an unquoted drive path may contain spaces, so the rest of the line goes too (safe side)
    assert runner.scrub_text(r"see D:\data\x.xlsx now") == "see <path>"
    assert runner.scrub_text(r"open 'C:\My Books\q.xlsx' failed") == "open '<path>' failed"
    assert runner.scrub_text(r"x \\srv\share\f.xlsx y") == "x <path> y"


def test_profile_refuses_non_windows(tmp_path):
    if sys.platform == "win32":
        pytest.skip("non-Windows guard")
    with pytest.raises(RunnerError, match="Windows"):
        runner.profile(opts(workbook=tmp_path / "a.xlsx", out_dir=tmp_path / "o"))


def test_profile_options_timeout_default():
    assert opts().timeout_s == 900.0


# --------------------------------------------------------------------------
# Clock, footer, failures, watchdog


def test_choose_probe():
    off, unc = runner.choose_probe([(100, 5_000, 300), (1_000, 6_050, 1_100), (2_000, 7_000, 2_400)])
    assert (off, unc) == (6_050 - 1_050, 50)
    with pytest.raises(RunnerError):
        runner.choose_probe([])


def test_check_drift():
    assert runner.check_drift(40, 50) == "ok"
    assert runner.check_drift(-60, 50) == "warn"
    assert runner.check_drift(900_000, 50_000) == "warn"  # above uncertainty but within 10x
    assert runner.check_drift(600_000, 50) == "warn"  # >10x but below 1 ms
    assert runner.check_drift(-2_000_000, 100_000) == "fail"


def test_check_vba_footer():
    ok = {"events": 4, "dropped": 0, "truncated": False, "open_spans": 0, "faults": 0}
    assert runner.check_vba_footer(ok) == []
    assert runner.check_vba_footer(dict(ok, truncated=True, dropped=3))
    assert runner.check_vba_footer(dict(ok, open_spans=1))
    assert "end_not_innermost" in runner.check_vba_footer(dict(ok, faults=2, first_fault="end_not_innermost"))[0]


def test_span_failures_split_critical():
    evs = [{"type": "B", "id": 1, "kind": "run.pass", "name": "pass"},
           {"type": "B", "id": 2, "kind": "calc.full", "name": "workbook"},
           {"type": "E", "id": 2, "status": "error", "attrs": {"error_code": 1004}},
           {"type": "B", "id": 3, "kind": "calc.name", "name": "N0001"},
           {"type": "E", "id": 3, "status": "error", "attrs": {"error_code": 13}},
           {"type": "E", "id": 1, "status": "ok"}]
    critical, other = runner.span_failures([(2, "off", evs)])
    assert critical == [{"pass": 2, "mode": "off", "kind": "calc.full", "status": "error", "error_code": 1004}]
    assert other == [{"pass": 2, "mode": "off", "kind": "calc.name", "status": "error", "error_code": 13}]
    assert "N0001" not in json.dumps(critical + other)


def test_watchdog_kills_and_raises():
    import threading
    killed = threading.Event()
    wd = runner.Watchdog()
    try:
        wd.set_kill(killed.set)
        with pytest.raises(RunnerError, match="timed out in profile_trace_on pass 1"):
            with wd.guard("profile_trace_on pass 1", 0.05):
                # stands in for a COM call that fails once Excel is gone
                assert killed.wait(5)
                raise OSError("RPC server unavailable")
        killed.clear()
        with wd.guard("quick", 5):
            pass
        assert not killed.wait(0.1) and wd.fired is None
        # a real error inside a guard that did not fire passes through unchanged
        with pytest.raises(ValueError):
            with wd.guard("other", 5):
                raise ValueError("x")
        # the deadline fired but the call returned anyway: still a timeout
        with pytest.raises(RunnerError, match="timed out in late"):
            with wd.guard("late", 0.02):
                assert killed.wait(5)
    finally:
        wd.stop()


# --------------------------------------------------------------------------
# Merge a hand-written VBA events file into a TraceWriter


def _vba_file(path, pass_id, mode, base_ns):
    """VBA-style event file: CRLF, \\u escapes, ids from 1, root parent 0, keys for ranges/names."""
    ns = iter(range(base_ns, base_ns + 10_000, 100))
    lines = []

    def b(i, parent, depth, kind, name, attrs):
        lines.append({"type": "B", "id": i, "parent": parent, "pass": pass_id, "depth": depth, "ns": next(ns),
                      "clock": "vba", "kind": kind, "name": name, "attrs": attrs})

    def e(i, status="ok", attrs=None):
        ev = {"type": "E", "id": i, "ns": next(ns), "clock": "vba", "status": status}
        if attrs:
            ev["attrs"] = attrs
        lines.append(ev)

    b(1, 0, 0, "run.pass", "pass", {"overhead_mode": mode, "calc_mode": "manual", "threads": 8, "multithreaded": True})
    b(2, 1, 1, "calc.full", "workbook", {"method": "Application.CalculateFull"}); e(2)
    b(3, 1, 1, "calc.recalc", "workbook", {"method": "Application.Calculate"}); e(3)
    if mode == "on":
        b(4, 1, 1, "calc.sheet", SHEET, {"method": "Worksheet.Calculate"}); e(4)
        b(5, 1, 1, "calc.range", "R0001", {"method": "Range.Calculate", "address": "$A$1:$B$2", "cells": 4}); e(5)
        b(6, 1, 1, "calc.name", "N0001", {"method": "Range.Calculate"}); e(6, "error", {"error_code": 1004})
        b(7, 1, 1, "vba.proc", "Module1.Refresh", {"method": "Application.Run"})
        b(8, 7, 2, "vba.proc", "Inner", {})
        lines.append({"type": "M", "id": 9, "parent": 8, "pass": pass_id, "ns": next(ns), "clock": "vba",
                      "kind": "marker", "name": "mid", "attrs": {}})
        e(8); e(7)
    e(1)
    n = len(lines)
    lines.append({"type": "vba_footer", "events": n, "dropped": 0, "truncated": False, "open_spans": 0,
                  "faults": 0, "attr_rejected": 0, "first_fault": "", "mode": mode, "pass": pass_id, "max_events": 100000})
    text = "\r\n".join(json.dumps(x, ensure_ascii=True, separators=(",", ":")) for x in lines) + "\r\n"
    path.write_bytes(text.encode("ascii"))
    return path


SHEET = "Caf\u00e9 \"Q3\""  # fails NAME_RE ('"'), so it is hashed even with --names clear
META_CLEAR = {"R0001": {"name": "Calc!$A$1:$B$2", "sheet": "Calc", "note": "auto 1/1"},
              "N0001": {"name": "Calc!Totals", "sheet": "Calc"}}


def _stages_and_events(tmp_path, trace):
    stages, events = [], {}
    t = 1_000_000
    for name in ("prepare_copy", "launch_excel", "open_workbook", "instrument", "warmup"):
        stages.append({"name": name, "start": t, "end": t + 500, "status": "ok", "attrs": {}})
        t += 1_000
    for k in (1, 2):
        for mode in ("off", "on"):
            # offset 0: VBA events must fall inside their host stage (DESIGN rule 9)
            f = _vba_file(tmp_path / f"vba-{mode}-{k}.jsonl", k, mode, base_ns=t + 100)
            evs, footer = trace.read_vba_events(f)
            assert runner.check_vba_footer(footer) == []
            events[(k, mode)] = evs
            stages.append({"name": f"profile_trace_{mode}", "start": t, "end": t + 50_000, "status": "ok",
                           "attrs": {"repeat": k}, "pass": k, "mode": mode})
            t += 100_000
    for name in ("collect_trace", "close_excel", "verify_output"):
        stages.append({"name": name, "start": t, "end": t + 500, "status": "ok", "attrs": {}})
        t += 1_000
    return stages, events


HEADER = {"run_id": "r", "created_utc": "2026-09-25T00:00:00Z", "tool_version": "0.1.0",
          "clock": {"host": "perf_counter_ns", "vba": "MicroTimer/QPC", "vba_offset_ns": 0, "vba_offset_uncertainty_ns": 1},
          "source": {"kind": "synthetic", "label": "synthetic:runner-merge"}}


def _write_both(tmp_path, stages, events, ident, meta, names_mode, expect):
    def merger(mode):
        def merge(w, st):
            if st.get("mode") == mode:
                runner.merge_vba_events(w, events[(st["pass"], mode)], ident=ident, meta=meta,
                                        attr_ok=runner.attr_value_ok)
        return merge

    on, off = tmp_path / "trace-on.jsonl", tmp_path / "trace-off.jsonl"
    hdr = dict(HEADER, redaction={"names": names_mode})
    runner.write_trace(on, dict(hdr, mode="trace-on", expect=expect), stages, max_events=10_000,
                       include=lambda s: True, merge=merger("on"))
    runner.write_trace(off, dict(hdr, mode="trace-off"), stages, max_events=10_000,
                       include=lambda s: s.get("mode") == "off", merge=merger("off"))
    return on, off


def test_merge_fake_vba_events_hashed(tmp_path):
    trace = pytest.importorskip("xlsprint.trace")
    ident = runner.make_ident("hashed", "s4lt")
    h = lambda s: trace.redact_name(s, "s4lt")  # noqa: E731
    stages, events = _stages_and_events(tmp_path, trace)
    plan = [("full",), ("recalc",), ("sheet", SHEET), ("name", "Calc", "$A$1", "Totals", "Calc"), ("macro", "Module1.Refresh")]
    passes = [(k, m, e) for (k, m), e in events.items()]
    assert runner.check_expected(passes, plan) == []
    critical, other = runner.span_failures(passes)
    assert critical == [] and len(other) == 2 and other[0]["kind"] == "calc.name"

    meta = {"R0001": {"name": h("Calc") + "!$A$1:$B$2", "sheet": "Calc", "note": "auto 1/1"},
            "N0001": {"name": h("Calc") + "!" + h("Totals"), "sheet": "Calc"}}
    expect = {"calc.sheet": [h(SHEET)], "vba.proc": [h("Module1.Refresh")]}
    on, off = _write_both(tmp_path, stages, events, ident, meta, "hashed", expect)
    for p in (on, off):
        res = trace.validate(p)
        assert res.ok, (p.name, res.errors)

    text = on.read_text(encoding="utf-8")
    for leak in ("Caf", "Calc!", '"Calc"', "Totals", "Refresh", "Inner", "R0001", "N0001", '"mid"'):
        assert leak not in text, leak
    hdr, evs, footer = trace.read_trace(on)
    by_id = {e["id"]: e for e in evs if e["type"] in ("B", "M")}
    rng = next(e for e in evs if e.get("kind") == "calc.range")
    assert rng["name"] == h("Calc") + "!$A$1:$B$2"
    assert rng["attrs"]["sheet"] == h("Calc") and rng["attrs"]["note"] == "auto 1/1"
    nm = next(e for e in evs if e.get("kind") == "calc.name")
    assert nm["name"] == h("Calc") + "!" + h("Totals") and nm["attrs"]["sheet"] == h("Calc")
    rp = [e for e in evs if e.get("kind") == "run.pass"]
    assert len(rp) == 2 and all(by_id[e["parent"]]["name"] == "profile_trace_on" for e in rp)
    inner = next(e for e in evs if e.get("kind") == "vba.proc" and e["depth"] == 3)
    assert inner["name"] == h("Inner") and by_id[inner["parent"]]["kind"] == "vba.proc"
    marker = next(e for e in evs if e["type"] == "M")
    assert marker["parent"] == inner["id"] and marker["name"] == h("mid")
    assert len({e["id"] for e in evs if e["type"] in ("B", "M")}) == len(by_id)

    _, off_evs, _ = trace.read_trace(off)
    assert {e["kind"] for e in off_evs if e["type"] == "B"} == {"host.stage", "run.pass", "calc.full", "calc.recalc"}


def test_merge_fake_vba_events_clear_rule10(tmp_path):
    trace = pytest.importorskip("xlsprint.trace")
    if trace.check_name(SHEET) is None:
        pytest.skip("trace.check_name does not implement NAME_RE (DESIGN rule 10) yet")
    ident = runner.make_ident("clear", "s4lt")
    stages, events = _stages_and_events(tmp_path, trace)
    expect = {"calc.sheet": [ident(SHEET)], "vba.proc": ["Module1.Refresh"]}
    on, off = _write_both(tmp_path, stages, events, ident, META_CLEAR, "clear", expect)
    for p in (on, off):
        res = trace.validate(p)
        assert res.ok, (p.name, res.errors)
    _, evs, _ = trace.read_trace(on)
    sheet_span = next(e for e in evs if e.get("kind") == "calc.sheet")
    assert sheet_span["name"] == trace.redact_name(SHEET, "s4lt")  # hashed although --names clear
    rng = next(e for e in evs if e.get("kind") == "calc.range")
    assert rng["name"] == "Calc!$A$1:$B$2" and rng["attrs"]["sheet"] == "Calc"
    assert {e["name"] for e in evs if e.get("kind") == "vba.proc"} == {"Module1.Refresh", "Inner"}


def test_merge_unknown_key_is_hashed(tmp_path):
    trace = pytest.importorskip("xlsprint.trace")
    evs, _ = trace.read_vba_events(_vba_file(tmp_path / "v.jsonl", 1, "on", 10_000))
    w = trace.TraceWriter(tmp_path / "t.jsonl", dict(HEADER, mode="trace-on", redaction={"names": "clear"}))
    sid = w.begin("host.stage", "profile_trace_on", ns=1)
    runner.merge_vba_events(w, evs, ident=runner.make_ident("hashed", "k"), meta={})
    w.end(sid, ns=2)
    w.close()
    _, out, _ = trace.read_trace(tmp_path / "t.jsonl")
    rng = next(e for e in out if e.get("kind") == "calc.range")
    assert rng["name"] == trace.redact_name("R0001", "k") and "sheet" not in rng["attrs"]


def test_check_expected_reports_missing():
    evs = [{"type": "B", "id": 1, "kind": "run.pass", "name": "pass"},
           {"type": "B", "id": 2, "kind": "calc.full", "name": "workbook"}]
    probs = runner.check_expected([(1, "on", evs), (1, "off", evs + [{"type": "B", "id": 3, "kind": "calc.sheet", "name": "S"}])],
                                  [("sheet", "Secret Sheet"), ("macro", "SecretMacro")])
    joined = " | ".join(probs)
    assert "on pass 1: no calc.recalc" in joined
    assert "1 planned sheet(s) have no calc.sheet" in joined and "macro #1" in joined
    assert "off pass 1: unexpected kinds ['calc.sheet']" in joined
    assert "Secret" not in joined  # identifiers are never echoed


# --------------------------------------------------------------------------
# Static lint of the VBA module


def _bas_lines():
    raw = BAS.read_bytes()
    raw.decode("ascii")  # the VBE imports in the ANSI code page: ASCII only
    text = raw.decode("ascii").replace("\r\n", "\n")
    joined, buf = [], ""
    for line in text.split("\n"):
        code = _strip_comment(line)
        if code.rstrip().endswith(" _"):
            buf += code.rstrip()[:-2] + " "
            continue
        joined.append((buf + code).strip())
        buf = ""
    return joined


def _strip_comment(line):
    out, in_str = [], False
    for ch in line:
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            break
        out.append(ch)
    return "".join(out)


def _procs():
    """{name: [lines]} for every Sub/Function body."""
    procs, cur = {}, None
    for ln in _bas_lines():
        m = re.match(r"^(?:Public |Private )?(?:Sub|Function) (\w+)", ln)
        if m:
            cur = m.group(1)
            procs[cur] = []
        elif re.match(r"^End (Sub|Function)$", ln):
            cur = None
        elif cur:
            procs[cur].append(ln)
    return procs


def test_bas_header_and_option_explicit():
    lines = BAS.read_text(encoding="ascii").splitlines()
    assert lines[0] == 'Attribute VB_Name = "XLSprintTimer"'
    code = [l for l in _bas_lines() if l]
    assert "Option Explicit" in code
    first_code = next(l for l in code if not l.startswith("Attribute"))
    assert first_code == "Option Explicit"


def test_bas_declares_ptrsafe_under_vba7():
    state, saw_vba7 = None, False
    for ln in _bas_lines():
        if ln.startswith("#If "):
            state = "vba7" if ln == "#If VBA7 Then" else "other"
            saw_vba7 = saw_vba7 or state == "vba7"
        elif ln == "#Else":
            state = "else"
        elif ln == "#End If":
            state = None
        elif re.search(r"\bDeclare\b", ln):
            assert state in ("vba7", "else"), f"Declare outside #If VBA7: {ln}"
            if state == "vba7":
                assert "Declare PtrSafe" in ln, ln
            assert ln.startswith("Private "), ln
    assert saw_vba7


def test_bas_blocks_balanced():
    counts = {k: 0 for k in ("if", "sub", "function", "select", "for", "with", "do")}
    for ln in _bas_lines():
        if not ln or ln.startswith("#") or ln.startswith("Attribute"):
            continue
        if re.match(r"^If .* Then$", ln):
            counts["if"] += 1
        elif ln == "End If":
            counts["if"] -= 1
        if re.match(r"^(Public |Private )?Sub \w+", ln):
            counts["sub"] += 1
        elif ln == "End Sub":
            counts["sub"] -= 1
        if re.match(r"^(Public |Private )?Function \w+", ln):
            counts["function"] += 1
        elif ln == "End Function":
            counts["function"] -= 1
        if ln.startswith("Select Case "):
            counts["select"] += 1
        elif ln == "End Select":
            counts["select"] -= 1
        if re.match(r"^For (Each )?\w+", ln):
            counts["for"] += 1
        elif re.match(r"^Next\b", ln):
            counts["for"] -= 1
        if ln.startswith("With "):
            counts["with"] += 1
        elif ln == "End With":
            counts["with"] -= 1
        if re.match(r"^Do\b", ln):
            counts["do"] += 1
        elif re.match(r"^Loop\b", ln):
            counts["do"] -= 1
        assert all(v >= 0 for v in counts.values()), ln
    assert counts == {k: 0 for k in counts}


def test_bas_is_non_interactive():
    code = "\n".join(_bas_lines())
    for word in ("MsgBox", "InputBox", "Stop", "SendKeys", "Debug.Print", "GetOpenFilename", "Shell"):
        assert not re.search(r"\b%s\b" % re.escape(word), code), word
    assert not re.search(r"^End$", code, re.M)  # a bare End would reset all module state


def test_bas_goto_labels_exist():
    for name, body in _procs().items():
        labels = {m.group(1) for l in body for m in [re.match(r"^(\w+):$", l)] if m}
        for l in body:
            for m in re.finditer(r"\b(?:GoTo|Resume) (\w+)", l):
                tgt = m.group(1)
                if tgt in ("0", "Next"):
                    continue
                assert tgt in labels, f"{name}: missing label {tgt}"


def test_bas_no_case_insensitive_name_clashes():
    code = "\n".join(_bas_lines())
    names = re.findall(r"^(?:Public |Private )?(?:Sub|Function|Const|Declare PtrSafe Function|Declare Function) (\w+)", code, re.M)
    names = [n for n in names]
    # the #Else declares repeat the VBA7 ones; count unique spellings only
    lowered = {}
    for n in set(names):
        lowered.setdefault(n.lower(), set()).add(n)
    clashes = {k: v for k, v in lowered.items() if len(v) > 1}
    assert not clashes
    procs = [n for n in re.findall(r"^(?:Public |Private )?(?:Sub|Function) (\w+)", code, re.M)]
    consts = re.findall(r"^(?:Public |Private )?Const (\w+)", code, re.M)
    assert not ({p.lower() for p in procs} & {c.lower() for c in consts})


def test_bas_local_names_do_not_shadow_attribute_builders():
    code = "\n".join(_bas_lines())
    builders = {
        name.lower()
        for name in re.findall(r"^(?:Public |Private )?Function (Attr[SNB])\b", code, re.M)
    }
    for proc, body in _procs().items():
        for line in body:
            if not re.match(r"^Dim\s+", line, re.I):
                continue
            declarations = re.sub(r"^Dim\s+", "", line, flags=re.I)
            local_names = re.findall(
                r"(?:^|,)\s*([A-Za-z_]\w*)(?:\s*\([^)]*\))?\s+As\b",
                declarations,
                re.I,
            )
            conflicts = builders & {name.lower() for name in local_names}
            assert not conflicts, f"{proc}: local declaration shadows {sorted(conflicts)}"


@pytest.mark.parametrize("hresult", [-2147418111, 0x80010001, -2147417846, 0x8001010A])
def test_excel_busy_com_hresult_is_retryable(hresult):
    exc = Exception("Excel is busy")
    exc.hresult = hresult
    assert runner._is_excel_busy_rejection(exc)


@pytest.mark.parametrize("hresult", [None, 0, -2147467259, "not-an-hresult"])
def test_non_busy_com_hresult_is_not_retryable(hresult):
    exc = Exception("not an Excel busy rejection")
    if hresult is not None:
        exc.hresult = hresult
    assert not runner._is_excel_busy_rejection(exc)


def test_runner_retries_only_excel_busy_com_rejections():
    src = Path(runner.__file__).read_text(encoding="utf-8")
    body = src.split("    def run_vba(proc: str, *args):", 1)[1].split("    def probe_clock", 1)[0]
    assert "_MAX_EXCEL_BUSY_RETRIES" in body
    assert "_is_excel_busy_rejection(exc)" in body
    assert "pythoncom.PumpWaitingMessages()" in body


def test_bas_public_api_used_by_runner():
    code = "\n".join(_bas_lines())
    src = Path(runner.__file__).read_text(encoding="utf-8")
    called = set(re.findall(r'run_vba\("(\w+)"', src))
    assert called == {"XSP_Init", "XSP_RunPass", "XSP_ClockProbe", "XSP_Version", "XSP_Bitness",
                      "XSP_SaveSettings"}
    # restoring settings would force a recalc in the throwaway instance; kept but not called
    assert "XSP_RestoreSettings" not in src.split("def _count_kinds")[0].replace("# ", "")
    assert re.search(r"^Public Function XSP_RestoreSettings\(", code, re.M)
    for proc in called:
        assert re.search(r"^Public Function %s\(" % proc, code, re.M), proc
    for proc in ("XSP_Begin", "XSP_End", "XSP_Marker", "XSP_Flush", "XSP_MicroTimer"):
        assert re.search(r"^Public (Sub|Function) %s\(" % proc, code, re.M), proc
    assert 'XSP_VBA_VERSION As String = "%s"' % runner.VBA_VERSION in code
    assert re.search(r"^Public Const XSP_MAX_EVENTS As Long = 100000$", code, re.M)
    # Init takes (outPath, mode, firstId, parentId, parentDepth, maxEvents)
    init = re.search(r"^Public Function XSP_Init\((.*?)\) As String", code, re.M).group(1)
    assert [p.split(" As ")[0].split()[-1] for p in init.split(",")] == [
        "outPath", "mode", "firstId", "parentId", "parentDepth", "maxEvents"]
    run_pass = re.search(r"^Public Function XSP_RunPass\((.*?)\) As String", code, re.M).group(1)
    assert [p.split(" As ")[0].split()[-1] for p in run_pass.split(",")] == ["passId", "mode", "planPath"]
    assert "plan = ReadTextFile(planPath)" in code
    assert 'run_vba("XSP_RunPass", k, mode, str(plan_path))' in src


def test_bas_begin_accepts_vba_proc_only():
    body = _procs()["XSP_Begin"]
    assert 'If kind = "vba.proc" Then' in body
    assert not any("calc." in l for l in body)


def test_bas_has_no_per_cell_array_scan():
    code = "\n".join(_bas_lines())
    for word in ("CurrentArray", "HasArray", "SpecialCells", "For Each c"):
        assert word not in code, word
    body = _procs()["StepRange"]
    assert "rng.Calculate" in body and "sid = SpanBegin(kind, key, spanAttributes)" in body


def test_bas_attr_allow_list_matches_trace():
    trace = pytest.importorskip("xlsprint.trace")
    body = " ".join(_procs()["KeyAllowed"])
    keys = set(re.findall(r'"(\w+)"', body))
    assert keys == set(trace.ALLOWED_ATTR_KEYS)
    # every attr key the module emits is allow-listed
    code = "\n".join(_bas_lines())
    used = set(re.findall(r'Attr[SNB]\((?:""|\w+), "(\w+)"', code))
    assert used and used <= keys


def test_runner_attr_filter_follows_contract():
    pytest.importorskip("xlsprint.trace")
    assert runner.attr_value_ok("caf\u00e9") and runner.attr_value_ok("Mod\u00e8le") and runner.attr_value_ok("h:0123456789")
    for bad in ("a\n", 'say "hi"', "=SUM(A1)", "x" * 257):
        assert not runner.attr_value_ok(bad)


def _vba_int(tok):
    tok = tok.strip()
    m = re.match(r"^&H([0-9A-Fa-f]+)&?$", tok)
    return int(m.group(1), 16) if m else int(tok)


def test_bas_value_pattern_within_contract():
    """ValueAllowed equals the contract over ASCII and is a subset of Unicode \\w beyond it."""
    trace = pytest.importorskip("xlsprint.trace")
    body = _procs()["ValueAllowed"]
    allowed = set()
    for l in body:
        m = re.match(r"^Case ([0-9A-Fa-f&H ,To]+)$", l)
        if not m:
            continue
        for part in m.group(1).split(","):
            if " To " in part:
                a, b = part.split(" To ")
                allowed.update(range(_vba_int(a), _vba_int(b) + 1))
            else:
                allowed.add(_vba_int(part))
    for c in range(0, 128):
        assert (c in allowed) == bool(trace.ATTR_VALUE_RE.fullmatch(chr(c))), (c, chr(c))
    wider = [hex(c) for c in allowed if c >= 128 and not trace.ATTR_VALUE_RE.fullmatch(chr(c))]
    assert not wider
    for ch in "\u00e9\u00c9\u00fc\u00df\u03a9\u0416\u05d0\u0628\u0e01\u3042\u30a2\u65e5\ud55c":
        assert ord(ch) in allowed, ch
    for ch in "\u00a0\u00d7\u00f7\u2019\u3000":  # NBSP, x, /, curly quote, ideographic space
        assert ord(ch) not in allowed
    assert not any(0xD800 <= c <= 0xDFFF for c in allowed)  # surrogates (emoji) drop the attr
    assert re.search(r"XSP_MAX_ATTR_LEN As Long = %d$" % trace.ATTR_MAX_LEN, "\n".join(_bas_lines()), re.M)


def _vba_literal(lit):
    assert lit[0] == lit[-1] == '"'
    return lit[1:-1].replace('""', '"')


def _escape_table():
    """Parse JsonStr's Select Case into {code: replacement} plus the pass-through range."""
    body = _procs()["JsonStr"]
    table, passthrough, default_hex = {}, None, False
    for l in body:
        m = re.match(r'^Case (\d+): r = r & ("(?:[^"]|"")*")$', l)
        if m:
            table[int(m.group(1))] = _vba_literal(m.group(2))
            continue
        m = re.match(r"^Case (\d+) To (\d+): r = r & ch$", l)
        if m:
            passthrough = (int(m.group(1)), int(m.group(2)))
            continue
        if l == 'Case Else: r = r & "\\u" & Right$("000" & Hex$(c), 4)':
            default_hex = True
    assert 'JsonStr = """" & r & """"' in body
    assert "c = AscW(ch) And &HFFFF&" in body
    return table, passthrough, default_hex


def _port_json_str(s):
    table, (lo, hi), default_hex = _escape_table()
    assert default_hex
    out = []
    # VBA strings are UTF-16 code units; AscW sees each unit, surrogates included.
    units = s.encode("utf-16-le")
    for i in range(0, len(units), 2):
        c = units[i] | (units[i + 1] << 8)
        if c in table:
            out.append(table[c])
        elif lo <= c <= hi:
            out.append(chr(c))
        else:
            out.append("\\u" + ("000" + "%X" % c)[-4:])
    return '"' + "".join(out) + '"'


def test_bas_json_escape_table():
    table, passthrough, default_hex = _escape_table()
    assert table == {34: '\\"', 92: "\\\\", 8: "\\b", 9: "\\t", 10: "\\n", 12: "\\f", 13: "\\r"}
    assert passthrough == (32, 126) and default_hex


@pytest.mark.parametrize("s", [
    'plain', 'quote " inside', "back\\slash", "cr\rlf\n tab\t", "".join(chr(c) for c in range(32)),
    "del\x7f", "Caf\u00e9", "\u65e5\u672c", "emoji \U0001F600", "'apos' $A$1:B2 !#@", "",
])
def test_bas_json_escape_roundtrip(s):
    lit = _port_json_str(s)
    assert all(32 <= ord(c) <= 126 for c in lit)
    assert json.loads(lit) == s


def test_spill_areas_are_not_extended():
    """Only CSE array_areas widen a step; calculating a spill's anchor recalculates the spill."""
    info = runner.adapt_plan_info({"sheets": [
        {"sheet": "S", "formula_cells": 4, "formula_area": "A1:B4", "used_range": "A1:D9",
         "array_areas": ["B4:B5"], "spill_areas": ["A4:A9", "C1:D9"]}]})
    assert info["sheets"][0]["array_areas"] == ["B4:B5"]
    vba, _ = runner.compile_plan([("range", "S", "$A$1:$B$4", "auto 1/1")], info, lambda s: s)
    assert vba == [("range", "S", "$A$1:$B$4,$B$5", "R0001")]


def test_on_and_off_traces_share_run_id(tmp_path):
    trace = pytest.importorskip("xlsprint.trace")
    stages, events = _stages_and_events(tmp_path, trace)
    ident = runner.make_ident("hashed", "k")
    meta = {"R0001": {"name": "h:x!$A$1", "sheet": "S"}, "N0001": {"name": "h:y", "sheet": "S"}}
    on, off = _write_both(tmp_path, stages, events, ident, meta, "hashed", {})
    assert trace.read_trace(on)[0]["run_id"] == trace.read_trace(off)[0]["run_id"] == HEADER["run_id"]
