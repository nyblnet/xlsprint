import copy
import json
import re
from dataclasses import dataclass, field
from types import SimpleNamespace

import pytest

from xlsprint import report

SECTION_IDS = ["run", "legend", "route", "timeline", "drilldown", "hotspots", "phases",
               "overhead", "formulas", "limits"]
EVIL = '<script>x</script>&"'
SENTINEL = "SUM(SENTINEL_FORMULA_TEXT)"


def _summ(med, n=5, lo=None, hi=None):
    return {"n": n, "median": med, "min": lo if lo is not None else med * 0.9,
            "max": hi if hi is not None else med * 1.2, "p25": med * 0.95, "p75": med * 1.1}


def make_model(kind="real", ok=True, off=True):
    rows = [
        {"id": 1, "kind": "host.stage", "name": "open_workbook", "pass": 0, "depth": 0,
         "start_rel_ns": 0, "dur_ns": 2_000_000_000, "status": "ok"},
        {"id": 2, "kind": "host.stage", "name": "profile_trace_on", "pass": 0, "depth": 0,
         "start_rel_ns": 2_000_000_000, "dur_ns": 3_000_000_000, "status": "ok"},
    ]
    nid = 3
    for p in (1, 2):
        base = 2_000_000_000 + (p - 1) * 1_000_000_000
        rows.append({"id": nid, "kind": "run.pass", "name": f"pass{p}", "pass": p, "depth": 1,
                     "start_rel_ns": base, "dur_ns": 900_000_000, "status": "ok"})
        rows.append({"id": nid + 1, "kind": "calc.full", "name": "full", "pass": p, "depth": 2,
                     "start_rel_ns": base + 1000, "dur_ns": 400_000_000, "status": "ok"})
        rows.append({"id": nid + 2, "kind": "calc.sheet", "name": EVIL, "pass": p, "depth": 2,
                     "start_rel_ns": base + 500_000_000, "dur_ns": 90_000_000, "status": "ok"})
        nid += 3
    return {
        "schema": report.REPORT_SCHEMA,
        "tool_version": "0.1.0",
        "meta": {
            "label": "Budget model", "source_kind": kind, "run_id": "abc123", "mode": "trace-on",
            "created_utc": "2026-09-25T12:00:00Z", "trace_tool_version": "0.1.0",
            "host": {"os": "Windows-10", "python": "3.12.1", "excel_version": "16.0",
                     "excel_build": "17928", "bitness": 64, "threads": 8, "calc_mode_original": "automatic"},
            "clock": {"host": "perf_counter_ns", "vba": "MicroTimer/QPC", "vba_offset_ns": 123,
                      "vba_offset_uncertainty_ns": 4000},
            "redaction_names": "clear", "repeats_on": 2, "repeats_off": 2 if off else None, "has_off": off,
        },
        "validation": {"ok": ok, "errors": [] if ok else ["trace-on: [2] E without matching B (id 7)"],
                       "warnings": [], "allow_invalid": False},
        "host_stages": [
            {"name": "open_workbook", "start_rel_ns": 0, "dur_ns": 2_000_000_000, "status": "ok"},
            {"name": "profile_trace_on", "start_rel_ns": 2_000_000_000, "dur_ns": 3_000_000_000, "status": "ok"},
        ],
        "passes": [{"pass": 1, "dur_ns": 900_000_000, "by_kind_union_ns": {"calc.full": 400_000_000}},
                   {"pass": 2, "dur_ns": 910_000_000, "by_kind_union_ns": {"calc.full": 410_000_000}}],
        "timeline": {"rows": rows, "truncated": 17, "max_rows": 2000},
        "drilldown": {
            "workbook": {"full": _summ(400_000_000), "recalc": _summ(50_000_000),
                         "fullrebuild": _summ(700_000_000), "volatility_ratio": 0.125},
            "sheets": [{"name": EVIL, "recalc": _summ(90_000_000), "share_of_recalc_sum": 0.6,
                        "formula": SENTINEL,
                        "children": [
                            {"kind": "calc.range", "name": "auto1", "summary": _summ(10_000_000),
                             "address": "A1:B10"},
                            {"kind": "calc.group", "name": "G0001", "summary": _summ(3_000_000),
                             "address": "C1:C100", "formula_text": SENTINEL,
                             "measurement": "measured (isolated Range.Calculate)"},
                        ]},
                       {"name": "Sheet2", "recalc": _summ(60_000_000), "share_of_recalc_sum": 0.4,
                        "children": []}],
        },
        "phases": [{"kind": "calc.full", "count": 2, "union_ns_median_per_pass": 400_000_000, "share_of_pass": 0.44},
                   {"kind": "calc.sheet", "count": 4, "union_ns_median_per_pass": 150_000_000, "share_of_pass": 0.16},
                   {"kind": "calc.range", "count": 4, "union_ns_median_per_pass": 20_000_000, "share_of_pass": 0.02}],
        "vba_procs": [{"kind": "vba.proc", "name": "RunModel", "summary": _summ(5_000_000),
                       "self_summary": _summ(1_000_000),
                       "children": [{"kind": "vba.proc", "name": "Inner", "summary": _summ(4_000_000),
                                     "self_summary": _summ(4_000_000), "children": []}]}],
        "hotspots": [{"kind": "calc.sheet", "name": EVIL, "sheet": EVIL, "median_ns": 90_000_000,
                      "median_self_ns": 70_000_000, "n": 2, "min_ns": 85_000_000, "max_ns": 95_000_000,
                      "measurement": "measured"},
                     {"kind": "calc.group", "name": "G0001", "sheet": "Sheet2", "median_ns": 3_000_000,
                      "median_self_ns": 3_000_000, "n": 2, "min_ns": 2_900_000, "max_ns": 3_100_000,
                      "measurement": "measured", "note": "measured (isolated Range.Calculate)"}],
        "overhead": ({"available": True, "reason": None, "rows": [
            {"kind": "run.pass", "name": "pass", "median_on_ns": 900_000_000, "median_off_ns": 899_000_000,
             "diff_ns": 1_000_000, "diff_min_ns": -2_000_000, "diff_max_ns": 3_000_000,
             "iqr_on_ns": 4_000_000, "iqr_off_ns": 5_000_000,
             "resolvable": False, "pairs": 2}]} if off else
            {"available": False, "rows": [], "reason": "No trace-off run was supplied."}),
        "formulas": {"schema": "xlsprint.formulas/1", "workbook_sha256_prefix": "ab12cd34ef",
                     "sheets": [{"sheet": EVIL, "formula_cells": 1200, "used_range": "A1:F2000"}],
                     "groups": [{"group": "G0001", "fingerprint": "0123456789abcdef", "sheet": EVIL,
                                 "cells": 1000, "areas": ["B2:B1001"], "functions": ["SUM", "OFFSET"],
                                 "udfs": 1, "volatile": True, "volatile_functions": ["OFFSET"],
                                 "single_threaded": False, "single_threaded_functions": [],
                                 "array": False, "dynamic_array": False, "cross_sheet": True,
                                 "external_ref": False, "whole_column_ref": False, "formula": SENTINEL}],
                     "groups_total": 1, "names_count": 1,
                     "limitations": ["Excel 4 macro sheets are not inspected."],
                     "totals": {"formula_cells": 1200, "groups": 1, "volatile_cells": 1000,
                                "udf_cells": 1000, "single_thread_cells": 0}},
        "section_errors": {},
        "secret_unknown_field": SENTINEL,
    }


def test_new_analysis_fields_rendered():
    out = report.render_html(make_model())
    assert "Full rebuild" in out
    assert "measured (isolated Range.Calculate)" in out
    assert "spread (IQR)" in out
    assert "Excel 4 macro sheets are not inspected." in out
    assert "[2] E without matching B" not in out  # valid model: no error list
    bad = report.render_html(make_model(ok=False))
    assert "[2] E without matching B (id 7)" in bad


def test_sections_present():
    out = report.render_html(make_model())
    for sid in SECTION_IDS:
        assert f'id="{sid}"' in out, sid
    assert out.startswith("<!DOCTYPE html>")


def test_render_is_deterministic():
    m = make_model()
    assert report.render_html(m) == report.render_html(copy.deepcopy(m))


def test_synthetic_banner_toggles():
    assert "SYNTHETIC — not a real-workbook measurement" in report.render_html(make_model("synthetic"))
    assert "SYNTHETIC" not in report.render_html(make_model("real"))


def test_invalid_banner_withholds_timing_unless_allowed():
    m = make_model(ok=False)
    out = report.render_html(m)
    assert "INVALID TRACE" in out
    assert "[2] E without matching B (id 7)" in out
    for sid in ("route", "timeline", "drilldown", "hotspots", "phases", "overhead"):
        assert f'id="{sid}"' not in out
    assert 'id="limits"' in out and 'id="formulas"' in out

    m["validation"]["allow_invalid"] = True
    out = report.render_html(m)
    assert "INVALID TRACE" in out
    assert 'id="timeline"' in out
    assert "INVALID TRACE" not in report.render_html(make_model())


def test_html_escaping():
    out = report.render_html(make_model())
    assert "<script>x</script>" not in out
    assert "&lt;script&gt;x&lt;/script&gt;&amp;&quot;" in out
    # Only our own two script elements exist.
    assert out.count("<script") == 2
    # The embedded JSON cannot close its script element early.
    blob = re.search(r'<script type="application/json" id="xlsprint-model">(.*?)</script>', out, re.S).group(1)
    assert "</" not in blob and "<" not in blob
    assert json.loads(blob)["meta"]["label"] == "Budget model"


def test_no_external_urls():
    out = report.render_html(make_model())
    assert not re.search(r"https?://", out)
    assert not re.search(r"""(src|href)\s*=\s*["']?\s*(//|[a-z]+:)""", out, re.I)
    assert "@import" not in out and "url(http" not in out


def test_every_number_has_a_class():
    out = report.render_html(make_model())
    nums = re.findall(r'<(?:span|tspan)\b[^>]*\bclass="num"[^>]*>', out)
    assert len(nums) > 50
    for tag in nums:
        m = re.search(r'data-class="([^"]+)"', tag)
        assert m and m.group(1) in {"measured", "derived", "static", "na"}, tag
    # Derived values show their formula.
    assert report.F_VOLATILITY in out
    assert "median recalc ÷ median full" in out


def test_unknown_fields_never_rendered():
    out = report.render_html(make_model())
    assert SENTINEL not in out
    assert "SENTINEL" not in out


def test_group_caveats_and_no_per_formula_time():
    out = report.render_html(make_model())
    assert report.GROUP_CAVEAT in out
    assert "measured (isolated Range.Calculate)" in out
    # Group rows never carry a per-cell division or a sum across groups.
    assert "per cell" not in out.lower() and "sum of groups" not in out.lower()


def test_timeline_truncation_and_pass_selector():
    out = report.render_html(make_model())
    assert "Timeline truncated" in out
    assert 'id="tl-pass-select"' in out
    assert 'data-pass="1"' in out and 'data-pass="2" hidden' in out


def test_overhead_not_resolvable_and_unavailable():
    assert "not resolvable (below run-to-run noise)" in report.render_html(make_model())
    out = report.render_html(make_model(off=False))
    assert "overhead is unavailable" in out or "No trace-off run" in out


def test_phases_note_unions():
    out = report.render_html(make_model())
    assert "unions, not sums" in out
    assert "RunModel" in out and "Inner" in out


def test_limits_always_present():
    out = report.render_html({})
    assert 'id="limits"' in out
    for phrase in ("no per-formula execution timer", "ScreenUpdating", "Multithread",
                   "medians", "Clock alignment", "PivotTables", "Application.Volatile"):
        assert phrase in out, phrase


# --------------------------------------------------------------------------
# build_model / render with a stand-in analysis module
# --------------------------------------------------------------------------

@dataclass
class FakeSpan:
    id: int
    parent: int
    pass_: int
    depth: int
    kind: str
    name: str
    clock: str
    start_ns: int
    end_ns: int
    status: str = "ok"
    attrs: dict = field(default_factory=dict)
    children: list = field(default_factory=list)

    @property
    def dur_ns(self):
        return self.end_ns - self.start_ns


def _fake_run(mode="trace-on", ok=True, kind="synthetic"):
    spans = {
        1: FakeSpan(1, 0, 1, 0, "run.pass", "p1", "vba", 0, 100, children=[2, 3]),
        2: FakeSpan(2, 1, 1, 1, "calc.full", "full", "vba", 0, 50),
        3: FakeSpan(3, 1, 1, 1, "vba.proc", "Macro1", "vba", 50, 90, children=[4]),
        4: FakeSpan(4, 3, 1, 2, "vba.proc", "Inner", "vba", 60, 80),
    }
    header = {"mode": mode, "run_id": "r1", "source": {"kind": kind, "label": "synthetic:test"},
              "host": {"threads": 4}, "clock": {"vba_offset_ns": 0, "vba_offset_uncertainty_ns": 10}}
    return SimpleNamespace(header=header, spans=spans, markers=[], mode=mode, source_kind=kind,
                           validation=SimpleNamespace(ok=ok, errors=[] if ok else ["unbalanced"], warnings=[]))


def _fake_analysis(runs=None):
    def summarize(vals):
        vals = sorted(vals)
        return {"n": len(vals), "median": vals[len(vals) // 2], "min": vals[0], "max": vals[-1],
                "p25": vals[0], "p75": vals[-1]}

    def self_time_ns(run, sid):
        sp = run.spans[sid]
        return sp.dur_ns - sum(run.spans[c].dur_ns for c in sp.children)

    return SimpleNamespace(
        summarize=summarize, self_time_ns=self_time_ns,
        load_run=lambda p: runs[p.name],
        host_stages=lambda r: [{"name": "warmup", "start_rel_ns": 0, "dur_ns": 5, "status": "ok"}],
        timeline=lambda r, max_rows=2000: [
            {"id": 1, "kind": "run.pass", "name": "p1", "pass": 1, "depth": 0, "start_rel_ns": 0,
             "dur_ns": 100, "status": "ok"}, {"truncated": 3}],
        passes=lambda r: [{"pass": 1, "dur_ns": 100, "by_kind_union_ns": {"calc.full": 50}}],
        drilldown=lambda r: {"workbook": {"full": summarize([50]), "recalc": summarize([5]),
                                          "volatility_ratio": 0.1}, "sheets": []},
        phase_breakdown=lambda r: [{"kind": "calc.full", "count": 1, "union_ns_median_per_pass": 50,
                                    "share_of_pass": 0.5}],
        hotspots=lambda r, top=25: [],
        overhead=lambda on, off: [{"kind": "run.pass", "name": "pass", "median_on_ns": 100,
                                   "median_off_ns": 90, "diff_ns": 10, "diff_min_ns": 5, "diff_max_ns": 15,
                                   "resolvable": True, "pairs": 1}],
    )


def test_build_model_with_fake_analysis(monkeypatch):
    A = _fake_analysis()
    monkeypatch.setattr(report, "_analysis", lambda: A)
    model = report.build_model(_fake_run(), _fake_run("trace-off"), None)
    assert model["timeline"]["truncated"] == 3
    assert all("truncated" not in r for r in model["timeline"]["rows"])
    assert model["meta"]["source_kind"] == "synthetic"
    assert model["meta"]["repeats_on"] == 1
    assert model["overhead"]["available"] is True
    procs = model["vba_procs"]
    assert [p["name"] for p in procs] == ["Macro1"]
    assert procs[0]["children"][0]["name"] == "Inner"
    assert procs[0]["self_summary"]["median"] == 20
    out = report.render_html(model)
    assert "SYNTHETIC" in out
    json.dumps(model)  # model is JSON-serialisable

    no_off = report.build_model(_fake_run(), None, None)
    assert no_off["overhead"]["available"] is False


def test_render_refuses_invalid_unless_allowed(monkeypatch, tmp_path):
    runs = {"on.jsonl": _fake_run(ok=False), "off.jsonl": _fake_run("trace-off")}
    A = _fake_analysis(runs)
    monkeypatch.setattr(report, "_analysis", lambda: A)
    out = tmp_path / "report.html"
    with pytest.raises(report.ReportError, match="invalid"):
        report.render([tmp_path / "on.jsonl", tmp_path / "off.jsonl"], None, out)
    assert not out.exists()
    report.render([tmp_path / "on.jsonl", tmp_path / "off.jsonl"], None, out, allow_invalid=True)
    text = out.read_text(encoding="utf-8")
    assert "INVALID TRACE" in text and "unbalanced" in text and 'id="timeline"' in text


def test_render_requires_one_trace_on(monkeypatch, tmp_path):
    runs = {"off.jsonl": _fake_run("trace-off")}
    monkeypatch.setattr(report, "_analysis", lambda: _fake_analysis(runs))
    with pytest.raises(report.ReportError, match="trace-on"):
        report.render([tmp_path / "off.jsonl"], None, tmp_path / "r.html")


def test_formulas_projection_drops_unsafe_strings():
    fm = report._formulas_model({
        "schema": "xlsprint.formulas/1",
        "groups": [{"group": "G1", "fingerprint": "abcdef", "sheet": "S", "cells": 3,
                    "areas": ["A1:A3", '=SUM("x")'], "functions": ["SUM", '=A1+"secret"'],
                    "formula": SENTINEL}],
        "totals": {"formula_cells": 3},
    })
    g = fm["groups"][0]
    assert g["areas"] == ["A1:A3"] and g["functions"] == ["SUM"]
    assert "formula" not in g


# --------------------------------------------------------------------------
# Integration with the real analysis module (skipped until it exists)
# --------------------------------------------------------------------------

def test_integration_synthetic(tmp_path):
    pytest.importorskip("xlsprint.analysis")
    synth = pytest.importorskip("xlsprint.synth")
    on, off = synth.synthetic_traces(tmp_path, seed=1, sheets=3, passes=5)
    fpath = None
    try:
        from xlsprint import formulas, synthbook
    except ImportError:
        formulas = None
    if formulas is not None:
        wb = synthbook.make_synthetic_workbook(tmp_path / "synth.xlsx", seed=1)
        fobj = formulas.inspect_workbook(wb, salt="test")
        fpath = tmp_path / "formulas.json"
        formulas.write_formulas_json(fobj, fpath)
    out = report.render([on, off], fpath, tmp_path / "report.html")
    text = out.read_text(encoding="utf-8")
    assert "Some sections could not be analysed" not in text
    assert "Hotspots" in text and "Drill-down" in text and "resolvable" in text
    if fpath is not None:
        assert "Formula groups" in text
        assert "Formula inspection limits" in text
    # Only one trace-on, trace-off optional: render without off says overhead unavailable.
    solo = report.render([on], None, tmp_path / "solo.html").read_text(encoding="utf-8")
    assert "overhead is unavailable" in solo
    for sid in SECTION_IDS:
        assert f'id="{sid}"' in text, sid
    assert "SYNTHETIC — not a real-workbook measurement" in text
    assert "INVALID TRACE" not in text
    assert not re.search(r"https?://", text)
    for tag in re.findall(r'<(?:span|tspan)\b[^>]*\bclass="num"[^>]*>', text):
        assert "data-class=" in tag


def test_integration_real_invalid_trace_refused(tmp_path):
    pytest.importorskip("xlsprint.analysis")
    synth = pytest.importorskip("xlsprint.synth")
    on, off = synth.synthetic_traces(tmp_path, seed=2, sheets=2, passes=3)
    lines = on.read_text(encoding="utf-8").splitlines()
    # Drop the first E event: its B stays open, so validation must fail.
    idx = next(i for i, ln in enumerate(lines) if json.loads(ln).get("type") == "E")
    broken = tmp_path / "broken-on.jsonl"
    broken.write_text("\n".join(lines[:idx] + lines[idx + 1:]) + "\n", encoding="utf-8")
    out = tmp_path / "report.html"
    with pytest.raises(report.ReportError):
        report.render([broken, off], None, out)
    assert not out.exists()
    try:
        report.render([broken, off], None, out, allow_invalid=True)
    except Exception as exc:  # analysis may refuse a structurally broken trace outright
        pytest.skip(f"analysis cannot load this invalid trace: {exc}")
    text = out.read_text(encoding="utf-8")
    assert "INVALID TRACE" in text


# --------------------------------------------------------------------------
# Review findings
# --------------------------------------------------------------------------

def _cell_classes(out, row_marker):
    row = re.search(r"<tr>(?:(?!</tr>).)*" + re.escape(row_marker) + r"(?:(?!</tr>).)*</tr>", out, re.S).group(0)
    return re.findall(r'<td[^>]*>(.*?)</td>', row, re.S)


def test_hotspot_badges():
    m = make_model()
    m["hotspots"] = [
        {"kind": "vba.proc", "name": "ProcWithKids", "sheet": None, "median_ns": 5_000_000,
         "median_self_ns": 1_000_000, "n": 5, "min_ns": 4_000_000, "max_ns": 6_000_000, "measurement": "derived"},
        {"kind": "calc.full", "name": "LeafFull", "sheet": None, "median_ns": 7_000_000,
         "median_self_ns": 7_000_000, "n": 5, "min_ns": 6_000_000, "max_ns": 8_000_000, "measurement": "measured"},
    ]
    out = report.render_html(m)
    cells = _cell_classes(out, "ProcWithKids")
    assert 'data-class="derived"' in cells[4]            # self: derived, with formula
    assert report.F_SELF in cells[4]
    assert 'data-class="measured"' in cells[5] and 'derived' not in cells[5]   # median total
    assert cells[7].count('data-class="measured"') == 2 and 'derived' not in cells[7]  # min–max
    leaf = _cell_classes(out, "LeafFull")
    assert 'data-class="measured"' in leaf[4] and 'derived' not in leaf[4]


def test_huge_numbers_do_not_crash():
    assert report.fmt_ns(10 ** 400) == "n/a"
    assert report.fmt_ns(float("inf")) == "n/a"
    m = make_model()
    m["host_stages"][0]["dur_ns"] = 10 ** 400
    m["timeline"]["rows"][0]["start_rel_ns"] = 10 ** 400
    m["drilldown"]["workbook"]["full"]["median"] = 10 ** 400
    m["meta"]["clock"]["vba_offset_ns"] = -(10 ** 400)
    m["hotspots"][0]["median_ns"] = 10 ** 400
    out = report.render_html(m)
    assert "1" * 50 not in out and "0" * 50 not in out
    blob = re.search(r'id="xlsprint-model">(.*?)</script>', out, re.S).group(1)
    json.loads(blob)


@pytest.mark.parametrize("header", [
    {"mode": "trace-on", "host": ["x"], "redaction": "clear", "clock": "qpc", "source": "synthetic"},
    {"mode": "trace-on", "host": "Windows", "source": {"kind": {"a": 1}, "label": ["x"]}, "run_id": {"x": 1}},
    ["not", "a", "dict"],
])
def test_build_model_tolerates_odd_headers(monkeypatch, header):
    monkeypatch.setattr(report, "_analysis", lambda: _fake_analysis())
    run = _fake_run()
    run.header = header
    model = report.build_model(run, None, None)
    assert model["meta"]["host"]["os"] is None
    report.render_html(model)


def test_valid_trace_section_errors_raise(monkeypatch):
    A = _fake_analysis()

    def boom(run, top=25):
        raise RuntimeError("analysis bug")
    A.hotspots = boom
    monkeypatch.setattr(report, "_analysis", lambda: A)
    with pytest.raises(RuntimeError, match="analysis bug"):
        report.build_model(_fake_run(), None, None)
    # On an invalid trace the same failure becomes a visible section error.
    model = report.build_model(_fake_run(ok=False), None, None)
    assert "hotspots" in model["section_errors"]


def test_drilldown_children_capped_in_html_and_json():
    m = make_model()
    kids = [{"kind": "calc.range", "name": f"blk{i:04d}", "summary": _summ(1000 + i), "address": f"A{i}"}
            for i in range(300)]
    m["drilldown"]["sheets"][0]["children"] = kids
    m["drilldown"]["sheets"][0]["children_truncated"] = 7
    m["drilldown"]["sheets"][1]["children"] = kids * 10
    out = report.render_html(m)
    assert "blk0199" in out and "blk0200" not in out
    assert "more not shown" in out
    blob = json.loads(re.search(r'id="xlsprint-model">(.*?)</script>', out, re.S).group(1))
    s0, s1 = blob["drilldown"]["sheets"]
    assert len(s0["children"]) == report.MAX_DD_CHILDREN_PER_SHEET
    assert s0["children_truncated"] == 7 + 100
    assert len(s0["children"]) + len(s1["children"]) <= report.MAX_DD_CHILDREN_TOTAL
    assert s1["children_truncated"] == 3000 - len(s1["children"])
    # Idempotent: capping an already capped model changes nothing.
    capped = report._apply_caps(m)
    assert report._apply_caps(capped)["drilldown"] == capped["drilldown"]


def test_timeline_rows_capped():
    m = make_model()
    row = m["timeline"]["rows"][-1]
    m["timeline"]["rows"] = [dict(row, id=i) for i in range(report.TIMELINE_MAX_ROWS + 50)]
    m["timeline"]["truncated"] = 0
    out = report.render_html(m)
    blob = json.loads(re.search(r'id="xlsprint-model">(.*?)</script>', out, re.S).group(1))
    assert len(blob["timeline"]["rows"]) == report.TIMELINE_MAX_ROWS
    assert blob["timeline"]["truncated"] == 50
    assert "Timeline truncated" in out


def test_overhead_unavailable_dict_and_formula(monkeypatch):
    A = _fake_analysis()
    A.overhead = lambda on, off: {"available": False, "reason": "run_id mismatch <on> vs <off>", "rows": []}
    monkeypatch.setattr(report, "_analysis", lambda: A)
    model = report.build_model(_fake_run(), _fake_run("trace-off"), None)
    assert model["overhead"]["available"] is False
    out = report.render_html(model)
    assert "run_id mismatch &lt;on&gt; vs &lt;off&gt;" in out

    m = make_model()
    m["overhead"]["rows"][0].update(pairs=2, resolvable=False, reason="only 2 paired passes")
    out = report.render_html(m)
    assert "not resolvable (below run-to-run noise)" in out and "fewer than 3 pairs" in out
    assert "only 2 paired passes" in out
    assert ("overhead = median over pairs of (on − off); resolvable when ≥3 pairs and "
            "min–max of paired differences excludes 0") in out


def test_overhead_rows_final_shape():
    m = make_model()
    base = m["overhead"]["rows"][0]
    m["overhead"]["rows"] = [
        dict(base, available=True, unavailable_reason=None, dropped_pairs=2, dropped_passes=[4, 5],
             pairs=3, resolvable=True, diff_min_ns=1_000, diff_max_ns=9_000),
        {"kind": "calc.full", "name": "Application.CalculateFull", "available": False,
         "unavailable_reason": "trace-on and trace-off run_id differ ('a' vs '<b>')",
         "median_on_ns": 123_456_789, "diff_ns": 987_654_321, "pairs": 0, "resolvable": False},
    ]
    out = report.render_html(m)
    assert "unavailable: trace-on and trace-off run_id differ (&#x27;a&#x27; vs &#x27;&lt;b&gt;&#x27;)" in out
    row = re.search(r"<tr>(?:(?!</tr>).)*CalculateFull(?:(?!</tr>).)*</tr>", out.split('id="overhead"')[1], re.S).group(0)
    assert "123.46 ms" not in row and "987.65 ms" not in row
    assert "dropped pairs" in out and "(pass 4, 5)" in out
    assert "for information only" in out
    assert "noise: IQR" not in out
    blob = json.loads(re.search(r'id="xlsprint-model">(.*?)</script>', out, re.S).group(1))
    assert blob["overhead"]["rows"][0]["dropped_passes"] == [4, 5]
    assert blob["overhead"]["rows"][1]["available"] is False


def test_overhead_all_unavailable_summary():
    m = make_model()
    m["overhead"]["rows"] = [
        {"kind": k, "name": k, "available": False, "unavailable_reason": "trace-on and trace-off run_id differ ('a' vs 'b')",
         "pairs": 0, "resolvable": False, "dropped_pairs": 0, "dropped_passes": []}
        for k in ("run.pass", "calc.full", "calc.recalc")]
    out = report.render_html(m)
    assert "Instrumentation overhead is unavailable: trace-on and trace-off run_id differ" in out
    assert out.count("unavailable: trace-on and trace-off run_id differ") == 4
