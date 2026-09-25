import pytest

from xlsprint import analysis
from xlsprint.analysis import (
    drilldown,
    host_stages,
    hotspots,
    load_run,
    overhead,
    passes,
    phase_breakdown,
    self_time_ns,
    summarize,
    timeline,
    union_ns,
)
from xlsprint.synth import synthetic_traces
from xlsprint.trace import TraceWriter


def _header(mode="trace-on", offset=0):
    return {
        "run_id": "0" * 32,
        "mode": mode,
        "clock": {"host": "perf_counter_ns", "vba": "MicroTimer/QPC", "vba_offset_ns": offset, "vba_offset_uncertainty_ns": 10},
        "source": {"kind": "synthetic", "label": "synthetic:unit"},
    }


class _VBA:
    """Tiny helper to write nested VBA spans with explicit ns."""

    def __init__(self, w):
        self.w = w

    def span(self, kind, name, start, end, pass_=1, attrs=None, body=None):
        sid = self.w.begin(kind, name, clock="vba", pass_=pass_, attrs=attrs, ns=start)
        if body:
            body()
        self.w.end(sid, ns=end)
        return sid


def _pass(v, p, t, full, recalc, extra=None):
    """Write run.pass p starting at t with calc.full and calc.recalc; return end time."""
    end_holder = {}

    def body():
        v.span("calc.full", "Application.CalculateFull", t + 1, t + 1 + full, pass_=p)
        r0 = t + 2 + full
        v.span("calc.recalc", "Application.Calculate", r0, r0 + recalc, pass_=p)
        cur = r0 + recalc + 1
        if extra:
            cur = extra(p, cur)
        end_holder["t"] = cur + 1

    rp = v.w.begin("run.pass", "XSP_RunPass", clock="vba", pass_=p, ns=t)
    body()
    v.w.end(rp, ns=end_holder["t"])
    return end_holder["t"]


def _run_with_passes(tmp_path, name, fulls, recalcs, mode="trace-on", extra=None):
    w = TraceWriter(tmp_path / name, _header(mode))
    v = _VBA(w)
    t = 1000
    for p, (f, r) in enumerate(zip(fulls, recalcs), 1):
        t = _pass(v, p, t, f, r, extra) + 10
    w.close()
    return load_run(tmp_path / name)


# ---------------------------------------------------------------- primitives


def test_union_ns():
    assert union_ns([]) == 0
    assert union_ns([(0, 10), (20, 30)]) == 20
    assert union_ns([(0, 10), (5, 15)]) == 15  # overlap counted once
    assert union_ns([(0, 100), (10, 20), (30, 40)]) == 100  # nested
    assert union_ns([(0, 10), (10, 20)]) == 20  # touching
    assert union_ns([(5, 5), (7, 3)]) == 0  # empty / inverted ignored
    assert union_ns(iter([(30, 40), (0, 10), (35, 50)])) == 30


def test_summarize():
    s = summarize([4, 1, 3, 2])
    assert s == {"n": 4, "median": 2.5, "min": 1, "max": 4, "p25": 1.75, "p75": 3.25}
    assert summarize([7]) == {"n": 1, "median": 7, "min": 7, "max": 7, "p25": 7, "p75": 7}
    assert summarize([])["n"] == 0 and summarize([])["median"] is None


def test_overlapping_children_not_double_counted(tmp_path):
    w = TraceWriter(tmp_path / "t.jsonl", _header())
    v = _VBA(w)

    def body():
        # Strict nesting forbids overlapping siblings, so overlap comes from a
        # grandchild (A1 inside A) that must not be counted twice.
        v.span("vba.proc", "A", 110, 140, body=lambda: v.span("vba.proc", "A1", 120, 130))
        v.span("vba.proc", "B", 150, 170)

    rp = w.begin("run.pass", "XSP_RunPass", clock="vba", pass_=1, ns=100)
    v.span("calc.full", "Application.CalculateFull", 101, 102)
    v.span("calc.recalc", "Application.Calculate", 103, 104)
    outer = v.span("vba.proc", "Outer", 105, 205, body=body)
    w.end(rp, ns=210)
    w.close()
    run = load_run(tmp_path / "t.jsonl")
    assert run.validation.ok, run.validation.errors
    # Outer 100ns, children A (30) + B (20); A1 is inside A and does not count again.
    assert self_time_ns(run, outer) == 100 - 50
    # The per-pass union of vba.proc spans equals Outer's extent, not the sum 100+30+10+20.
    assert passes(run)[0]["by_kind_union_ns"]["vba.proc"] == 100
    ph = {r["kind"]: r for r in phase_breakdown(run)}
    assert ph["vba.proc"]["union_ns_median_per_pass"] == 100
    assert ph["vba.proc"]["count"] == 4


def test_self_time_clips_children_to_parent():
    run = analysis.Run(header={}, spans={}, markers=[], validation=None)
    run.spans[1] = analysis.Span(1, 0, 1, 0, "vba.proc", "P", "vba", 0, 100, "ok", {}, [2, 3])
    run.spans[2] = analysis.Span(2, 1, 1, 1, "vba.proc", "C", "vba", 10, 60, "ok", {})
    run.spans[3] = analysis.Span(3, 1, 1, 1, "vba.proc", "D", "vba", 40, 150, "ok", {})  # past parent end
    assert self_time_ns(run, 1) == 100 - 90
    assert self_time_ns(run, 2) == 50


def test_vba_offset_alignment(tmp_path):
    offset = 5_000_000
    w = TraceWriter(tmp_path / "t.jsonl", _header(offset=offset))
    stage = w.begin("host.stage", "profile_trace_on", ns=1_000)
    rp = w.begin("run.pass", "XSP_RunPass", clock="vba", pass_=1, ns=offset + 1_100)
    f = w.begin("calc.full", "Application.CalculateFull", clock="vba", pass_=1, ns=offset + 1_200)
    w.end(f, ns=offset + 1_300)
    r = w.begin("calc.recalc", "Application.Calculate", clock="vba", pass_=1, ns=offset + 1_400)
    w.end(r, ns=offset + 1_450)
    w.end(rp, ns=offset + 1_500)
    w.end(stage, ns=2_000)
    w.close()
    run = load_run(tmp_path / "t.jsonl")
    assert run.validation.ok, run.validation.errors
    s = run.spans[rp]
    assert (s.start_ns, s.end_ns, s.dur_ns) == (1_100, 1_500, 400)
    # Host stage self time subtracts the aligned VBA child.
    assert self_time_ns(run, stage) == 1_000 - 400
    rows = {row["id"]: row for row in timeline(run)}
    assert rows[stage]["start_rel_ns"] == 0
    assert rows[rp]["start_rel_ns"] == 100
    assert rows[f]["start_rel_ns"] == 200
    assert host_stages(run) == [{"name": "profile_trace_on", "start_rel_ns": 0, "dur_ns": 1_000, "status": "ok"}]


# ---------------------------------------------------------------- views


def test_drilldown_and_hotspots(tmp_path):
    def extra(p, cur):
        v = _VBA(w_holder["w"])
        for sheet, dur in (("Inputs", 10), ("Model", 30)):
            v.span("calc.sheet", sheet, cur, cur + dur + p, pass_=p, attrs={"sheet": sheet})
            cur += dur + p + 1
        # A range and a name over the same cells: listed separately.
        v.span("calc.range", "Model!A1:B10", cur, cur + 8, pass_=p, attrs={"sheet": "Model", "address": "Model!A1:B10"})
        cur += 9
        v.span("calc.name", "Block", cur, cur + 7, pass_=p, attrs={"address": "Model!A1:B10"})
        cur += 8
        v.span("calc.group", "G0001", cur, cur + 5, pass_=p, attrs={"sheet": "Inputs", "group": "G0001"})
        return cur + 6

    w_holder = {}
    w = w_holder["w"] = TraceWriter(tmp_path / "d.jsonl", _header())
    v = _VBA(w)
    t = 1000
    for p, (full, recalc) in enumerate([(100, 20), (200, 40), (300, 60)], 1):
        t = _pass(v, p, t, full, recalc, extra) + 10
    w.close()
    run = load_run(tmp_path / "d.jsonl")
    assert run.validation.ok, run.validation.errors

    d = drilldown(run)
    assert d["workbook"]["full"]["median"] == 200 and d["workbook"]["full"]["n"] == 3
    assert d["workbook"]["recalc"]["median"] == 40
    assert d["workbook"]["volatility_ratio"] == pytest.approx(0.2)
    sheets = {s["name"]: s for s in d["sheets"]}
    assert [s["name"] for s in d["sheets"]] == ["Model", "Inputs"]
    assert sheets["Model"]["recalc"]["median"] == 32
    assert sum(s["share_of_recalc_sum"] for s in d["sheets"]) == pytest.approx(1.0)
    model_children = {(c["kind"], c["name"]): c for c in sheets["Model"]["children"]}
    assert set(model_children) == {("calc.range", "Model!A1:B10"), ("calc.name", "Block")}
    assert model_children[("calc.range", "Model!A1:B10")]["summary"]["n"] == 3
    assert model_children[("calc.name", "Block")]["address"] == "Model!A1:B10"
    grp = sheets["Inputs"]["children"][0]
    assert grp["kind"] == "calc.group" and "isolated" in grp["measurement"]

    hs = hotspots(run, top=3)
    assert [h["kind"] for h in hs] == ["calc.full", "calc.recalc", "calc.sheet"]
    assert hs[0]["median_self_ns"] == 200 and hs[0]["measurement"] == "measured"
    all_hs = hotspots(run)
    names = {(h["kind"], h["name"]) for h in all_hs}
    assert ("calc.range", "Model!A1:B10") in names and ("calc.name", "Block") in names
    assert next(h for h in all_hs if h["kind"] == "calc.range")["address"] == "Model!A1:B10"
    assert [h for h in all_hs if h["kind"] == "calc.group"][0]["note"].startswith("measured (isolated")
    assert not any(h["kind"] in ("run.pass", "host.stage") for h in all_hs)


def test_hotspot_measurement_label_derived_for_parents(tmp_path):
    w = TraceWriter(tmp_path / "h.jsonl", _header())
    v = _VBA(w)
    rp = w.begin("run.pass", "XSP_RunPass", clock="vba", pass_=1, ns=0)
    v.span("calc.full", "Application.CalculateFull", 1, 2)
    v.span("calc.recalc", "Application.Calculate", 3, 4)
    v.span("vba.proc", "Macro", 10, 110, body=lambda: v.span("vba.proc", "Inner", 20, 90))
    w.end(rp, ns=200)
    w.close()
    hs = {h["name"]: h for h in hotspots(load_run(tmp_path / "h.jsonl"))}
    assert hs["Macro"]["measurement"] == "derived" and hs["Macro"]["median_self_ns"] == 30
    assert hs["Inner"]["measurement"] == "measured" and hs["Inner"]["median_self_ns"] == 70
    assert list(hs)[0] == "Inner"  # ranked by self time, not total


def test_timeline_bounded(tmp_path):
    run = _run_with_passes(tmp_path, "t.jsonl", [10] * 10, [5] * 10)
    assert len(run.spans) == 30
    rows = timeline(run, max_rows=7)
    assert len(rows) == 8 and rows[-1] == {"truncated": 23}
    kept = rows[:-1]
    assert all(r["kind"] == "run.pass" for r in kept)  # shallowest first
    assert [r["start_rel_ns"] for r in kept] == sorted(r["start_rel_ns"] for r in kept)
    assert "truncated" not in timeline(run)[-1]


def test_passes_rows(tmp_path):
    run = _run_with_passes(tmp_path, "p.jsonl", [10, 20], [5, 6])
    rows = passes(run)
    assert [r["pass"] for r in rows] == [1, 2]
    assert rows[1]["by_kind_union_ns"] == {"calc.full": 20, "calc.recalc": 6}


# ---------------------------------------------------------------- overhead


def test_overhead_resolvable(tmp_path):
    on = _run_with_passes(tmp_path, "on.jsonl", [110, 111, 112, 113, 114], [50] * 5)
    off = _run_with_passes(tmp_path, "off.jsonl", [100, 101, 102, 103, 104], [50] * 5, mode="trace-off")
    rows = {r["kind"]: r for r in overhead(on, off)}
    full = rows["calc.full"]
    assert full["available"] is True and full["pairs"] == 5 and full["dropped_pairs"] == 0
    assert full["diff_ns"] == 10 and (full["diff_min_ns"], full["diff_max_ns"]) == (10, 10)
    assert full["resolvable"] is True
    assert rows["run.pass"]["resolvable"] is True
    # Identical recalc timings: the per-pair range [0, 0] includes 0.
    assert rows["calc.recalc"]["diff_ns"] == 0 and rows["calc.recalc"]["resolvable"] is False


def test_overhead_diff_is_median_of_pair_differences(tmp_path):
    on = _run_with_passes(tmp_path, "on.jsonl", [100, 150, 120, 90, 130], [50] * 5)
    off = _run_with_passes(tmp_path, "off.jsonl", [105, 95, 140, 110, 100], [50] * 5, mode="trace-off")
    full = {r["kind"]: r for r in overhead(on, off)}["calc.full"]
    # Per-pair diffs: -5, 55, -20, -20, 30 -> median -5 (median-of-medians would say 15).
    assert full["diff_ns"] == -5
    assert (full["diff_min_ns"], full["diff_max_ns"]) == (-20, 55)
    assert full["resolvable"] is False  # range straddles 0


def test_overhead_needs_three_pairs(tmp_path):
    on = _run_with_passes(tmp_path, "on.jsonl", [110, 120], [5] * 2)
    off = _run_with_passes(tmp_path, "off.jsonl", [100, 100], [5] * 2, mode="trace-off")
    full = {r["kind"]: r for r in overhead(on, off)}["calc.full"]
    assert full["pairs"] == 2 and full["diff_min_ns"] == 10 and full["resolvable"] is False


def test_overhead_pairs_by_pass_number_not_position(tmp_path):
    on = _run_with_passes(tmp_path, "on.jsonl", [110, 120, 130, 140], [5] * 4)
    off = _run_with_passes(tmp_path, "off.jsonl", [100, 110, 120, 130], [5] * 4, mode="trace-off")
    # Drop off pass 1: positional pairing would compare on1 with off2.
    for sid in [s.id for s in off.spans.values() if s.pass_ == 1]:
        del off.spans[sid]
    full = {r["kind"]: r for r in overhead(on, off)}["calc.full"]
    assert full["pairs"] == 3 and full["dropped_pairs"] == 1 and full["dropped_passes"] == [1]
    assert (full["diff_min_ns"], full["diff_max_ns"]) == (10, 10)
    assert full["median_on_ns"] == 130


def test_overhead_skips_non_ok_passes(tmp_path):
    on = _run_with_passes(tmp_path, "on.jsonl", [110, 120, 130, 140], [5] * 4)
    off = _run_with_passes(tmp_path, "off.jsonl", [100, 110, 120, 130], [5] * 4, mode="trace-off")
    rp2 = [s for s in on.spans.values() if s.kind == "run.pass" and s.pass_ == 2][0]
    rp2.status = "error"
    for row in overhead(on, off):
        assert row["pairs"] == 3 and row["dropped_passes"] == [2], row


def test_overhead_unavailable_when_run_ids_differ(tmp_path):
    on = _run_with_passes(tmp_path, "on.jsonl", [110] * 3, [5] * 3)
    off = _run_with_passes(tmp_path, "off.jsonl", [100] * 3, [5] * 3, mode="trace-off")
    off.header["run_id"] = "f" * 32
    rows = overhead(on, off)
    assert len(rows) == 3
    for row in rows:
        assert row["available"] is False and "run_id" in row["unavailable_reason"]
        assert row["diff_ns"] is None and row["pairs"] == 0 and row["resolvable"] is False


def test_drilldown_caps_children_per_sheet(tmp_path):
    n = analysis.MAX_CHILDREN_PER_SHEET + 25

    def extra(p, cur):
        v = _VBA(w)
        v.span("calc.sheet", "S", cur, cur + 10, pass_=p)
        cur += 11
        for i in range(n):
            v.span("calc.name", "N%04d" % i, cur, cur + 1 + i, pass_=p, attrs={"sheet": "S"})
            cur += 2 + i
        return cur

    w = TraceWriter(tmp_path / "cap.jsonl", _header())
    t = 0
    for p in (1, 2):
        t = _pass(_VBA(w), p, t, 5, 5, extra) + 1
    w.close()
    sheet = drilldown(load_run(tmp_path / "cap.jsonl"))["sheets"][0]
    assert len(sheet["children"]) == analysis.MAX_CHILDREN_PER_SHEET
    assert sheet["children_truncated"] == 25
    assert sheet["children"][0]["name"] == "N%04d" % (n - 1)  # largest kept


# ---------------------------------------------------------------- loading + synth


def test_load_run_tolerates_invalid(tmp_path):
    w = TraceWriter(tmp_path / "bad.jsonl", _header())
    rp = w.begin("run.pass", "XSP_RunPass", clock="vba", pass_=1, ns=0)
    f = w.begin("calc.full", "Application.CalculateFull", clock="vba", pass_=1, ns=1)
    w.end(f, ns=5)
    w.close()  # run.pass left open
    run = load_run(tmp_path / "bad.jsonl")
    assert not run.validation.ok
    assert list(run.spans) == [f] and rp not in run.spans
    assert run.source_kind == "synthetic" and run.mode == "trace-on"


def test_synthetic_traces_valid_and_deterministic(tmp_path):
    on, off = synthetic_traces(tmp_path / "a", seed=3, sheets=4, passes=4)
    on2, off2 = synthetic_traces(tmp_path / "b", seed=3, sheets=4, passes=4)
    assert on.read_bytes() == on2.read_bytes() and off.read_bytes() == off2.read_bytes()
    on3, _ = synthetic_traces(tmp_path / "c", seed=4, sheets=4, passes=4)
    assert on.read_bytes() != on3.read_bytes()

    r_on, r_off = load_run(on), load_run(off)
    assert r_on.validation.ok, r_on.validation.errors
    assert r_off.validation.ok, r_off.validation.errors
    assert r_on.source_kind == "synthetic" and r_on.header["source"]["label"] == "synthetic:generated"
    assert r_on.mode == "trace-on" and r_off.mode == "trace-off"
    kinds = {s.kind for s in r_on.spans.values()}
    assert kinds >= {"host.stage", "run.pass", "calc.full", "calc.recalc", "calc.sheet", "calc.range", "calc.name", "calc.group", "vba.proc"}
    assert {s.kind for s in r_off.spans.values()} == {"host.stage", "run.pass", "calc.full", "calc.recalc"}
    assert sum(1 for s in r_on.spans.values() if s.kind == "calc.group") == 4
    # vba.proc with a nested vba.proc child in every pass
    nested = [s for s in r_on.spans.values() if s.kind == "vba.proc" and r_on.spans[s.parent].kind == "vba.proc"]
    assert len(nested) == 4
    d = drilldown(r_on)
    assert [s["name"] for s in d["sheets"]] and {s["name"] for s in d["sheets"]} <= {"Inputs", "Model", "Report", "Lookup"}
    rows = {r["kind"]: r for r in overhead(r_on, r_off)}
    assert r_on.header["run_id"] == r_off.header["run_id"]
    assert rows["run.pass"]["available"] and rows["run.pass"]["pairs"] == 4 and rows["run.pass"]["diff_ns"] > 0
    assert len(host_stages(r_on)) >= 5


def test_synthetic_trace_has_no_formula_text(tmp_path):
    on, off = synthetic_traces(tmp_path, seed=0)
    text = on.read_text() + off.read_text()
    assert "=" not in text and '\\"' not in text


def test_stray_spans_excluded_from_views(tmp_path):
    run = _run_with_passes(tmp_path, "s.jsonl", [100, 200, 300], [10, 20, 30])
    # Simulate an --allow-invalid load: a calc.full at the root and a sheet claiming another pass.
    run.spans[900] = analysis.Span(900, 0, 0, 0, "calc.full", "Application.CalculateFull", "vba", 0, 10**12, "ok", {})
    rp1 = [s for s in run.spans.values() if s.kind == "run.pass" and s.pass_ == 1][0]
    run.spans[901] = analysis.Span(901, rp1.id, 9, 1, "calc.sheet", "Ghost", "vba", rp1.start_ns, rp1.start_ns + 1, "ok", {})
    rp1.children.append(901)
    d = drilldown(run)
    assert d["workbook"]["full"]["n"] == 3 and d["workbook"]["full"]["max"] == 300
    assert not d["sheets"]
    assert all(h["name"] != "Ghost" for h in hotspots(run))


def test_overhead_row_unavailable_without_pairs(tmp_path):
    on = _run_with_passes(tmp_path, "on.jsonl", [110] * 3, [5] * 3)
    off = _run_with_passes(tmp_path, "off.jsonl", [100] * 3, [5] * 3, mode="trace-off")
    for sid in [s.id for s in off.spans.values() if s.kind == "calc.recalc"]:
        del off.spans[sid]
    rows = {r["kind"]: r for r in overhead(on, off)}
    assert rows["calc.full"]["available"] is True and rows["calc.full"]["unavailable_reason"] is None
    rec = rows["calc.recalc"]
    assert rec["available"] is False and rec["pairs"] == 0 and rec["dropped_passes"] == [1, 2, 3]
    assert "calc.recalc" in rec["unavailable_reason"] and rec["diff_ns"] is None
