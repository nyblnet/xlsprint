import json
import re

import pytest

from xlsprint import trace
from xlsprint.trace import SCHEMA, TraceError, TraceWriter, read_trace, read_vba_events, redact_name, validate


def _header(**kw):
    h = {
        "type": "header",
        "schema": SCHEMA,
        "run_id": "0" * 32,
        "mode": "trace-on",
        "clock": {"host": "perf_counter_ns", "vba": "MicroTimer/QPC", "vba_offset_ns": 0, "vba_offset_uncertainty_ns": 1000},
        "source": {"kind": "synthetic", "label": "synthetic:unit"},
    }
    h.update(kw)
    return h


def _b(i, kind, name, parent, depth, ns, clock="vba", pass_=1, attrs=None):
    return {"type": "B", "id": i, "parent": parent, "pass": pass_, "depth": depth, "ns": ns,
            "clock": clock, "kind": kind, "name": name, "attrs": attrs or {}}


def _e(i, ns, clock="vba", status="ok"):
    return {"type": "E", "id": i, "ns": ns, "clock": clock, "status": status}


def _events():
    return [
        _b(1, "host.stage", "profile_trace_on", 0, 0, 100, clock="host", pass_=0),
        _b(2, "run.pass", "XSP_RunPass", 1, 1, 1000),
        _b(3, "calc.full", "Application.CalculateFull", 2, 2, 1100),
        _e(3, 1200),
        _b(4, "calc.recalc", "Application.Calculate", 2, 2, 1300),
        _e(4, 1400),
        {"type": "M", "id": 5, "parent": 2, "pass": 1, "ns": 1450, "clock": "vba", "kind": "marker", "name": "AfterCalculate", "attrs": {}},
        _b(6, "calc.sheet", "Model", 2, 2, 1500, attrs={"method": "Worksheet.Calculate", "sheet": "Model"}),
        _e(6, 1600),
        _e(2, 1700),
        _e(1, 5000, clock="host"),
    ]


def _write(tmp_path, events=None, header=None, footer=None, name="t.jsonl"):
    events = _events() if events is None else events
    header = _header() if header is None else header
    if footer is None:
        footer = {"type": "footer", "events": len(events), "dropped": 0, "truncated": False, "open_spans_at_close": 0}
    lines = []
    if header is not False:
        lines.append(json.dumps(header))
    lines += [json.dumps(e) for e in events]
    if footer is not False:
        lines.append(json.dumps(footer))
    p = tmp_path / name
    p.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return p


def _errors_with(res, tag):
    return [e for e in res.errors if e.startswith(tag)]


# ---------------------------------------------------------------- validator


def test_baseline_valid(tmp_path):
    res = validate(_write(tmp_path))
    assert res.ok, res.errors
    assert res.stats["passes"] == [1]
    assert res.stats["by_type"] == {"B": 5, "E": 5, "M": 1}


def test_rule1_missing_header(tmp_path):
    res = validate(_write(tmp_path, header=False))
    assert not res.ok and _errors_with(res, "[1]")


def test_rule1_missing_footer(tmp_path):
    res = validate(_write(tmp_path, footer=False))
    assert not res.ok and any("footer missing" in e for e in res.errors)


def test_rule1_unknown_schema(tmp_path):
    res = validate(_write(tmp_path, header=_header(schema="xlsprint.trace/99")))
    assert not res.ok and any("unknown schema" in e for e in res.errors)


def test_rule1_missing_source_kind(tmp_path):
    h = _header()
    del h["source"]
    res = validate(_write(tmp_path, header=h))
    assert not res.ok and any("source.kind" in e for e in res.errors)


def test_rule1_bad_json_line(tmp_path):
    p = _write(tmp_path)
    lines = p.read_text().splitlines()
    lines.insert(2, "{not json")
    p.write_text("\n".join(lines) + "\n")
    assert not validate(p).ok


def test_rule1_vba_events_need_offset(tmp_path):
    h = _header()
    del h["clock"]["vba_offset_ns"]
    res = validate(_write(tmp_path, header=h))
    assert not res.ok and any("vba_offset_ns" in e for e in res.errors)


def test_rule2_unbalanced_end(tmp_path):
    ev = _events()
    ev.insert(4, _e(99, 1250))
    res = validate(_write(tmp_path, ev))
    assert any("no matching open B" in e for e in _errors_with(res, "[2]"))


def test_rule2_non_innermost_end(tmp_path):
    ev = _events()
    # Close run.pass (2) while calc.full (3) is still open.
    ev[3] = _e(2, 1200)
    res = validate(_write(tmp_path, ev))
    assert any("not innermost" in e for e in _errors_with(res, "[2]"))


def test_rule2_open_span_at_footer(tmp_path):
    ev = _events()[:-1]  # drop the host stage's E
    res = validate(_write(tmp_path, ev))
    assert any("still open at footer" in e for e in _errors_with(res, "[2]"))


def test_rule3_ns_backwards_same_clock(tmp_path):
    ev = _events()
    ev[4]["ns"] = 1150  # calc.recalc B before previous E (1200) on the vba clock
    ev[5]["ns"] = 1160
    res = validate(_write(tmp_path, ev))
    assert any("went backwards" in e for e in _errors_with(res, "[3]"))


def test_rule3_clocks_are_independent(tmp_path):
    # Host events at 100/5000 and VBA events in between interleave freely.
    assert validate(_write(tmp_path)).ok


def test_rule3_end_before_begin(tmp_path):
    ev = [
        _b(1, "run.pass", "XSP_RunPass", 0, 0, 1000),
        _b(2, "calc.full", "Application.CalculateFull", 1, 1, 1100),
        _e(2, 1200),
        _b(3, "calc.recalc", "Application.Calculate", 1, 1, 1300),
        _e(3, 1400),
        _b(4, "calc.sheet", "S", 1, 1, 1500),
        _b(5, "calc.range", "S!A1", 4, 2, 1500),
        _e(5, 1500),
        _e(4, 1500),
        _e(1, 1500),
    ]
    assert validate(_write(tmp_path, ev)).ok
    # E(5)=1500 now precedes its B(5)=1600 (this also trips the backwards check).
    ev[5]["ns"] = 1600
    ev[6]["ns"] = 1600
    res = validate(_write(tmp_path, ev))
    assert any("ends before its B" in e for e in _errors_with(res, "[3]"))


def test_rule4_duplicate_id(tmp_path):
    ev = _events()
    ev[4]["id"] = 3
    ev[5]["id"] = 3
    res = validate(_write(tmp_path, ev))
    assert any("duplicate id" in e for e in _errors_with(res, "[4]"))


@pytest.mark.parametrize("field,value", [("parent", 1), ("depth", 1), ("parent", 0)])
def test_rule4_bad_parent_or_depth(tmp_path, field, value):
    ev = _events()
    ev[2][field] = value  # calc.full should be parent=2, depth=2
    res = validate(_write(tmp_path, ev))
    assert any(field in e for e in _errors_with(res, "[4]"))


def test_rule9_pass_mismatch_inside_run_pass(tmp_path):
    ev = _events()
    ev[2]["pass"] = 7
    res = validate(_write(tmp_path, ev))
    assert any("differs from enclosing run.pass" in e for e in _errors_with(res, "[9]"))


def test_rule5_truncated(tmp_path):
    ev = _events()
    footer = {"type": "footer", "events": len(ev), "dropped": 0, "truncated": True}
    assert _errors_with(validate(_write(tmp_path, ev, footer=footer)), "[5]")


def test_rule5_dropped(tmp_path):
    ev = _events()
    footer = {"type": "footer", "events": len(ev), "dropped": 3, "truncated": False}
    assert _errors_with(validate(_write(tmp_path, ev, footer=footer)), "[5]")


def test_rule5_count_mismatch(tmp_path):
    ev = _events()
    footer = {"type": "footer", "events": len(ev) + 1, "dropped": 0, "truncated": False}
    res = validate(_write(tmp_path, ev, footer=footer))
    assert any("observed" in e for e in _errors_with(res, "[5]"))


def test_rule6_disallowed_attr_key(tmp_path):
    ev = _events()
    ev[7]["attrs"]["formula"] = "SUM(A1:A3)"
    res = validate(_write(tmp_path, ev))
    assert any("not allow-listed" in e for e in _errors_with(res, "[6]"))


@pytest.mark.parametrize("value", ["=SUM(A1:A3)", 'say "hi"', "x" * 257, None, ["a"], {"a": 1}])
def test_rule6_bad_attr_value(tmp_path, value):
    ev = _events()
    ev[7]["attrs"]["note"] = value
    assert _errors_with(validate(_write(tmp_path, ev)), "[6]")


def test_rule6_extra_event_key(tmp_path):
    ev = _events()
    ev[3]["value"] = 42  # an E carrying a cell value
    assert _errors_with(validate(_write(tmp_path, ev)), "[6]")


def test_rule7_missing_expected_instrumentation(tmp_path):
    h = _header(expect={"vba.proc": ["Macro_Refresh"], "calc.sheet": ["*"]})
    res = validate(_write(tmp_path, header=h))
    errs = _errors_with(res, "[7]")
    assert len(errs) == 1 and "Macro_Refresh" in errs[0]


def test_rule7_wildcard(tmp_path):
    h = _header(expect={"calc.group": ["*"]})
    assert _errors_with(validate(_write(tmp_path, header=h)), "[7]")
    h = _header(expect={"calc.sheet": ["*"], "calc.full": ["Application.CalculateFull"]})
    assert validate(_write(tmp_path, header=h)).ok


def test_rule8_trace_on_without_run_pass(tmp_path):
    ev = [_b(1, "host.stage", "warmup", 0, 0, 100, clock="host", pass_=0), _e(1, 200, clock="host")]
    res = validate(_write(tmp_path, ev))
    assert any("no run.pass" in e for e in _errors_with(res, "[8]"))


@pytest.mark.parametrize("drop_id,kind", [(3, "calc.full"), (4, "calc.recalc")])
def test_rule8_pass_lacks_full_or_recalc(tmp_path, drop_id, kind):
    ev = [e for e in _events() if e["id"] != drop_id]
    res = validate(_write(tmp_path, ev))
    assert any(kind in e for e in _errors_with(res, "[8]"))


def test_rule8_trace_off_only_warns(tmp_path):
    ev = [e for e in _events() if e["id"] not in (4, 5, 6)]
    res = validate(_write(tmp_path, ev, header=_header(mode="trace-off")))
    assert res.ok, res.errors
    assert any("calc.recalc" in w for w in res.warnings)


def test_kind_clock_mismatch_rejected(tmp_path):
    ev = _events()
    ev[2]["clock"] = "host"
    ev[3]["clock"] = "host"
    assert not validate(_write(tmp_path, ev)).ok


def test_name_that_looks_like_formula_rejected(tmp_path):
    ev = _events()
    ev[7]["name"] = "=A1*2"
    assert _errors_with(validate(_write(tmp_path, ev)), "[10]")


# ---------------------------------------------------------------- writer


def _writer(tmp_path, **kw):
    h = _header()
    del h["type"], h["schema"]
    return TraceWriter(tmp_path / "w.jsonl", h, **kw)


def test_writer_roundtrip_valid(tmp_path):
    w = _writer(tmp_path)
    sid = w.begin("host.stage", "profile_trace_on", ns=0)
    assert w.open_span_id == sid and w.open_depth == 1
    rp = w.begin("run.pass", "XSP_RunPass", clock="vba", pass_=1, ns=10)
    for i, kind in enumerate(("calc.full", "calc.recalc")):
        s = w.begin(kind, kind, clock="vba", pass_=1, ns=20 + i * 10)
        w.end(s, ns=25 + i * 10)
    w.marker("AfterCalculate", clock="vba", pass_=1, ns=50)
    w.end(rp, ns=60)
    w.end(sid, ns=100)
    footer = w.close()
    assert footer["events"] == 9 and footer["truncated"] is False and footer["dropped"] == 0
    assert not (tmp_path / "w.jsonl.part").exists()
    header, events, foot = read_trace(tmp_path / "w.jsonl")
    assert header["schema"] == SCHEMA and foot == footer
    assert events[1]["parent"] == events[0]["id"] and events[1]["depth"] == 1
    res = validate(tmp_path / "w.jsonl")
    assert res.ok, res.errors
    assert not any("bytes" in x for x in res.warnings)


def test_writer_uses_perf_counter_for_host(tmp_path, monkeypatch):
    ticks = iter([1000, 2000])
    monkeypatch.setattr(trace.time, "perf_counter_ns", lambda: next(ticks))
    w = _writer(tmp_path)
    sid = w.begin("host.stage", "warmup")
    w.end(sid)
    w.close()
    _, events, _ = read_trace(tmp_path / "w.jsonl")
    assert [e["ns"] for e in events] == [1000, 2000]


def test_writer_vba_requires_ns(tmp_path):
    w = _writer(tmp_path)
    with pytest.raises(TraceError):
        w.begin("run.pass", "XSP_RunPass", clock="vba", pass_=1)


def test_writer_rejects_non_innermost_end(tmp_path):
    w = _writer(tmp_path)
    a = w.begin("host.stage", "warmup")
    w.begin("host.stage", "instrument")
    with pytest.raises(TraceError):
        w.end(a)
    with pytest.raises(TraceError):
        w.end(12345)


@pytest.mark.parametrize("attrs", [{"formula": "x"}, {"note": "=1+1"}, {"note": 'a"b'}, {"note": None}])
def test_writer_rejects_bad_attrs(tmp_path, attrs):
    w = _writer(tmp_path)
    with pytest.raises(TraceError):
        w.begin("host.stage", "warmup", attrs=attrs)
    assert w.open_depth == 0


def test_writer_rejects_backwards_ns(tmp_path):
    w = _writer(tmp_path)
    w.begin("run.pass", "XSP_RunPass", clock="vba", pass_=1, ns=100)
    with pytest.raises(TraceError):
        w.begin("calc.full", "x", clock="vba", pass_=1, ns=50)


def test_writer_span_marks_error_and_reraises(tmp_path):
    w = _writer(tmp_path)
    with pytest.raises(ValueError):
        with w.span("host.stage", "open_workbook"):
            raise ValueError("boom")
    w.close()
    _, events, _ = read_trace(tmp_path / "w.jsonl")
    assert events[-1]["status"] == "error"


def test_writer_close_does_not_fabricate_end(tmp_path):
    w = _writer(tmp_path)
    w.begin("host.stage", "warmup")
    footer = w.close()
    assert footer["open_spans_at_close"] == 1
    _, events, _ = read_trace(tmp_path / "w.jsonl")
    assert [e["type"] for e in events] == ["B"]
    assert not validate(tmp_path / "w.jsonl").ok


@pytest.mark.parametrize("limit", [{"max_events": 5}, {"max_bytes": 1200}])
def test_writer_bounded_truncation(tmp_path, limit):
    w = _writer(tmp_path, **limit)
    rp = w.begin("run.pass", "XSP_RunPass", clock="vba", pass_=1, ns=1)
    ns = 2
    for i in range(20):
        s = w.begin("calc.sheet", "S%d" % i, clock="vba", pass_=1, ns=ns)
        w.end(s, ns=ns + 1)
        ns += 2
    w.end(rp, ns=ns)
    footer = w.close()
    assert footer["truncated"] is True
    assert footer["dropped"] == 42 - footer["events"]
    assert footer["open_spans_at_close"] == 0  # nesting stays balanced logically
    if "max_events" in limit:
        assert footer["events"] == 5
    else:
        assert (tmp_path / "w.jsonl").stat().st_size <= 1200  # footer included
    res = validate(tmp_path / "w.jsonl")
    assert not res.ok and _errors_with(res, "[5]")


def test_writer_update_header_late(tmp_path):
    w = _writer(tmp_path)
    s = w.begin("host.stage", "launch_excel")
    w.end(s)
    w.update_header({"clock": {"vba_offset_ns": 777}})
    w.close()
    header, _, _ = read_trace(tmp_path / "w.jsonl")
    assert header["clock"]["vba_offset_ns"] == 777
    assert header["clock"]["vba"] == "MicroTimer/QPC"


def test_writer_header_requires_source_kind(tmp_path):
    with pytest.raises(TraceError):
        TraceWriter(tmp_path / "x.jsonl", {"mode": "trace-on"})


def test_reserve_ids_and_append_event(tmp_path):
    w = _writer(tmp_path)
    stage = w.begin("host.stage", "profile_trace_on", ns=0)
    first = w.reserve_ids(100)
    # VBA-produced events: parent 0 / depth 0 relative to VBA, rebased here.
    vba = [
        {"type": "B", "id": first, "parent": 0, "pass": 1, "depth": 0, "ns": 500, "clock": "vba", "kind": "run.pass", "name": "XSP_RunPass", "attrs": {}},
        {"type": "B", "id": first + 1, "parent": first, "pass": 1, "depth": 1, "ns": 510, "clock": "vba", "kind": "calc.full", "name": "Application.CalculateFull", "attrs": {}},
        {"type": "E", "id": first + 1, "ns": 520, "clock": "vba", "status": "ok"},
        {"type": "B", "id": first + 2, "parent": first, "pass": 1, "depth": 1, "ns": 530, "clock": "vba", "kind": "calc.recalc", "name": "Application.Calculate", "attrs": {}},
        {"type": "E", "id": first + 2, "ns": 540, "clock": "vba", "status": "ok"},
        {"type": "E", "id": first, "ns": 550, "clock": "vba", "status": "ok"},
    ]
    for ev in vba:
        w.append_event(ev)
    nxt = w.begin("host.stage", "collect_trace", ns=600)
    assert nxt >= first + 100
    w.end(nxt, ns=700)
    w.end(stage, ns=800)
    w.close()
    _, events, _ = read_trace(tmp_path / "w.jsonl")
    rp = [e for e in events if e.get("kind") == "run.pass"][0]
    assert rp["parent"] == stage and rp["depth"] == 1
    assert validate(tmp_path / "w.jsonl").ok


@pytest.mark.parametrize(
    "bad",
    [
        {"type": "B", "id": 1, "parent": 0, "pass": 1, "ns": 5, "clock": "vba", "kind": "calc.full", "name": "x", "attrs": {}},  # duplicate id
        {"type": "B", "id": 50, "parent": 999, "pass": 1, "ns": 5, "clock": "vba", "kind": "calc.full", "name": "x", "attrs": {}},  # bad parent
        {"type": "B", "id": 51, "parent": 0, "pass": 1, "ns": 5, "clock": "vba", "kind": "calc.full", "name": "x", "attrs": {"cellvalue": 3}},
        {"type": "B", "id": 52, "parent": 0, "pass": 1, "ns": 5, "clock": "vba", "kind": "calc.full", "name": "x", "attrs": {}, "formula": "=1"},
        {"type": "B", "id": 53, "parent": 0, "pass": 1, "ns": 5, "clock": "vba", "kind": "nope", "name": "x", "attrs": {}},
        {"type": "E", "id": 77, "ns": 5, "clock": "vba", "status": "ok"},
    ],
)
def test_append_event_rejects(tmp_path, bad):
    w = _writer(tmp_path)
    w.begin("host.stage", "profile_trace_on")  # id 1
    with pytest.raises(TraceError):
        w.append_event(bad)


def test_redact_name():
    import hashlib

    assert redact_name("Model", "salt") == "h:" + hashlib.sha256(b"saltModel").hexdigest()[:10]
    assert redact_name("Model", "salt") != redact_name("Model", "other")


def test_read_vba_events(tmp_path):
    p = tmp_path / "vba.jsonl"
    lines = [
        json.dumps({"type": "B", "id": 1, "parent": 0, "pass": 1, "depth": 0, "ns": 1.0e9, "clock": "vba", "kind": "run.pass", "name": "XSP_RunPass", "attrs": {}}),
        json.dumps({"type": "E", "id": 1, "ns": 1000000500, "clock": "vba", "status": "ok"}),
        json.dumps({"type": "vba_footer", "events": 2, "dropped": 0, "truncated": False, "open_spans": 0}),
    ]
    p.write_bytes(b"\xef\xbb\xbf" + "\r\n".join(lines).encode() + b"\r\n")
    events, footer = read_vba_events(p)
    assert events[0]["ns"] == 1_000_000_000 and isinstance(events[0]["ns"], int)
    assert footer == {"type": "vba_footer", "events": 2, "dropped": 0, "truncated": False, "open_spans": 0}


def test_read_vba_events_requires_footer(tmp_path):
    p = tmp_path / "vba.jsonl"
    p.write_text(json.dumps({"type": "E", "id": 1, "ns": 1, "clock": "vba", "status": "ok"}) + "\n")
    with pytest.raises(TraceError):
        read_vba_events(p)


def test_vba_file_format_from_excel_module(tmp_path):
    # ASCII JSONL, CRLF, \uXXXX escapes, extra vba_footer keys kept for the runner.
    p = tmp_path / "vba.jsonl"
    lines = [
        '{"type":"B","id":1,"parent":0,"pass":1,"depth":0,"ns":5000,"clock":"vba","kind":"run.pass","name":"XSP_RunPass","attrs":{}}',
        '{"type":"B","id":2,"parent":1,"pass":1,"depth":1,"ns":5100,"clock":"vba","kind":"calc.sheet","name":"Mod\\u00e8le","attrs":{"sheet":"h:0123456789"}}',
        '{"type":"E","id":2,"ns":5200,"clock":"vba","status":"ok"}',
        '{"type":"E","id":1,"ns":5300,"clock":"vba","status":"ok"}',
        '{"type":"vba_footer","events":4,"dropped":0,"truncated":false,"open_spans":0,"faults":0,'
        '"attr_rejected":0,"first_fault":"","mode":"on","pass":1,"max_events":100000}',
    ]
    p.write_bytes("\r\n".join(lines).encode("ascii") + b"\r\n")
    events, footer = read_vba_events(p)
    assert events[1]["name"] == "Modèle"
    assert footer["faults"] == 0 and footer["mode"] == "on" and footer["max_events"] == 100000
    assert footer["open_spans"] == 0 and footer["truncated"] is False

    # Runner replay: full header incl. type/schema, host spans with explicit ns,
    # VBA events remapped onto a reserved id block.
    h = _header()
    w = TraceWriter(tmp_path / "t.jsonl", h)
    stage = w.begin("host.stage", "profile_trace_on", ns=4000)
    base = w.reserve_ids(len(events)) - 1
    for ev in events:
        ev = dict(ev, id=ev["id"] + base)
        if ev["type"] in ("B", "M"):
            ev["parent"] = ev["parent"] + base if ev["parent"] else stage
        w.append_event(ev)
    w.end(stage, ns=6000)
    w.close()
    res = validate(tmp_path / "t.jsonl")
    # Only rule 8 fires (this fixture has no calc.full/recalc); structure is valid.
    assert all(e.startswith("[8]") for e in res.errors), res.errors


def test_non_ascii_attr_value_allowed_but_formula_chars_rejected(tmp_path):
    # Sheet names may be non-ASCII; '=' and '"' (formula text, literals) may not.
    w = _writer(tmp_path)
    sid = w.begin("host.stage", "warmup", attrs={"sheet": "Mod\u00e8le"})
    w.end(sid)
    with pytest.raises(TraceError):
        w.begin("host.stage", "open_workbook", attrs={"note": "=SUM(A1)"})
    with pytest.raises(TraceError):
        w.begin("host.stage", "open_workbook", attrs={"note": 'say "hi"'})



@pytest.mark.parametrize("value", ["a\n", "a\r", "a\r\n", "x\u00a0y", "tab\tx", "=A1", 'q"'])
def test_attr_pattern_is_fully_anchored(value):
    assert trace.check_attr("note", value) is not None
    assert not trace.ATTR_VALUE_RE.match(value)


@pytest.mark.parametrize("value", ["Mod\u00e8le", "\u65e5\u672c", "Sheet 1!$A$1:$B$2", "'Q1 (draft)'!A1"])
def test_attr_pattern_allows_unicode_word_chars(value):
    assert trace.check_attr("sheet", value) is None


def test_attr_pattern_matches_design_regex_exactly():
    design = re.compile(r"^[\w .:$!'#\-/()\[\],@]*$")
    for c in range(0x3000):
        ch = chr(c)
        assert (trace.check_attr("note", ch) is None) == bool(design.fullmatch(ch)), (c, ch)


def test_validator_rejects_newline_in_attr(tmp_path):
    ev = _events()
    ev[7]["attrs"]["note"] = "ok\n"
    assert _errors_with(validate(_write(tmp_path, ev)), "[6]")


# ---------------------------------------------------------------- rules 9, 10 and robustness


@pytest.mark.parametrize("name", ['Sheet1!B2 SUM(A1:A9)*1.07+"x"', 'a"b', "a*b", "a+b", "a^b", "a<b", "a>b", "=A1", "tab\tname", "x\n"])
def test_rule10_bad_names_rejected(tmp_path, name):
    assert trace.check_name(name) is not None
    w = _writer(tmp_path)
    with pytest.raises(TraceError):
        w.begin("host.stage", name)
    ev = _events()
    ev[7]["name"] = name
    assert _errors_with(validate(_write(tmp_path, ev)), "[10]")


@pytest.mark.parametrize("name", ["Model", "Mod\u00e8le", "'Q1 & Q2'!$A$1:$B$2", "Growth %", "h:0123456789", "Macro_Refresh.LoadInputs", "G0001"])
def test_rule10_structural_names_allowed(name):
    assert trace.check_name(name) is None
    assert trace.NAME_RE.fullmatch(name)


def test_rule10_name_length_and_hash_always_passes():
    assert trace.check_name("a" * 256) is None
    assert trace.check_name("a" * 257) is not None
    assert trace.check_name(redact_name('bad "name" *+', "salt")) is None


@pytest.mark.parametrize("ns", [True, 1.5, 1000.0, -1, 2**63, "1000", None])
def test_rule9_ns_must_be_bounded_int(tmp_path, ns):
    ev = _events()
    ev[2]["ns"] = ns
    res = validate(_write(tmp_path, ev))
    assert any("0 <= ns < 2**63" in e for e in _errors_with(res, "[9]"))


def test_writer_rejects_out_of_range_ns(tmp_path):
    w = _writer(tmp_path)
    for bad in (-1, 2**63, 5.0, True):
        with pytest.raises(TraceError):
            w.begin("run.pass", "XSP_RunPass", clock="vba", pass_=1, ns=bad)


@pytest.mark.parametrize(
    "key,value",
    [
        ("host", []),
        ("host", "x"),
        ("redaction", "hashed"),
        ("redaction", {"names": "plain"}),
        ("clock", {"vba_offset_ns": 0, "vba_offset_uncertainty_ns": "lots"}),
        ("clock", {"vba_offset_ns": 1.5}),
        ("source", {"kind": "real", "label": {"x": 1}}),
        ("limits", {"max_events": "many"}),
        ("expect", {"vba.proc": [["x"]]}),
        ("expect", {"vba.proc": [{"a": 1}]}),
        ("expect", {"vba.proc": "Macro"}),
        ("expect", {"not.a.kind": ["x"]}),
        ("expect", ["vba.proc"]),
        ("run_id", 7),
        ("mode", ["trace-on"]),
    ],
)
def test_rule9_header_structure(tmp_path, key, value):
    res = validate(_write(tmp_path, header=_header(**{key: value})))
    assert not res.ok
    assert _errors_with(res, "[9]") or _errors_with(res, "[1]")


def test_rule9_root_level_calc_rejected(tmp_path):
    ev = _events()
    # calc.full outside the host stage and any run.pass.
    ev += [_b(50, "calc.full", "Application.CalculateFull", 0, 0, 6000, pass_=0), _e(50, 6100)]
    res = validate(_write(tmp_path, ev))
    assert any("not inside any run.pass" in e for e in _errors_with(res, "[9]"))


def test_rule9_calc_claiming_other_pass_rejected(tmp_path):
    ev = _events()
    ev[7]["pass"] = 9
    res = validate(_write(tmp_path, ev))
    assert any("differs from enclosing run.pass" in e for e in _errors_with(res, "[9]"))


def test_rule9_offset_sign_checked(tmp_path):
    offset = 10_000_000
    h = _header(clock={"vba_offset_ns": offset, "vba_offset_uncertainty_ns": 50})
    ev = _events()
    for e in ev:
        if e["clock"] == "vba":
            e["ns"] += offset
    assert validate(_write(tmp_path, ev, header=h)).ok
    h["clock"]["vba_offset_ns"] = -offset
    res = validate(_write(tmp_path, ev, header=h))
    assert any("clock alignment" in e for e in _errors_with(res, "[9]"))


def test_rule9_offset_uncertainty_tolerated(tmp_path):
    ev = _events()
    ev[0]["ns"] = 1040  # host stage begins 40 ns after its VBA child run.pass (1000)
    h = _header()
    h["clock"]["vba_offset_uncertainty_ns"] = 100
    assert validate(_write(tmp_path, ev, header=h)).ok
    h["clock"]["vba_offset_uncertainty_ns"] = 10
    assert _errors_with(validate(_write(tmp_path, ev, header=h)), "[9]")


def test_strict_int_fields(tmp_path):
    ev = _events()
    ev[1]["parent"] = True  # True == 1 in Python; must still be rejected
    assert _errors_with(validate(_write(tmp_path, ev)), "[4]")
    ev = _events()
    footer = {"type": "footer", "events": float(len(ev)), "dropped": 0, "truncated": False}
    assert _errors_with(validate(_write(tmp_path, ev, footer=footer)), "[5]")


def test_rule9_header_limits_enforced(tmp_path):
    ev = _events()
    res = validate(_write(tmp_path, ev, header=_header(limits={"max_events": 5, "max_bytes": 10**9})))
    assert any("max_events" in e for e in _errors_with(res, "[9]"))
    res = validate(_write(tmp_path, ev, header=_header(limits={"max_events": 1000, "max_bytes": 500})))
    assert any("max_bytes" in e for e in _errors_with(res, "[9]"))


def test_writer_limits_header_reflects_writer(tmp_path):
    h = _header(limits={"max_events": 1, "max_bytes": 1})
    del h["type"], h["schema"]
    w = TraceWriter(tmp_path / "w.jsonl", h, max_events=7, max_bytes=10**6)
    assert w.header["limits"] == {"max_events": 7, "max_bytes": 10**6}
    with pytest.raises(TraceError):
        w.update_header({"limits": {"max_events": 10**9}})


def test_writer_header_growth_after_events_stays_in_bounds(tmp_path):
    w = _writer(tmp_path, max_bytes=2000)
    for i in range(5):
        w.marker("m", ns=i)
    w.update_header({"host": {"os": "x" * 900}})
    footer = w.close()
    size = (tmp_path / "w.jsonl").stat().st_size
    assert size <= 2000
    assert footer["events"] + footer["dropped"] == 5
    if footer["dropped"]:
        assert footer["truncated"] is True
    with pytest.raises(TraceError):
        w2 = _writer(tmp_path / "sub", max_bytes=2000)
        w2.update_header({"host": {"os": "x" * 5000}})


def test_writer_many_small_events_respect_max_bytes(tmp_path):
    w = _writer(tmp_path, max_bytes=2000)
    for i in range(50):
        w.marker("m", ns=i)
    footer = w.close()
    assert (tmp_path / "w.jsonl").stat().st_size <= 2000
    assert footer["truncated"] is True and footer["events"] + footer["dropped"] == 50


def _fuzz_values():
    return [None, True, False, 0, -1, 1.5, 2**64, 10**400, "", "x", "=A1", [], [1], {}, {"a": 1}, [["x"]]]


def test_validate_never_raises_on_mutations(tmp_path):
    import random

    from xlsprint.synth import synthetic_traces

    on, _off = synthetic_traces(tmp_path / "src", seed=11, sheets=2, passes=2)
    lines = [json.loads(l) for l in on.read_text().splitlines()]
    rng = random.Random(1234)
    values = _fuzz_values()
    count = 0
    for trial in range(300):
        doc = json.loads(json.dumps(lines))
        for _ in range(rng.randint(1, 4)):
            i = rng.randrange(len(doc))
            op = rng.randrange(7)
            if op == 0 and isinstance(doc[i], dict) and doc[i]:
                k = rng.choice(sorted(doc[i]))
                doc[i][k] = rng.choice(values)
            elif op == 1 and isinstance(doc[i], dict) and doc[i]:
                del doc[i][rng.choice(sorted(doc[i]))]
            elif op == 2:
                doc[i] = rng.choice(values)
            elif op == 3:
                del doc[i]
                if not doc:
                    doc = [{}]
            elif op == 4:
                doc.insert(i, json.loads(json.dumps(doc[rng.randrange(len(doc))])))
            elif op == 5 and isinstance(doc[0], dict):
                sub = rng.choice(["host", "clock", "source", "redaction", "limits", "expect"])
                doc[0][sub] = rng.choice(values)
            elif op == 6 and isinstance(doc[i], dict) and isinstance(doc[i].get("attrs"), dict):
                doc[i]["attrs"][rng.choice(["note", "sheet", "bogus"])] = rng.choice(values)
        p = tmp_path / ("f%d.jsonl" % trial)
        text = "\n".join(json.dumps(d) for d in doc) + "\n"
        if trial % 25 == 0:
            text = text[: rng.randrange(len(text))]  # truncated file
        p.write_text(text)
        res = validate(p)  # must not raise
        assert isinstance(res.ok, bool)
        count += not res.ok
        from xlsprint import analysis

        analysis.load_run(p)  # tolerant loading must not raise either
    assert count > 250  # nearly every mutation is caught
