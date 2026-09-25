"""Deterministic synthetic traces for tests, demos, and ``selftest``.

The generated timings are made up (seeded random numbers). They follow the
shape of a real profile: host stages, then R passes of Microsoft's drill-down
plan on the VBA clock, and a matching trace-off run that makes exactly the
same calls but records only run.pass, calc.full and calc.recalc. Headers say
``source.kind = "synthetic"`` so the report shows the SYNTHETIC banner.
"""

from __future__ import annotations

import random
from pathlib import Path
from typing import List, Tuple

from . import __version__
from .trace import TraceWriter

MS = 1_000_000
US = 1_000

_SHEET_NAMES = ["Inputs", "Model", "Report", "Lookup", "Summary", "Scenarios", "Checks", "Outputs"]
_HOST_BASE_NS = 1_000_000_000_000
_GAP_NS = 20 * US  # idle time between consecutive calls inside a pass
_EVENT_COST_NS = 3 * US  # per-event instrumentation cost when tracing is on


def _sheet_names(n: int) -> List[str]:
    return [_SHEET_NAMES[i] if i < len(_SHEET_NAMES) else "Sheet%d" % (i + 1) for i in range(n)]


class _Clock:
    """Simulated host clock; the VBA clock is host + offset."""

    def __init__(self, offset: int):
        self.t = _HOST_BASE_NS
        self.offset = offset

    def host(self) -> int:
        return self.t

    def vba(self) -> int:
        return self.t + self.offset

    def advance(self, ns: float) -> None:
        self.t += max(1, int(ns))


def _header(mode: str, run_id: str, offset: int, sheets: int, expect=None) -> dict:
    h = {
        "run_id": run_id,
        "mode": mode,
        "created_utc": "2026-01-01T00:00:00Z",
        "tool_version": __version__,
        "host": {
            "os": "synthetic",
            "python": "synthetic",
            "excel_version": "synthetic",
            "excel_build": "synthetic",
            "bitness": 64,
            "threads": 8,
            "calc_mode_original": "automatic",
        },
        "clock": {
            "host": "perf_counter_ns",
            "vba": "MicroTimer/QPC",
            "vba_offset_ns": offset,
            "vba_offset_uncertainty_ns": 4000,
        },
        "source": {"kind": "synthetic", "label": "synthetic:generated"},
        "redaction": {"names": "clear"},
    }
    if expect:
        h["expect"] = expect
    return h


def _stage(w: TraceWriter, clk: _Clock, name: str, dur_ns: float) -> None:
    sid = w.begin("host.stage", name, ns=clk.host())
    clk.advance(dur_ns)
    w.end(sid, ns=clk.host())


def _jitter(rng: random.Random, sd: float = 0.04) -> float:
    return min(1.3, max(0.7, rng.gauss(1.0, sd)))


def _write_pass(w: TraceWriter, clk: _Clock, rng: random.Random, p: int, sheets: List[str], cost: dict, record_all: bool) -> None:
    """One drill-down pass. Timing always advances; only recording differs by mode."""

    extra = _EVENT_COST_NS if record_all else 0

    def span(kind, name, dur, attrs=None, recorded=True, body=None):
        clk.advance(_GAP_NS)
        if recorded:
            sid = w.begin(kind, name, clock="vba", pass_=p, attrs=attrs, ns=clk.vba())
            clk.advance(extra)
        clk.advance(dur)
        if body:
            body()
        if recorded:
            clk.advance(extra)
            w.end(sid, ns=clk.vba())

    rp = w.begin("run.pass", "XSP_RunPass", clock="vba", pass_=p, attrs={"repeat": p, "calc_mode": "manual"}, ns=clk.vba())
    sheet_total = sum(cost[s] for s in sheets)
    span("calc.full", "Application.CalculateFull", sheet_total * 2.8 * _jitter(rng), {"method": "Application.CalculateFull"})
    span("calc.recalc", "Application.Calculate", sheet_total * 0.35 * _jitter(rng), {"method": "Application.Calculate"})
    if record_all:
        clk.advance(_GAP_NS)
        w.marker("AfterCalculate", clock="vba", pass_=p, ns=clk.vba())
    for s in sheets:
        span("calc.sheet", s, cost[s] * _jitter(rng), {"method": "Worksheet.Calculate", "sheet": s}, record_all)
    for i, s in enumerate(sheets):
        half = 100 + 50 * i
        for addr, share in (("A1:D%d" % half, 0.45), ("E1:H%d" % half, 0.35)):
            full_addr = "%s!%s" % (s, addr)
            span("calc.range", full_addr, cost[s] * share * _jitter(rng),
                 {"method": "Range.Calculate", "sheet": s, "address": full_addr, "areas": 1, "iteration": False},
                 record_all)
    name_sheet = sheets[0]
    span("calc.name", "InputBlock", cost[name_sheet] * 0.4 * _jitter(rng),
         {"method": "Range.Calculate", "sheet": name_sheet, "address": "%s!A1:D100" % name_sheet}, record_all)
    grp_sheet = sheets[1] if len(sheets) > 1 else sheets[0]
    span("calc.group", "G0001", cost[grp_sheet] * 0.3 * _jitter(rng),
         {"method": "Range.Calculate", "sheet": grp_sheet, "group": "G0001", "cells": 400, "areas": 1}, record_all)

    def macro_body():
        span("vba.proc", "LoadInputs", 6 * MS * _jitter(rng), {"note": "user span"}, record_all)
        clk.advance(2 * MS * _jitter(rng))

    span("vba.proc", "Macro_Refresh", 3 * MS * _jitter(rng), {"method": "Application.Run"}, record_all, macro_body)
    clk.advance(_GAP_NS)
    w.end(rp, ns=clk.vba())


def synthetic_traces(out_dir, *, seed=0, sheets=3, passes=5) -> Tuple[Path, Path]:
    """Write ``trace-on.jsonl`` and ``trace-off.jsonl`` into ``out_dir``; return both paths.

    Deterministic for a given (seed, sheets, passes). Both traces validate.
    """
    if sheets < 1 or passes < 1:
        raise ValueError("sheets and passes must be >= 1")
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    rng = random.Random(seed)
    names = _sheet_names(sheets)
    cost = {s: rng.uniform(5, 40) * MS for s in names}
    offset = rng.randrange(10**12, 10**13)
    # One profile run produces both files; overhead pairing requires the same run_id.
    run_id = "%032x" % rng.getrandbits(128)
    on_path, off_path = out / "trace-on.jsonl", out / "trace-off.jsonl"

    # Same per-pass random stream for on and off, so off makes the same calls
    # with the same underlying timings; on adds per-event instrumentation cost.
    pass_seeds = [rng.getrandbits(32) for _ in range(passes)]

    on = TraceWriter(on_path, _header("trace-on", run_id, offset, sheets,
                                      expect={"calc.sheet": ["*"], "vba.proc": ["Macro_Refresh", "LoadInputs"]}))
    clk = _Clock(offset)
    for stage, dur in (("prepare_copy", 80 * MS), ("launch_excel", 1500 * MS), ("open_workbook", 900 * MS),
                       ("instrument", 120 * MS), ("warmup", 400 * MS)):
        _stage(on, clk, stage, dur * _jitter(rng))
    prof = on.begin("host.stage", "profile_trace_on", ns=clk.host())
    for p in range(1, passes + 1):
        clk.advance(5 * MS)
        _write_pass(on, clk, random.Random(pass_seeds[p - 1]), p, names, cost, record_all=True)
    clk.advance(5 * MS)
    on.end(prof, ns=clk.host())
    for stage, dur in (("collect_trace", 60 * MS), ("close_excel", 700 * MS), ("verify_output", 40 * MS)):
        _stage(on, clk, stage, dur * _jitter(rng))
    on.close()

    off = TraceWriter(off_path, _header("trace-off", run_id, offset, sheets))
    clk = _Clock(offset)
    prof = off.begin("host.stage", "profile_trace_off", ns=clk.host())
    for p in range(1, passes + 1):
        clk.advance(5 * MS)
        _write_pass(off, clk, random.Random(pass_seeds[p - 1]), p, names, cost, record_all=False)
    clk.advance(5 * MS)
    off.end(prof, ns=clk.host())
    off.close()
    return on_path, off_path
