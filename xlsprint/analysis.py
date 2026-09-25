"""Span tree and statistics over a validated trace.

Measurement classes (DESIGN.md): a span's ``dur_ns`` is DIRECT (``measured``):
the wall time between its balanced B and E events. The median, quartiles and
min-max over repeats of the *same* operation are still ``measured`` (always
shown with ``n``). Arithmetic that combines different measurements is
DERIVED: self time (duration minus the union of child intervals),
interval-union totals over several spans, shares, volatility ratio, and
overhead (on minus off).

Totals over several spans are always the length of the union of their
intervals, never the sum of durations (DESIGN.md, "Double-counting rule").
All times are in the host timebase: VBA-clock readings are shifted by
``-header.clock.vba_offset_ns``.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Tuple

from . import trace as _trace
from .trace import ValidationResult

# Kinds that are drill-down measurements (ranked as hotspots and phases).
CALC_KINDS = ("calc.full", "calc.fullrebuild", "calc.recalc", "calc.sheet", "calc.range", "calc.name", "calc.group")
MEASURED_KINDS = CALC_KINDS + ("vba.proc",)
SHEET_CHILD_KINDS = ("calc.range", "calc.name", "calc.group")
OVERHEAD_KINDS = ("run.pass", "calc.full", "calc.recalc")

GROUP_MEASUREMENT = "measured (isolated Range.Calculate)"
UNASSIGNED_SHEET = "(no sheet)"
MAX_CHILDREN_PER_SHEET = 200  # drill-down rows per sheet; the rest are counted, not listed  # drill-down bucket for children whose sheet is unknown


@dataclass
class Span:
    id: int
    parent: int
    pass_: int
    depth: int
    kind: str
    name: str
    clock: str
    start_ns: int  # host timebase
    end_ns: int  # host timebase
    status: str
    attrs: dict
    children: List[int] = field(default_factory=list)

    @property
    def dur_ns(self) -> int:
        """DIRECT: end minus begin on the span's own clock (the offset cancels)."""
        return self.end_ns - self.start_ns


@dataclass
class Run:
    header: dict
    spans: Dict[int, Span]
    markers: List[dict]
    validation: ValidationResult

    @property
    def source_kind(self) -> str:
        src = self.header.get("source")
        kind = src.get("kind") if isinstance(src, dict) else None
        return kind if _trace._in(kind, _trace.SOURCE_KINDS) else "real"

    @property
    def mode(self) -> str:
        mode = self.header.get("mode")
        return mode if _trace._in(mode, _trace.MODES) else "trace-on"


# --------------------------------------------------------------------------
# Loading


def _vba_offset(header: dict) -> int:
    clock = header.get("clock")
    off = clock.get("vba_offset_ns") if isinstance(clock, dict) else None
    return off if isinstance(off, int) and not isinstance(off, bool) else 0


def load_run(path) -> Run:
    """Validate and load a trace into a span tree.

    Loading is tolerant so an invalid trace can still be shown with an
    INVALID banner: spans without a matching E are left out, and lines that
    are not well-formed events are skipped. ``run.validation`` says whether
    any of that happened.
    """
    validation = _trace.validate(path)
    objs, _errs, _ = _trace._parse_lines(path)
    header, body, _footer = _trace._split(objs)
    header = header or {}
    offset = _vba_offset(header)

    def host_ns(ev: dict) -> Optional[int]:
        ns = ev.get("ns")
        if not _trace._is_ns(ns):
            return None
        return ns - offset if ev.get("clock") == "vba" else ns

    def as_int(v: Any, default: int = 0) -> int:
        return v if _trace._is_int(v) else default

    def as_str(v: Any, default: str = "") -> str:
        return v if isinstance(v, str) else default

    def as_attrs(v: Any) -> dict:
        return dict(v) if isinstance(v, dict) else {}

    # Malformed values are coerced or skipped here; run.validation records them.
    spans: Dict[int, Span] = {}
    order: Dict[int, int] = {}  # span id -> line index of its B (parents must precede children)
    markers: List[dict] = []
    open_b: Dict[int, Tuple[int, dict]] = {}
    for idx, (_no, ev) in enumerate(body):
        if not isinstance(ev, dict) or not _trace._is_int(ev.get("id")):
            continue
        typ, eid = ev.get("type"), ev["id"]
        if typ == "B" and eid not in open_b and eid not in spans:
            open_b[eid] = (idx, ev)
        elif typ == "E" and eid in open_b:
            b_idx, b = open_b.pop(eid)
            start, end = host_ns(b), host_ns(ev)
            if start is None or end is None or end < start:
                continue
            attrs = as_attrs(b.get("attrs"))
            attrs.update(as_attrs(ev.get("attrs")))
            status = ev.get("status")
            spans[eid] = Span(
                id=eid,
                parent=as_int(b.get("parent")),
                pass_=as_int(b.get("pass")),
                depth=as_int(b.get("depth")),
                kind=as_str(b.get("kind")),
                name=as_str(b.get("name")),
                clock=as_str(b.get("clock"), "host"),
                start_ns=start,
                end_ns=end,
                status=status if _trace._in(status, _trace.STATUSES) else "error",
                attrs=attrs,
            )
            order[eid] = b_idx
        elif typ == "M":
            ns = host_ns(ev)
            if ns is not None:
                markers.append(
                    {
                        "id": eid,
                        "parent": as_int(ev.get("parent")),
                        "pass": as_int(ev.get("pass")),
                        "name": as_str(ev.get("name")),
                        "clock": as_str(ev.get("clock"), "host"),
                        "ns": ns,
                        "attrs": as_attrs(ev.get("attrs")),
                    }
                )
    for s in sorted(spans.values(), key=lambda s: (s.start_ns, s.depth)):
        if s.parent in spans and order[s.parent] < order[s.id]:
            spans[s.parent].children.append(s.id)
    return Run(header=header, spans=spans, markers=markers, validation=validation)


# --------------------------------------------------------------------------
# Primitives


def union_ns(intervals: Iterable[Tuple[int, int]]) -> int:
    """Length of the union of half-open intervals [start, end). Empty ones are ignored."""
    ivs = sorted((a, b) for a, b in intervals if b > a)
    total = 0
    cur_a = cur_b = None
    for a, b in ivs:
        if cur_b is None or a > cur_b:
            if cur_b is not None:
                total += cur_b - cur_a
            cur_a, cur_b = a, b
        elif b > cur_b:
            cur_b = b
    if cur_b is not None:
        total += cur_b - cur_a
    return total


def self_time_ns(run: Run, span_id: int) -> int:
    """DERIVED: duration minus |union of child intervals| clipped to the span, floored at 0."""
    s = run.spans[span_id]
    kids = []
    for cid in s.children:
        c = run.spans[cid]
        kids.append((max(c.start_ns, s.start_ns), min(c.end_ns, s.end_ns)))
    return max(0, s.dur_ns - union_ns(kids))


def _num(x: float):
    return int(x) if float(x).is_integer() else x


def _quantile(sorted_vals: List[int], q: float) -> float:
    # Linear interpolation between closest ranks (same as numpy's default).
    if len(sorted_vals) == 1:
        return sorted_vals[0]
    pos = (len(sorted_vals) - 1) * q
    lo = math.floor(pos)
    hi = min(lo + 1, len(sorted_vals) - 1)
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (pos - lo)


def summarize(values: List[int]) -> dict:
    """{n, median, min, max, p25, p75} of values; with n == 0 the statistics are None.

    Over repeats of the same operation (e.g. one sheet's Worksheet.Calculate
    across passes) these are ``measured`` per DESIGN.md; over derived
    values (self times, differences) they stay ``derived``. Quartiles use
    linear interpolation between closest ranks.
    """
    vals = sorted(values)
    if not vals:
        return {"n": 0, "median": None, "min": None, "max": None, "p25": None, "p75": None}
    return {
        "n": len(vals),
        "median": _num(_quantile(vals, 0.5)),
        "min": vals[0],
        "max": vals[-1],
        "p25": _num(_quantile(vals, 0.25)),
        "p75": _num(_quantile(vals, 0.75)),
    }


def _iqr(summary: dict) -> float:
    return (summary["p75"] - summary["p25"]) if summary["n"] else 0


def _origin(run: Run) -> int:
    starts = [s.start_ns for s in run.spans.values()] + [m["ns"] for m in run.markers]
    return min(starts) if starts else 0


def _by_start(spans: Iterable[Span]) -> List[Span]:
    return sorted(spans, key=lambda s: (s.start_ns, s.depth, s.id))


def _in_its_pass(run: Run, span: Span) -> bool:
    """True if a run.pass ancestor carries the span's pass number.

    Measured kinds outside their pass only occur in invalid traces (rule 9);
    views skip them so a forced (--allow-invalid) report is not skewed.
    """
    seen = set()
    p = span.parent
    while p in run.spans and p not in seen:
        seen.add(p)
        anc = run.spans[p]
        if anc.kind == "run.pass":
            return anc.pass_ == span.pass_
        p = anc.parent
    return False


def _of_kind(run: Run, kind: str) -> List[Span]:
    spans = (s for s in run.spans.values() if s.kind == kind)
    if kind in MEASURED_KINDS:
        spans = (s for s in spans if _in_its_pass(run, s))
    return _by_start(spans)


def _measured_spans(run: Run, kinds: Tuple[str, ...] = MEASURED_KINDS) -> List[Span]:
    return [s for s in run.spans.values() if s.kind in kinds and _in_its_pass(run, s)]


# --------------------------------------------------------------------------
# Views


def host_stages(run: Run) -> List[dict]:
    """Host stages in start order: {name, start_rel_ns, dur_ns, status}.

    ``dur_ns`` is DIRECT (perf_counter_ns). ``start_rel_ns`` is relative to the
    earliest event in the run.
    """
    t0 = _origin(run)
    return [
        {"name": s.name, "start_rel_ns": s.start_ns - t0, "dur_ns": s.dur_ns, "status": s.status}
        for s in _of_kind(run, "host.stage")
    ]


def timeline(run: Run, *, max_rows: int = 2000) -> List[dict]:
    """All spans in start order: {id, kind, name, pass, depth, start_rel_ns, dur_ns, status}.

    Bounded: when there are more than ``max_rows`` spans, the shallowest and
    then longest ``max_rows`` are kept and a final row ``{"truncated": k}``
    says how many were omitted. VBA-span starts are only as accurate as the
    header's ``vba_offset_uncertainty_ns`` relative to host spans.
    """
    spans = list(run.spans.values())
    omitted = 0
    if len(spans) > max_rows:
        keep = sorted(spans, key=lambda s: (s.depth, -s.dur_ns, s.start_ns))[:max_rows]
        omitted = len(spans) - len(keep)
        spans = keep
    t0 = _origin(run)
    rows: List[dict] = [
        {
            "id": s.id,
            "kind": s.kind,
            "name": s.name,
            "pass": s.pass_,
            "depth": s.depth,
            "start_rel_ns": s.start_ns - t0,
            "dur_ns": s.dur_ns,
            "status": s.status,
        }
        for s in _by_start(spans)
    ]
    if omitted:
        rows.append({"truncated": omitted})
    return rows


def _descendants(run: Run, span_id: int) -> List[Span]:
    out, todo = [], list(run.spans[span_id].children)
    while todo:
        c = run.spans[todo.pop()]
        out.append(c)
        todo.extend(c.children)
    return out


def passes(run: Run) -> List[dict]:
    """One row per run.pass: {pass, id, status, dur_ns, by_kind_union_ns}.

    ``dur_ns`` is DIRECT. ``by_kind_union_ns[kind]`` is DERIVED: the union of
    that kind's descendant intervals within the pass (nested spans of one kind,
    e.g. vba.proc inside vba.proc, are not counted twice).
    """
    rows = []
    for rp in sorted(_of_kind(run, "run.pass"), key=lambda s: (s.pass_, s.start_ns)):
        per_kind: Dict[str, List[Tuple[int, int]]] = {}
        for d in _descendants(run, rp.id):
            per_kind.setdefault(d.kind, []).append((d.start_ns, d.end_ns))
        rows.append(
            {
                "pass": rp.pass_,
                "id": rp.id,
                "status": rp.status,
                "dur_ns": rp.dur_ns,
                "by_kind_union_ns": {k: union_ns(v) for k, v in sorted(per_kind.items())},
            }
        )
    return rows


def _sheet_of(span: Span) -> Optional[str]:
    """Sheet a range/name/group span belongs to: attrs.sheet, else the prefix of an A1 address."""
    sheet = span.attrs.get("sheet")
    if isinstance(sheet, str) and sheet:
        return sheet
    if span.kind == "calc.sheet":
        return span.name
    for cand in (span.attrs.get("address"), span.name):
        if isinstance(cand, str) and "!" in cand:
            sh = cand.rsplit("!", 1)[0]
            if len(sh) >= 2 and sh[0] == sh[-1] == "'":
                sh = sh[1:-1].replace("''", "'")
            return sh
    return None


def drilldown(run: Run) -> dict:
    """Microsoft's workbook -> sheet -> range drill-down, summarized across passes.

    - ``workbook.full`` / ``workbook.recalc``: summaries of DIRECT calc.full /
      calc.recalc durations (one sample per span, normally one per pass).
    - ``workbook.volatility_ratio``: DERIVED, median recalc / median full.
      None when either is missing or median full is 0.
    - ``sheets[].recalc``: summary of Worksheet.Calculate durations.
    - ``sheets[].share_of_recalc_sum``: DERIVED, this sheet's median divided by
      the sum of all sheet medians. The sheet calcs are separate sequential
      calls; this is a share of that sum, not of the workbook recalc.
    - ``sheets[].children``: calc.range / calc.name / calc.group keyed by
      (kind, name) across passes. A range and a name covering the same cells
      stay separate rows; nothing here adds them up. At most
      MAX_CHILDREN_PER_SHEET rows (highest median first) are kept;
      ``children_truncated`` counts the rest (0 when none were dropped).
    """
    full = summarize([s.dur_ns for s in _of_kind(run, "calc.full")])
    recalc = summarize([s.dur_ns for s in _of_kind(run, "calc.recalc")])
    ratio = None
    if full["n"] and recalc["n"] and full["median"]:
        ratio = recalc["median"] / full["median"]
    workbook = {"full": full, "recalc": recalc, "volatility_ratio": ratio}
    rebuild = _of_kind(run, "calc.fullrebuild")
    if rebuild:
        workbook["fullrebuild"] = summarize([s.dur_ns for s in rebuild])

    sheet_durs: Dict[str, List[int]] = {}
    for s in _of_kind(run, "calc.sheet"):
        sheet_durs.setdefault(s.name, []).append(s.dur_ns)
    child_spans: Dict[Tuple[Optional[str], str, str], List[Span]] = {}
    for s in _by_start(_measured_spans(run, SHEET_CHILD_KINDS)):
        child_spans.setdefault((_sheet_of(s), s.kind, s.name), []).append(s)

    by_sheet: Dict[Optional[str], List[Tuple[str, str, List[Span]]]] = {}
    for (sheet, kind, cname), group in child_spans.items():
        by_sheet.setdefault(sheet, []).append((kind, cname, group))
    sheet_names: List[Optional[str]] = list(sheet_durs)
    sheet_names += [n for n in by_sheet if n is not None and n not in sheet_durs]
    if None in by_sheet:
        sheet_names.append(None)  # rendered as UNASSIGNED_SHEET

    summaries = {n: summarize(sheet_durs.get(n, [])) for n in sheet_names}
    denom = sum(v["median"] for v in summaries.values() if v["n"])
    sheets = []
    for name in sheet_names:
        children = []
        for kind, cname, group in by_sheet.get(name, []):
            first = group[0]
            children.append(
                {
                    "kind": kind,
                    "name": cname,
                    "summary": summarize([g.dur_ns for g in group]),
                    "address": first.attrs.get("address") if isinstance(first.attrs.get("address"), str) else (cname if kind == "calc.range" else None),
                    "measurement": GROUP_MEASUREMENT if kind == "calc.group" else "measured",
                }
            )
        children.sort(key=lambda c: (-(c["summary"]["median"] or 0), c["kind"], c["name"]))
        truncated = max(0, len(children) - MAX_CHILDREN_PER_SHEET)
        children = children[:MAX_CHILDREN_PER_SHEET]
        summ = summaries[name]
        sheets.append(
            {
                "name": UNASSIGNED_SHEET if name is None else name,
                "recalc": summ,
                "share_of_recalc_sum": (summ["median"] / denom) if (summ["n"] and denom) else 0.0,
                "children": children,
                "children_truncated": truncated,
            }
        )
    sheets.sort(key=lambda s: -(s["recalc"]["median"] or 0))
    return {"workbook": workbook, "sheets": sheets}


def phase_breakdown(run: Run) -> List[dict]:
    """Per kind: {kind, count, union_ns_median_per_pass, share_of_pass, passes}.

    DERIVED. For each pass, the kind's spans are unioned (never summed); the
    median of those per-pass unions is reported, taken over the passes in
    which the kind occurs (``passes``). ``share_of_pass`` is that median
    divided by the median run.pass duration (0.0 without passes).
    """
    rows = passes(run)
    pass_median = summarize([r["dur_ns"] for r in rows])["median"] or 0
    per_kind: Dict[str, List[int]] = {}
    for r in rows:
        for kind, ns in r["by_kind_union_ns"].items():
            if kind in MEASURED_KINDS:
                per_kind.setdefault(kind, []).append(ns)
    counts: Dict[str, int] = {}
    for s in run.spans.values():
        if s.kind in per_kind:
            counts[s.kind] = counts.get(s.kind, 0) + 1
    out = []
    for kind, vals in per_kind.items():
        med = summarize(vals)["median"]
        out.append(
            {
                "kind": kind,
                "count": counts.get(kind, 0),
                "passes": len(vals),
                "union_ns_median_per_pass": med,
                "share_of_pass": (med / pass_median) if pass_median else 0.0,
            }
        )
    out.sort(key=lambda r: -r["union_ns_median_per_pass"])
    return out


def hotspots(run: Run, top: int = 25) -> List[dict]:
    """Rank (kind, name) by median self time across passes.

    Each row: {kind, name, sheet, address, median_ns, median_self_ns, n, min_ns,
    max_ns, measurement}. ``median_ns`` summarizes DIRECT durations.
    ``measurement`` is "measured" when no sample had children (self time is
    then the direct duration) and "derived" when self time subtracts a child
    union. calc.group rows carry ``note`` = isolated Range.Calculate; their
    time is not a share of a normal recalc. Rows are never added together.
    """
    groups: Dict[Tuple[str, str], List[Span]] = {}
    for s in _measured_spans(run):
        groups.setdefault((s.kind, s.name), []).append(s)
    rows = []
    for (kind, name), spans in groups.items():
        durs = summarize([s.dur_ns for s in spans])
        selfs = summarize([self_time_ns(run, s.id) for s in spans])
        row = {
            "kind": kind,
            "name": name,
            "sheet": _sheet_of(spans[0]) if kind in SHEET_CHILD_KINDS + ("calc.sheet",) else None,
            "address": spans[0].attrs.get("address") if isinstance(spans[0].attrs.get("address"), str)
            else (name if kind == "calc.range" and "!" in name else None),
            "median_ns": durs["median"],
            "median_self_ns": selfs["median"],
            "n": durs["n"],
            "min_ns": durs["min"],
            "max_ns": durs["max"],
            "measurement": "derived" if any(s.children for s in spans) else "measured",
        }
        if kind == "calc.group":
            row["note"] = GROUP_MEASUREMENT
        rows.append(row)
    rows.sort(key=lambda r: (-r["median_self_ns"], -r["median_ns"], r["kind"], r["name"]))
    return rows[:top]


def _ok_per_pass(run: Run, kind: str) -> Tuple[Dict[int, Span], List[int]]:
    """({pass: first status-ok span of ``kind``}, [passes excluded as not ok]).

    A pass counts only if its run.pass span has status ok, and (for calc
    kinds) the span itself is ok.
    """
    rp_ok = {s.pass_: s.status == "ok" for s in _of_kind(run, "run.pass") if s.pass_ >= 1}
    firsts: Dict[int, Span] = {}
    bad: List[int] = []
    for s in _of_kind(run, kind):
        p = s.pass_
        if p < 1 or p in firsts or p in bad:
            continue
        if s.status == "ok" and rp_ok.get(p, False):
            firsts[p] = s
        else:
            bad.append(p)
    return firsts, sorted(bad)


def _overhead_row(kind: str, name: str, **values) -> dict:
    row = {
        "kind": kind,
        "name": name,
        "median_on_ns": None,
        "median_off_ns": None,
        "diff_ns": None,
        "diff_min_ns": None,
        "diff_max_ns": None,
        "iqr_on_ns": None,
        "iqr_off_ns": None,
        "resolvable": False,
        "pairs": 0,
        "dropped_pairs": 0,
        "dropped_passes": [],
        "available": True,
        "unavailable_reason": None,
    }
    row.update(values)
    return row


def overhead(on: Run, off: Run) -> List[dict]:
    """Instrumentation overhead for run.pass, calc.full and calc.recalc (DERIVED).

    Pairs are matched by pass number: trace-off pass k with trace-on pass k.
    Always returns one row per kind. A row with ``available: False`` carries
    a plain-text ``unavailable_reason`` and None for every number: either the
    runs' header ``run_id`` differ (all rows), or that kind has no pairable
    pass. Only
    passes whose run.pass (and, for calc kinds, the span itself) has status
    ok are paired; ``dropped_passes`` lists pass numbers present in either
    run that could not be paired, ``dropped_pairs`` counts them.

    Per row: ``diff_ns`` = median of the per-pair differences (on - off);
    ``diff_min_ns``/``diff_max_ns`` = their range; ``median_on_ns`` /
    ``median_off_ns`` and the interquartile spreads ``iqr_on_ns`` /
    ``iqr_off_ns`` describe the paired samples of each run. ``resolvable`` is
    True only with >= 3 pairs and when [diff_min_ns, diff_max_ns] excludes 0;
    otherwise the report says "not resolvable (below run-to-run noise)".
    """
    on_id, off_id = on.header.get("run_id"), off.header.get("run_id")
    same_run = isinstance(on_id, str) and bool(on_id) and on_id == off_id
    rows = []
    for kind in OVERHEAD_KINDS:
        if not same_run:
            rows.append(_overhead_row(kind, kind, available=False,
                                      unavailable_reason="trace-on and trace-off run_id differ (%r vs %r)" % (on_id, off_id)))
            continue
        on_p, on_bad = _ok_per_pass(on, kind)
        off_p, off_bad = _ok_per_pass(off, kind)
        paired = sorted(set(on_p) & set(off_p))
        dropped = sorted((set(on_p) | set(off_p) | set(on_bad) | set(off_bad)) - set(paired))
        names = [on_p[p].name for p in paired] or [kind]
        name = max(sorted(set(names)), key=names.count)
        if not paired:
            rows.append(_overhead_row(kind, name, dropped_pairs=len(dropped), dropped_passes=dropped, available=False,
                                      unavailable_reason="no pass has a status-ok %s in both runs" % kind))
            continue
        on_v = [on_p[p].dur_ns for p in paired]
        off_v = [off_p[p].dur_ns for p in paired]
        diffs = [a - b for a, b in zip(on_v, off_v)]
        s_on, s_off, s_diff = summarize(on_v), summarize(off_v), summarize(diffs)
        rows.append(
            _overhead_row(
                kind,
                name,
                median_on_ns=s_on["median"],
                median_off_ns=s_off["median"],
                diff_ns=s_diff["median"],
                diff_min_ns=s_diff["min"],
                diff_max_ns=s_diff["max"],
                iqr_on_ns=_num(_iqr(s_on)),
                iqr_off_ns=_num(_iqr(s_off)),
                resolvable=len(paired) >= 3 and (s_diff["min"] > 0 or s_diff["max"] < 0),
                pairs=len(paired),
                dropped_pairs=len(dropped),
                dropped_passes=dropped,
            )
        )
    return rows
