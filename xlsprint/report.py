"""Standalone HTML report.

``build_model`` turns analysed runs into a plain view-model dict,
``render_html`` turns that dict into one self-contained HTML page (inline CSS,
inline vanilla JS, inline SVG, no network), and ``render`` wires the two to
trace files on disk.

Every number on the page is wrapped by ``_num`` and so carries a measurement
class badge (measured / derived / static / n/a). ``render_html`` reads only the
fields it knows about, and the model embedded as JSON is pruned to the same
allow-list, so unexpected fields in a model never reach the page.
"""

from __future__ import annotations

import html
import json
import math
import re
from pathlib import Path
from typing import Any, Iterable

from . import __version__

REPORT_SCHEMA = "xlsprint.report/1"
TIMELINE_MAX_ROWS = 2000
MAX_FORMULA_GROUPS = 500
MAX_PROC_NODES = 300
MAX_PROC_DEPTH = 8
MAX_DD_SHEETS = 500
MAX_DD_CHILDREN_PER_SHEET = 200
MAX_DD_CHILDREN_TOTAL = 2000
MAX_LIST_ROWS = 500
MAX_PASS_ROWS = 1000


class ReportError(Exception):
    """The report cannot be produced (for example: invalid trace, fail closed)."""


# ---------------------------------------------------------------------------
# Measurement classes
# ---------------------------------------------------------------------------

MEASURED = "measured"
DERIVED = "derived"
STATIC = "static"
NA = "na"

_BADGE_TEXT = {MEASURED: "measured", DERIVED: "derived", STATIC: "static", NA: "n/a"}

# Formulas for derived values, shown next to them and in hover titles.
F_VOLATILITY = "volatility ratio = median recalc ÷ median full"
F_SHARE = "share = sheet median recalc ÷ sum of all sheet median recalcs"
F_SELF = "self time = span duration − |union of child spans|, median over passes"
F_UNION = "union time = length of the union of all spans of this kind in a pass, median over passes"
F_SHARE_PASS = "share of pass = union time ÷ run.pass duration"
F_OVERHEAD = ("overhead = median over pairs of (on − off); resolvable when ≥3 pairs and "
              "min–max of paired differences excludes 0")
F_OVERHEAD_RANGE = "min–max of per-pair differences (on − off)"
F_IQR = "IQR = p75 − p25 of the run's durations"
F_OFFSET = "offset = VBA clock − host midpoint (h0+h1)/2 at the clock probe"
F_OFFSET_UNC = "uncertainty = ±(h1 − h0)/2 of the clock probe"

STAGE_ORDER = [
    "prepare_copy", "launch_excel", "open_workbook", "instrument", "warmup",
    "profile_trace_off", "profile_trace_on", "collect_trace", "close_excel",
    "verify_output", "render_report",
]
STAGE_LABELS = {
    "prepare_copy": "prepare copy",
    "launch_excel": "launch Excel",
    "open_workbook": "open workbook",
    "instrument": "instrument",
    "warmup": "warmup",
    "profile_trace_off": "trace-off passes",
    "profile_trace_on": "trace-on passes",
    "collect_trace": "collect trace",
    "close_excel": "close Excel",
    "verify_output": "verify output",
    "render_report": "render report",
}
KIND_LABELS = {
    "host.stage": "host stage",
    "run.pass": "pass",
    "calc.full": "CalculateFull",
    "calc.fullrebuild": "CalculateFullRebuild",
    "calc.recalc": "Calculate",
    "calc.sheet": "Worksheet.Calculate",
    "calc.range": "Range.Calculate (range)",
    "calc.name": "Range.Calculate (name)",
    "calc.group": "Range.Calculate (group, isolated)",
    "vba.proc": "VBA procedure",
    "marker": "marker",
}
KIND_CSS = {
    "host.stage": "k-host", "run.pass": "k-pass", "calc.full": "k-full",
    "calc.fullrebuild": "k-full", "calc.recalc": "k-recalc", "calc.sheet": "k-sheet",
    "calc.range": "k-range", "calc.name": "k-name", "calc.group": "k-group",
    "vba.proc": "k-proc",
}
PHASE_ORDER = ["run.pass", "calc.full", "calc.fullrebuild", "calc.recalc", "calc.sheet",
               "calc.range", "calc.name", "calc.group", "vba.proc"]

FORMULA_TOTALS = [
    ("formula_cells", "formula cells"), ("groups", "groups"), ("volatile_cells", "volatile cells"),
    ("udf_cells", "UDF cells"), ("single_thread_cells", "single-threaded cells"),
    ("array_cells", "array cells"), ("dynamic_array_cells", "dynamic-array cells"),
    ("cross_sheet_cells", "cross-sheet cells"), ("external_ref_cells", "external-ref cells"),
    ("whole_column_ref_cells", "whole-column-ref cells"), ("whole_row_ref_cells", "whole-row-ref cells"),
    ("data_tables", "data tables"), ("names", "defined names"), ("range_names", "range names"),
]
FORMULA_TOTALS = dict(FORMULA_TOTALS)

GROUP_CAVEAT = "isolated Range.Calculate — not a share of normal recalc"

# Defensive filters for strings copied from formulas.json. Formula text always
# contains "=" or other characters these patterns exclude.
_SAFE_TOKEN_RE = re.compile(r"^[A-Za-z0-9_.:]{1,64}$")
_SAFE_ADDR_RE = re.compile(r"^[A-Za-z0-9_ .:$!'#\-/()\[\],@]{1,256}$")


# ---------------------------------------------------------------------------
# Model
# ---------------------------------------------------------------------------

def _analysis():
    from . import analysis  # imported lazily; tests may substitute this getter
    return analysis


def _get(obj: Any, name: str, default: Any = None) -> Any:
    if isinstance(obj, dict):
        return obj.get(name, default)
    return getattr(obj, name, default)


def _dict(v: Any) -> dict:
    return v if isinstance(v, dict) else {}


def _scalar(v: Any) -> Any:
    return v if v is None or isinstance(v, (str, int, float, bool)) else None


def _validation_dict(run: Any, label: str) -> dict:
    v = _get(run, "validation")
    ok = bool(_get(v, "ok", False)) if v is not None else False
    errors = [f"{label}: {e}" for e in (_get(v, "errors", None) or [])]
    warnings = [f"{label}: {w}" for w in (_get(v, "warnings", None) or [])]
    if v is None:
        errors.append(f"{label}: no validation result available")
    return {"ok": ok, "errors": errors, "warnings": warnings}


def _proc_tree(run: Any, A: Any) -> list:
    """Merge vba.proc spans (and everything nested below them) across passes by name path."""
    spans = _get(run, "spans", None) or {}
    nodes: dict = {}
    roots: list = []

    def visit(sid: int, parent_path: tuple, depth: int) -> None:
        sp = spans[sid]
        path = parent_path + ((sp.kind, sp.name),)
        node = nodes.get(path)
        if node is None:
            if len(nodes) >= MAX_PROC_NODES:
                return
            node = nodes[path] = {"kind": sp.kind, "name": sp.name, "durs": [], "selfs": [], "kids": []}
            (nodes[parent_path]["kids"] if parent_path else roots).append(path)
        node["durs"].append(sp.dur_ns)
        node["selfs"].append(A.self_time_ns(run, sid))
        if depth < MAX_PROC_DEPTH:
            for c in sp.children:
                if c in spans:
                    visit(c, path, depth + 1)

    for sid in sorted(spans):
        sp = spans[sid]
        if sp.kind != "vba.proc":
            continue
        parent = spans.get(sp.parent)
        if parent is not None and parent.kind == "vba.proc":
            continue
        visit(sid, (), 0)

    def out(path: tuple) -> dict:
        n = nodes[path]
        return {"kind": n["kind"], "name": n["name"], "summary": A.summarize(n["durs"]),
                "self_summary": A.summarize(n["selfs"]), "children": [out(k) for k in n["kids"]]}

    return [out(p) for p in roots]


def _formulas_model(formulas: dict | None) -> dict | None:
    if not formulas:
        return None

    def toks(values: Any) -> list:
        return [str(v) for v in (values or []) if isinstance(v, str) and _SAFE_TOKEN_RE.match(v)]

    def addr(v: Any) -> Any:
        return v if isinstance(v, str) and _SAFE_ADDR_RE.match(v) else None

    def structural(v: Any) -> Any:
        return v if isinstance(v, str) and len(v) <= 256 and "=" not in v and '"' not in v else None

    groups_in = [g for g in (formulas.get("groups") or []) if isinstance(g, dict)]
    groups_in.sort(key=lambda g: (-(g.get("cells") or 0), str(g.get("group", ""))))
    groups = []
    for g in groups_in[:MAX_FORMULA_GROUPS]:
        groups.append({
            "group": addr(g.get("group")),
            "fingerprint": g.get("fingerprint") if isinstance(g.get("fingerprint"), str)
            and re.match(r"^[0-9a-fA-F]{1,64}$", g["fingerprint"]) else None,
            "sheet": g.get("sheet") if isinstance(g.get("sheet"), str) else None,
            "cells": g.get("cells"),
            "areas": [a for a in (addr(x) for x in (g.get("areas") or [])) if a],
            "areas_total": g.get("areas_total"),
            "lambda_calls": g.get("lambda_calls"),
            "length_bucket": addr(g.get("length_bucket")),
            "functions": toks(g.get("functions")),
            "udfs": g.get("udfs"),
            "volatile": bool(g.get("volatile")),
            "volatile_functions": toks(g.get("volatile_functions")),
            "single_threaded": bool(g.get("single_threaded")),
            "single_threaded_functions": toks(g.get("single_threaded_functions")),
            "array": bool(g.get("array")),
            "dynamic_array": bool(g.get("dynamic_array")),
            "cross_sheet": bool(g.get("cross_sheet")),
            "external_ref": bool(g.get("external_ref")),
            "whole_column_ref": bool(g.get("whole_column_ref")),
            "whole_row_ref": bool(g.get("whole_row_ref")),
            "annotations": [],
        })
    sheets = [{"sheet": s.get("sheet") if isinstance(s.get("sheet"), str) else None,
               "formula_cells": s.get("formula_cells"), "used_range": addr(s.get("used_range")),
               "data_tables": s.get("data_tables")}
              for s in (formulas.get("sheets") or []) if isinstance(s, dict)]
    totals_in = formulas.get("totals") or {}
    totals = {k: totals_in.get(k) for k in FORMULA_TOTALS if k in totals_in}
    # Limitations are tool-authored prose; anything that looks like formula text is dropped.
    limitations = [x for x in (formulas.get("limitations") or [])
                   if isinstance(x, str) and len(x) <= 400 and "=" not in x and '"' not in x][:30]
    redaction = formulas.get("redaction") if isinstance(formulas.get("redaction"), dict) else {}
    names = []
    for item in (formulas.get("names") or []):
        if not isinstance(item, dict):
            continue
        names.append({
            "name": structural(item.get("name")),
            "scope": structural(item.get("scope")),
            "sheet": structural(item.get("sheet")),
            "refers_to_range": addr(item.get("refers_to_range")),
            "is_range": bool(item.get("is_range")),
            "hidden": bool(item.get("hidden")),
            "formula_cells": item.get("formula_cells"),
            "annotations": [],
        })
    semantics_in = formulas.get("semantics") if isinstance(formulas.get("semantics"), dict) else {}
    semantic_annotations = []
    for item in (semantics_in.get("annotations") or [])[:500]:
        if not isinstance(item, dict):
            continue
        annotation = {
            key: value for key, limit in (("id", 80), ("label", 120), ("intent", 600),
                                          ("category", 80), ("source", 120))
            if isinstance((value := item.get(key)), str) and len(value) <= limit
        }
        if not all(isinstance(annotation.get(key), str) for key in ("id", "label", "intent")):
            continue
        binding = item.get("binding") if item.get("binding") in ("bound", "unmatched") else "unmatched"
        targets = []
        for target in (item.get("targets") or [])[:100]:
            if not isinstance(target, dict):
                continue
            clean = {key: structural(target.get(key)) for key in
                     ("target", "kind", "name", "scope", "sheet", "address", "span_kind", "group")
                     if target.get(key) is not None}
            if isinstance(target.get("span_kinds"), list):
                clean["span_kinds"] = [v for v in target["span_kinds"] if isinstance(v, str) and len(v) <= 40][:10]
            if clean:
                targets.append(clean)
        semantic_annotations.append({**annotation, "binding": binding, "targets": targets,
                                     "timed_match_count": 0})
    return {
        "schema": formulas.get("schema") if isinstance(formulas.get("schema"), str) else None,
        "workbook_sha256_prefix": addr(formulas.get("workbook_sha256_prefix")),
        "sheets": sheets,
        "groups": groups,
        "names": names[:1000],
        "groups_total": len(groups_in),
        "names_count": len(formulas.get("names") or []),
        "totals": totals,
        "limitations": limitations,
        "redaction_names": redaction.get("names") if isinstance(redaction.get("names"), str) else None,
        "semantics": {"schema": "xlsprint.semantics/1",
                      "workbook_bound": bool(semantics_in.get("workbook_bound")),
                      "annotations": semantic_annotations,
                      "unbound_count": int(semantics_in.get("unbound_count") or 0)},
    }


def _same_address(left: Any, right: Any) -> bool:
    if not isinstance(left, str) or not isinstance(right, str):
        return False
    clean = lambda value: re.sub(r"[\s$]", "", value).upper()
    return clean(left.rsplit("!", 1)[-1]) == clean(right.rsplit("!", 1)[-1])


def _target_matches(row: dict, target: dict, *, sheet: str | None = None,
                    formula_group: bool = False, defined_name: bool = False) -> bool:
    target_type = target.get("target")
    if target_type == "defined_name":
        if defined_name:
            return (row.get("name") == target.get("name")
                    and (target.get("scope") is None or row.get("scope") == target.get("scope")))
        return (row.get("kind") == target.get("span_kind") == "calc.name"
                and row.get("name") == target.get("name")
                and (target.get("sheet") is None or (row.get("sheet") or sheet) == target.get("sheet"))
                and (target.get("address") is None or _same_address(row.get("address"), target.get("address"))))
    if target_type == "formula_group":
        if formula_group:
            return row.get("group") == target.get("group") \
                and (target.get("sheet") is None or row.get("sheet") == target.get("sheet"))
        return (row.get("kind") == target.get("span_kind") == "calc.group"
                and row.get("name") == target.get("group")
                and (target.get("sheet") is None or (row.get("sheet") or sheet) == target.get("sheet")))
    kinds = target.get("span_kinds") or ([target.get("span_kind")] if target.get("span_kind") else [])
    kind = target.get("kind") or target.get("span_kind")
    if target_type == "span" and kind and row.get("kind") != kind:
        return False
    if kinds and row.get("kind") not in kinds:
        return False
    if target_type == "range" or (target_type == "span" and target.get("address")):
        row_sheet = row.get("sheet") if row.get("sheet") is not None else sheet
        if target.get("sheet") is not None and row_sheet != target.get("sheet"):
            return False
        if not _same_address(row.get("address"), target.get("address")):
            return False
    if target.get("name") is not None and row.get("name") != target.get("name"):
        return False
    if target.get("sheet") is not None and target_type == "span" and "address" not in target:
        row_sheet = row.get("sheet") if row.get("sheet") is not None else sheet
        if row_sheet != target.get("sheet"):
            return False
    return bool(kinds or target_type == "span")


def _apply_semantics(model: dict) -> None:
    """Attach labels to matching report rows without changing any measurements."""
    formulas = model.get("formulas") or {}
    semantics = formulas.get("semantics") or {}
    annotations = semantics.get("annotations") or []
    if not annotations:
        return
    timed_matches: dict[str, set[str]] = {a["id"]: set() for a in annotations}

    def add(row: dict, *, sheet: str | None = None, formula_group: bool = False,
            defined_name: bool = False, timed: bool = False, row_key: str = "") -> None:
        if not isinstance(row, dict):
            return
        for annotation in annotations:
            for target in annotation.get("targets") or []:
                if _target_matches(row, target, sheet=sheet, formula_group=formula_group,
                                   defined_name=defined_name):
                    row.setdefault("annotations", []).append({key: annotation[key] for key in
                                      ("id", "label", "intent", "category", "source") if key in annotation})
                    if timed:
                        timed_matches[annotation["id"]].add(row_key)
                    break

    for row in formulas.get("names") or []:
        add(row, defined_name=True)
    for row in formulas.get("groups") or []:
        add(row, formula_group=True)
        for annotation in annotations:
            if any(target.get("target") == "range" and target.get("sheet") == row.get("sheet")
                   and any(_same_address(area, target.get("address")) for area in (row.get("areas") or []))
                   for target in annotation.get("targets") or []):
                row.setdefault("annotations", []).append({key: annotation[key] for key in
                                  ("id", "label", "intent", "category", "source") if key in annotation})
    for sheet_row in (model.get("drilldown") or {}).get("sheets") or []:
        for row in sheet_row.get("children") or []:
            add(row, sheet=sheet_row.get("name"), timed=True,
                row_key=f"drilldown:{sheet_row.get('name')}:{row.get('kind')}:{row.get('name')}:{row.get('address')}")
    for row in model.get("hotspots") or []:
        add(row, timed=True, row_key=f"hotspot:{row.get('kind')}:{row.get('sheet')}:{row.get('name')}")
    for row in (model.get("timeline") or {}).get("rows") or []:
        add(row, timed=True, row_key=f"timeline:{row.get('id')}")

    def add_proc(rows: list, prefix: str = "") -> None:
        for index, row in enumerate(rows):
            key = f"{prefix}/{index}:{row.get('name')}"
            add(row, timed=True, row_key=f"proc:{key}")
            add_proc(row.get("children") or [], key)

    add_proc(model.get("vba_procs") or [])
    for annotation in annotations:
        annotation["timed_match_count"] = len(timed_matches[annotation["id"]])


def build_model(on: Any, off: Any = None, formulas: dict | None = None) -> dict:
    """Pure view-model for ``render_html``, built from the analysis API."""
    A = _analysis()
    header = _dict(_get(on, "header", None))
    host = _dict(header.get("host"))
    clock = _dict(header.get("clock"))
    source = _dict(header.get("source"))
    redaction = _dict(header.get("redaction"))

    v_on = _validation_dict(on, "trace-on")
    validation = {"ok": v_on["ok"], "errors": list(v_on["errors"]),
                  "warnings": list(v_on["warnings"]), "allow_invalid": False}
    if off is not None:
        v_off = _validation_dict(off, "trace-off")
        validation["ok"] = validation["ok"] and v_off["ok"]
        validation["errors"] += v_off["errors"]
        validation["warnings"] += v_off["warnings"]

    section_errors: dict = {}

    def section(name: str, fn, default):
        # A valid trace that breaks analysis is a bug and must surface. An
        # invalid trace (rendered only with allow_invalid) may not analyse.
        if validation["ok"]:
            return fn()
        try:
            return fn()
        except Exception as exc:  # noqa: BLE001
            section_errors[name] = f"{type(exc).__name__}: {exc}"
            return default

    passes_on = section("passes", lambda: A.passes(on), [])
    passes_off = section("passes_off", lambda: A.passes(off), []) if off is not None else []

    tl_rows = section("timeline", lambda: A.timeline(on, max_rows=TIMELINE_MAX_ROWS), [])
    truncated = 0
    rows = []
    for r in tl_rows:
        if "truncated" in r and "id" not in r:
            truncated = int(r.get("truncated") or 0)
        else:
            rows.append(r)

    if off is None:
        overhead = {"available": False, "rows": [],
                    "reason": "No trace-off run was supplied, so instrumentation overhead is unavailable."}
    else:
        oh = section("overhead", lambda: A.overhead(on, off), None)
        if oh is None:
            overhead = {"available": False, "rows": [],
                        "reason": "Overhead could not be computed from these traces."}
        elif isinstance(oh, dict):
            # analysis may report the whole comparison as unavailable, e.g. run_id mismatch.
            available = oh.get("available", True) is not False
            reason = oh.get("reason") if isinstance(oh.get("reason"), str) else None
            overhead = {"available": available, "rows": list(oh.get("rows") or []) if available else [],
                        "reason": reason or (None if available else "Overhead unavailable for these traces.")}
        else:
            overhead = {"available": True, "rows": list(oh), "reason": None}

    meta = {
        "label": _scalar(source.get("label")),
        "source_kind": _scalar(source.get("kind")),
        "run_id": _scalar(header.get("run_id")),
        "mode": _scalar(header.get("mode")),
        "created_utc": _scalar(header.get("created_utc")),
        "trace_tool_version": _scalar(header.get("tool_version")),
        "host": {k: _scalar(host.get(k)) for k in
                 ("os", "python", "excel_version", "excel_build", "bitness", "threads", "calc_mode_original")},
        "clock": {k: _scalar(clock.get(k)) for k in
                  ("host", "vba", "vba_offset_ns", "vba_offset_uncertainty_ns")},
        "redaction_names": _scalar(redaction.get("names")),
        "repeats_on": len(passes_on),
        "repeats_off": len(passes_off) if off is not None else None,
        "has_off": off is not None,
    }

    model = _apply_caps({
        "schema": REPORT_SCHEMA,
        "tool_version": __version__,
        "meta": meta,
        "validation": validation,
        "host_stages": section("host_stages", lambda: A.host_stages(on), []),
        "passes": passes_on,
        "timeline": {"rows": rows, "truncated": truncated, "max_rows": TIMELINE_MAX_ROWS},
        "drilldown": section("drilldown", lambda: A.drilldown(on), {}),
        "phases": section("phases", lambda: A.phase_breakdown(on), []),
        "vba_procs": section("vba_procs", lambda: _proc_tree(on, A), []),
        "hotspots": section("hotspots", lambda: A.hotspots(on), []),
        "overhead": overhead,
        "formulas": _formulas_model(formulas),
        "section_errors": section_errors,
    })
    _apply_semantics(model)
    return model


def _apply_caps(model: dict) -> dict:
    """Bound every list that grows with the workbook. Idempotent; returns a shallow-copied model."""
    m = dict(model)
    tl = m.get("timeline")
    if isinstance(tl, dict) and isinstance(tl.get("rows"), list) and len(tl["rows"]) > TIMELINE_MAX_ROWS:
        extra = len(tl["rows"]) - TIMELINE_MAX_ROWS
        prev = tl.get("truncated") if _is_num(tl.get("truncated")) else 0
        m["timeline"] = dict(tl, rows=tl["rows"][:TIMELINE_MAX_ROWS], truncated=int(prev) + extra)
    dd = m.get("drilldown")
    if isinstance(dd, dict) and isinstance(dd.get("sheets"), list):
        budget = MAX_DD_CHILDREN_TOTAL
        sheets = []
        for sh in dd["sheets"][:MAX_DD_SHEETS]:
            if not isinstance(sh, dict):
                continue
            kids = sh.get("children") if isinstance(sh.get("children"), list) else []
            keep = max(0, min(len(kids), MAX_DD_CHILDREN_PER_SHEET, budget))
            budget -= keep
            prev = sh.get("children_truncated") if _is_num(sh.get("children_truncated")) else 0
            sheets.append(dict(sh, children=kids[:keep], children_truncated=int(prev) + len(kids) - keep))
        m["drilldown"] = dict(dd, sheets=sheets,
                              sheets_truncated=max(0, len(dd["sheets"]) - MAX_DD_SHEETS))
    for key, cap in (("hotspots", MAX_LIST_ROWS), ("host_stages", MAX_LIST_ROWS), ("phases", MAX_LIST_ROWS),
                     ("passes", MAX_PASS_ROWS), ("vba_procs", MAX_PROC_NODES)):
        if isinstance(m.get(key), list) and len(m[key]) > cap:
            m[key] = m[key][:cap]
    oh = m.get("overhead")
    if isinstance(oh, dict) and isinstance(oh.get("rows"), list) and len(oh["rows"]) > MAX_LIST_ROWS:
        m["overhead"] = dict(oh, rows=oh["rows"][:MAX_LIST_ROWS])
    fm = m.get("formulas")
    if isinstance(fm, dict) and isinstance(fm.get("groups"), list) and len(fm["groups"]) > MAX_FORMULA_GROUPS:
        m["formulas"] = dict(fm, groups=fm["groups"][:MAX_FORMULA_GROUPS])
    return m


# ---------------------------------------------------------------------------
# Allow-list for the embedded JSON copy of the model
# ---------------------------------------------------------------------------

S = "scalar"          # a str/int/float/bool/None
SL = "scalar-list"    # a list of scalars


class _Map:
    """A dict with arbitrary (string) keys whose values follow ``spec``."""

    def __init__(self, spec: Any) -> None:
        self.spec = spec


_SUMMARY = {"n": S, "median": S, "min": S, "max": S, "p25": S, "p75": S}
_ANNOTATION = {"id": S, "label": S, "intent": S, "category": S, "source": S}
_ANNOTATIONS = [_ANNOTATION]
_SEMANTIC_TARGET = {"target": S, "kind": S, "name": S, "scope": S, "sheet": S,
                    "address": S, "span_kind": S, "group": S, "span_kinds": SL}
_PROC: dict = {"kind": S, "name": S, "summary": _SUMMARY, "self_summary": _SUMMARY,
               "annotations": _ANNOTATIONS}
_PROC["children"] = [_PROC]
_MODEL_SPEC = {
    "schema": S, "tool_version": S,
    "meta": {"label": S, "source_kind": S, "run_id": S, "mode": S, "created_utc": S,
             "trace_tool_version": S,
             "host": {"os": S, "python": S, "excel_version": S, "excel_build": S, "bitness": S,
                      "threads": S, "calc_mode_original": S},
             "clock": {"host": S, "vba": S, "vba_offset_ns": S, "vba_offset_uncertainty_ns": S},
             "redaction_names": S, "repeats_on": S, "repeats_off": S, "has_off": S},
    "validation": {"ok": S, "errors": SL, "warnings": SL, "allow_invalid": S},
    "host_stages": [{"name": S, "start_rel_ns": S, "dur_ns": S, "status": S}],
    "passes": [{"pass": S, "id": S, "status": S, "dur_ns": S, "by_kind_union_ns": _Map(S)}],
    "timeline": {"rows": [{"id": S, "kind": S, "name": S, "pass": S, "depth": S,
                           "start_rel_ns": S, "dur_ns": S, "status": S,
                           "annotations": _ANNOTATIONS}],
                 "truncated": S, "max_rows": S},
    "drilldown": {"workbook": {"full": _SUMMARY, "recalc": _SUMMARY, "fullrebuild": _SUMMARY,
                               "volatility_ratio": S},
                  "sheets": [{"name": S, "recalc": _SUMMARY, "share_of_recalc_sum": S,
                              "children": [{"kind": S, "name": S, "summary": _SUMMARY, "address": S,
                                           "measurement": S, "annotations": _ANNOTATIONS}],
                              "children_truncated": S}],
                  "sheets_truncated": S},
    "phases": [{"kind": S, "count": S, "passes": S, "union_ns_median_per_pass": S, "share_of_pass": S}],
    "vba_procs": [_PROC],
    "hotspots": [{"kind": S, "name": S, "sheet": S, "median_ns": S, "median_self_ns": S, "n": S,
                  "address": S, "min_ns": S, "max_ns": S, "measurement": S, "note": S,
                  "annotations": _ANNOTATIONS}],
    "overhead": {"available": S, "reason": S,
                 "rows": [{"kind": S, "name": S, "median_on_ns": S, "median_off_ns": S, "diff_ns": S,
                           "diff_min_ns": S, "diff_max_ns": S, "iqr_on_ns": S, "iqr_off_ns": S,
                           "resolvable": S, "pairs": S, "reason": S, "available": S,
                           "unavailable_reason": S, "dropped_pairs": S, "dropped_passes": SL}]},
    "formulas": {"schema": S, "workbook_sha256_prefix": S,
                 "groups_total": S, "names_count": S,
                 "limitations": SL, "redaction_names": S,
                 "semantics": {"schema": S, "workbook_bound": S, "unbound_count": S,
                               "annotations": [{"id": S, "label": S, "intent": S, "category": S,
                                               "source": S, "binding": S, "timed_match_count": S,
                                               "targets": [_SEMANTIC_TARGET]}]},
                 "sheets": [{"sheet": S, "formula_cells": S, "used_range": S, "data_tables": S}],
                 "groups": [{"group": S, "fingerprint": S, "sheet": S, "cells": S, "areas": SL, "areas_total": S,
                             "lambda_calls": S, "length_bucket": S, "whole_row_ref": S,
                             "functions": SL, "udfs": S, "volatile": S, "volatile_functions": SL,
                             "single_threaded": S, "single_threaded_functions": SL, "array": S,
                             "dynamic_array": S, "cross_sheet": S, "external_ref": S,
                             "whole_column_ref": S, "annotations": _ANNOTATIONS}],
                 "names": [{"name": S, "scope": S, "sheet": S, "refers_to_range": S,
                            "is_range": S, "hidden": S, "formula_cells": S,
                            "annotations": _ANNOTATIONS}],
                 "totals": {k: S for k in FORMULA_TOTALS}},
    "section_errors": _Map(S),
}


def _is_scalar(v: Any) -> bool:
    return v is None or isinstance(v, (str, int, float, bool))


def _finite(v: Any) -> Any:
    if isinstance(v, float) and not math.isfinite(v):
        return None
    if isinstance(v, int) and not isinstance(v, bool) and not -_NUM_LIMIT < v < _NUM_LIMIT:
        return None
    return v


def _prune(value: Any, spec: Any) -> Any:
    if spec == S:
        return _finite(value) if _is_scalar(value) else None
    if spec == SL:
        return [_finite(v) for v in value if _is_scalar(v)] if isinstance(value, list) else []
    if isinstance(spec, _Map):
        if not isinstance(value, dict):
            return {}
        return {str(k): _prune(v, spec.spec) for k, v in value.items() if isinstance(k, str)}
    if isinstance(spec, list):
        return [_prune(v, spec[0]) for v in value] if isinstance(value, list) else []
    if isinstance(spec, dict):
        if not isinstance(value, dict):
            return None
        return {k: _prune(value[k], sub) for k, sub in spec.items() if k in value}
    return None


def _json_embed(model: dict) -> str:
    text = json.dumps(_prune(model, _MODEL_SPEC), sort_keys=True, ensure_ascii=True,
                      separators=(",", ":"), allow_nan=False, default=str)
    # Neutralise "</script>", "<!--" and friends inside the script element.
    return text.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")


# ---------------------------------------------------------------------------
# HTML helpers
# ---------------------------------------------------------------------------

def _e(v: Any) -> str:
    return html.escape("" if v is None else str(v), quote=True)


# Anything beyond ~31 years in ns (or a non-finite float) is not a real timing;
# it renders as n/a instead of overflowing float conversion.
_NUM_LIMIT = 10 ** 18


def _is_num(v: Any) -> bool:
    return (isinstance(v, (int, float)) and not isinstance(v, bool) and v == v
            and -_NUM_LIMIT < v < _NUM_LIMIT)


def fmt_ns(v: Any) -> str:
    if not _is_num(v):
        return "n/a"
    x = float(v)
    a = abs(x)
    if a >= 1e9:
        return f"{x / 1e9:.3f} s"
    if a >= 1e6:
        return f"{x / 1e6:.2f} ms"
    if a >= 1e3:
        return f"{x / 1e3:.1f} µs"
    return f"{x:.0f} ns"


def _fmt_count(v: Any) -> str:
    if isinstance(v, float) and not v.is_integer():
        return f"{v:.2f}"
    return f"{int(v):,}"


def _num(value: Any, cls: str, kind: str = "ns", formula: str | None = None) -> str:
    """A number with its measurement-class badge. Missing values render as n/a."""
    if not _is_num(value):
        return '<span class="num" data-class="na" title="n/a: not available">—</span>'
    if kind == "ns":
        text = fmt_ns(value)
    elif kind == "ratio":
        text = f"{float(value):.3f}"
    elif kind == "pct":
        text = f"{float(value) * 100:.1f}%"
    else:
        text = _fmt_count(value)
    title = _BADGE_TEXT[cls] + (f": {formula}" if formula else "")
    return f'<span class="num" data-class="{cls}" title="{_e(title)}">{_e(text)}</span>'


def _na(reason: str) -> str:
    return f'<span class="num" data-class="na" title="n/a: {_e(reason)}">n/a</span>'


def _badge(cls: str, extra: str = "") -> str:
    text = _BADGE_TEXT[cls] + (f" ({extra})" if extra else "")
    return f'<span class="badge" data-class="{cls}">{_e(text)}</span>'


def _range(summary: Any, cls: str = MEASURED) -> str:
    if not isinstance(summary, dict) or not _is_num(summary.get("min")):
        return _num(None, cls)
    return f'{_num(summary.get("min"), cls)} – {_num(summary.get("max"), cls)}'


def _sv(v: Any) -> str:
    """data-v attribute value used by client-side sorting."""
    if _is_num(v):
        return _e(repr(float(v)))
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, (int, float)):  # out-of-range or non-finite: sorts as missing
        return ""
    return _e("" if v is None else str(v).lower())


def _yes(v: Any) -> str:
    return '<span class="flag on">yes</span>' if v else '<span class="flag">no</span>'


def _kind(kind: Any) -> str:
    return f'<span class="kind {KIND_CSS.get(kind, "k-other")}">{_e(KIND_LABELS.get(kind, kind))}</span>'


def _section(sid: str, title: str, body: str, intro: str = "") -> str:
    intro_html = f'<p class="intro">{intro}</p>' if intro else ""
    return f'<section id="{sid}"><h2>{_e(title)}</h2>{intro_html}{body}</section>'


def _med(summary: Any) -> Any:
    return summary.get("median") if isinstance(summary, dict) else None


# ---------------------------------------------------------------------------
# Sections
# ---------------------------------------------------------------------------

def _header(model: dict) -> str:
    meta = model.get("meta") or {}
    host = meta.get("host") or {}
    clock = meta.get("clock") or {}
    val = model.get("validation") or {}
    out = []
    label = meta.get("label") or "(unlabelled run)"
    kind = meta.get("source_kind")
    out.append(f'<h1>{_e(label)}</h1>')
    if kind == "synthetic":
        out.append('<div class="banner synth" role="alert"><strong>SYNTHETIC — not a real-workbook '
                   'measurement.</strong> These timings come from generated test data.</div>')
    elif kind != "real":
        out.append(f'<div class="banner warn" role="alert"><strong>Unknown source kind '
                   f'({_e(kind)}).</strong> Treat these numbers as unverified.</div>')
    if not val.get("ok", False):
        errs = "".join(f"<li>{_e(x)}</li>" for x in (val.get("errors") or []))
        tail = ("Rendered with --allow-invalid. Timing numbers below may be wrong."
                if val.get("allow_invalid") else "Timing sections are withheld.")
        out.append(f'<div class="banner invalid" role="alert"><strong>INVALID TRACE.</strong> '
                   f'The trace failed validation. {_e(tail)}<ul>{errs}</ul></div>')
    warns = val.get("warnings") or []
    if warns:
        out.append('<details class="warnings"><summary>Validation warnings '
                   f'({_num(len(warns), STATIC, "count")})</summary><ul>'
                   + "".join(f"<li>{_e(w)}</li>" for w in warns) + "</ul></details>")

    rows = [
        ("Source", _e(kind or "unknown")),
        ("Run id", f"<code>{_e(meta.get('run_id'))}</code>"),
        ("Recorded (UTC)", _e(meta.get("created_utc"))),
        ("Excel", f"{_e(host.get('excel_version'))} build {_e(host.get('excel_build'))}, "
                  f"{_num(host.get('bitness'), STATIC, 'count')}-bit"),
        ("Calc threads", _num(host.get("threads"), STATIC, "count")),
        ("Original calc mode", _e(host.get("calc_mode_original"))
         + ' <span class="muted">(set to manual while profiling, restored before close)</span>'),
        ("Host", f"{_e(host.get('os'))}, Python {_e(host.get('python'))}"),
        ("Passes (repeats)", f"trace-on {_num(meta.get('repeats_on'), STATIC, 'count')}"
         + (f", trace-off {_num(meta.get('repeats_off'), STATIC, 'count')}" if meta.get("has_off")
            else ', trace-off <span class="muted">none</span>')),
        ("Clocks", f"host {_e(clock.get('host'))}; VBA {_e(clock.get('vba'))}"),
        ("VBA clock offset", f"{_num(clock.get('vba_offset_ns'), DERIVED, 'ns', F_OFFSET)} "
         f"± {_num(clock.get('vba_offset_uncertainty_ns'), DERIVED, 'ns', F_OFFSET_UNC)} "
         '<span class="muted">VBA and host spans are never compared more finely than this.</span>'),
        ("Name redaction", _e(meta.get("redaction_names"))),
    ]
    dl = "".join(f"<div><dt>{k}</dt><dd>{v}</dd></div>" for k, v in rows)
    return (f'<header id="run"><p class="eyebrow">XLSprint calculation profile</p>{"".join(out)}'
            f'<dl class="facts">{dl}</dl></header>')


def _legend() -> str:
    items = [
        (MEASURED, "Direct wall time of one Excel calc call or host stage, between balanced begin/end "
                   "events (MicroTimer or perf_counter). Includes all Excel work that call triggered. "
                   "Over repeats we show the median and the min–max of these direct timings."),
        (DERIVED, "Arithmetic on measured values: self time, interval-union totals, ratios, shares, "
                  "overhead. The formula is shown next to the value or in its hover title."),
        (STATIC, "Counts and settings read without timing: formula counts, group sizes, function "
                 "features, thread count, number of repeats."),
        (NA, "Not measurable by this method, or missing from this trace."),
    ]
    body = "".join(f'<div class="leg">{_badge(c)}<p>{_e(t)}</p></div>' for c, t in items)
    body += ('<p class="muted">Each number carries a small marker with its class: '
             '<span class="num" data-class="measured" title="example">1.0 ms</span> '
             '<span class="num" data-class="derived" title="example">0.25</span> '
             '<span class="num" data-class="static" title="example">12</span> '
             '<span class="num" data-class="na" title="example">—</span>. '
             'Hover a number to see its class and, for derived values, its formula.</p>')
    return _section("legend", "Measurement legend", f'<div class="legend">{body}</div>')


def _svg_text(x: float, y: float, text: str, cls: str = "rt-t") -> str:
    return f'<text x="{x:.0f}" y="{y:.0f}" class="{cls}" text-anchor="middle">{_e(text)}</text>'


def _svg_num(x: float, y: float, value: Any, cls: str, prefix: str = "") -> str:
    short = {MEASURED: "m", DERIVED: "d", STATIC: "s", NA: "n/a"}[cls if _is_num(value) else NA]
    shown = fmt_ns(value) if _is_num(value) else "not run"
    dc = cls if _is_num(value) else NA
    return (f'<text x="{x:.0f}" y="{y:.0f}" class="rt-n" text-anchor="middle">{_e(prefix)}'
            f'<tspan class="num" data-class="{dc}">{_e(shown)}</tspan>'
            f'<tspan class="rt-b" dx="3">{short}</tspan><title>{_e(_BADGE_TEXT[dc])}</title></text>')


def _clip(s: Any, n: int = 18) -> str:
    s = "" if s is None else str(s)
    return s if len(s) <= n else s[: n - 1] + "…"


def _route(model: dict) -> str:
    stages = [s for s in (model.get("host_stages") or []) if isinstance(s, dict)]
    stages.sort(key=lambda s: (s.get("start_rel_ns") if _is_num(s.get("start_rel_ns")) else 0,
                               STAGE_ORDER.index(s.get("name")) if s.get("name") in STAGE_ORDER else 99))
    phases = {p.get("kind"): p for p in (model.get("phases") or []) if isinstance(p, dict)}
    bw, bh, gap, x0 = 116, 58, 26, 12
    host_w = x0 * 2 + len(stages) * (bw + gap) - gap if stages else 0
    pw, pgap = 132, 22
    steps = [
        ("full", [("calc.full", "")]),
        ("recalc", [("calc.recalc", "")]),
        ("sheets", [("calc.sheet", "")]),
        ("ranges / names", [("calc.range", "range "), ("calc.name", "name ")]),
        ("groups", [("calc.group", "")]),
        ("macros", [("vba.proc", "")]),
    ]
    pass_w = x0 * 2 + len(steps) * (pw + pgap) - pgap
    width = max(host_w, pass_w, 320)
    height = 272
    parts = [f'<svg class="route" viewBox="0 0 {width} {height}" width="{width}" height="{height}" '
             'role="img" aria-labelledby="route-title"><title id="route-title">Host stages flow '
             'left to right; one profiling pass expands below</title>',
             '<defs><marker id="arr" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" '
             'markerHeight="7" orient="auto-start-reverse"><path d="M0,0L10,5L0,10z" class="rt-arrow"/>'
             '</marker></defs>']
    y = 26
    parts.append(f'<text x="{x0}" y="16" class="rt-cap">Host stages (measured, perf_counter)</text>')
    on_box = None
    for i, st in enumerate(stages):
        x = x0 + i * (bw + gap)
        name = st.get("name")
        status = st.get("status")
        bad = " rt-bad" if status not in (None, "ok") else ""
        hl = " rt-on" if name == "profile_trace_on" else ""
        parts.append(f'<g><rect x="{x}" y="{y}" width="{bw}" height="{bh}" rx="6" class="rt-box{hl}{bad}"/>'
                     + _svg_text(x + bw / 2, y + 22, _clip(STAGE_LABELS.get(name, name)))
                     + _svg_num(x + bw / 2, y + 42, st.get("dur_ns"), MEASURED)
                     + (f'<title>{_e(name)}: status {_e(status)}</title>' if bad else "") + '</g>')
        if i:
            parts.append(f'<line x1="{x - gap + 2}" y1="{y + bh / 2}" x2="{x - 3}" y2="{y + bh / 2}" '
                         'class="rt-line" marker-end="url(#arr)"/>')
        if name == "profile_trace_on":
            on_box = x
    if not stages:
        parts.append(f'<text x="{x0}" y="{y + 30}" class="rt-t" text-anchor="start">No host stages were '
                     'recorded in this trace.</text>')
    py = 150
    if on_box is not None:
        cx = on_box + bw / 2
        parts.append(f'<path d="M{cx - 20},{y + bh} L{x0},{py - 4} M{cx + 20},{y + bh} '
                     f'L{pass_w - x0},{py - 4}" class="rt-zoom"/>')
    parts.append(f'<text x="{x0}" y="{py + bh + 34}" class="rt-cap">One pass, in Microsoft\'s drill-down '
                 f'order — median union time per pass (derived)</text>')
    for i, (label, kinds) in enumerate(steps):
        x = x0 + i * (pw + pgap)
        present = any(k in phases for k, _ in kinds)
        parts.append(f'<g><rect x="{x}" y="{py}" width="{pw}" height="{bh + 16}" rx="6" '
                     f'class="rt-box{"" if present else " rt-missing"}"/>'
                     + _svg_text(x + pw / 2, py + 20, label))
        for j, (k, prefix) in enumerate(kinds):
            val = (phases.get(k) or {}).get("union_ns_median_per_pass")
            parts.append(_svg_num(x + pw / 2, py + 42 + j * 18, val, DERIVED, prefix))
        parts.append(f'<title>{_e(" / ".join(k for k, _ in kinds))}. {_e(F_UNION)}</title></g>')
        if i:
            parts.append(f'<line x1="{x - pgap + 2}" y1="{py + (bh + 16) / 2}" x2="{x - 3}" '
                         f'y2="{py + (bh + 16) / 2}" class="rt-line" marker-end="url(#arr)"/>')
    parts.append("</svg>")
    note = ('<p class="muted">Range and name timings sit in one box but are separate numbers: a range '
            'and a name covering the same cells are never added. Dashed boxes were not run in this '
            'profile.</p>')
    return _section("route", "Route", f'<div class="scroll">{"".join(parts)}</div>{note}',
                    "The end-to-end run as it happened, and what one profiling pass does inside Excel.")


def _tl_rows(rows: list, lo: float, hi: float) -> str:
    span = max(hi - lo, 1.0)
    out = []
    for r in rows:
        start = r.get("start_rel_ns") if _is_num(r.get("start_rel_ns")) else lo
        dur = r.get("dur_ns") if _is_num(r.get("dur_ns")) else 0
        left = (start - lo) / span * 100
        width = dur / span * 100
        depth = r.get("depth") if isinstance(r.get("depth"), int) else 0
        status = r.get("status")
        bad = status not in (None, "ok")
        kind = r.get("kind")
        tip = f"{KIND_LABELS.get(kind, kind)}: {r.get('name')} — {fmt_ns(dur)}" + (f" ({status})" if bad else "")
        out.append(
            f'<div class="tl-row{" tl-bad" if bad else ""}" title="{_e(tip)}">'
            f'<div class="tl-label" style="padding-left:{min(depth, 12) * 0.75:.2f}rem">'
            f'<span class="dot {KIND_CSS.get(kind, "k-other")}"></span>{_e(r.get("name"))}'
            + (f' <span class="status">{_e(status)}</span>' if bad else "") + '</div>'
            f'<div class="tl-track"><div class="tl-bar {KIND_CSS.get(kind, "k-other")}" '
            f'style="left:{left:.3f}%;width:{width:.3f}%"></div></div>'
            f'<div class="tl-dur">{_num(dur, MEASURED)}</div></div>')
    return "".join(out)


def _tl_block(rows: list, caption: str) -> str:
    if not rows:
        return '<p class="muted">No spans.</p>'
    lo = min(r.get("start_rel_ns") for r in rows if _is_num(r.get("start_rel_ns"))) if any(
        _is_num(r.get("start_rel_ns")) for r in rows) else 0
    hi = max((r.get("start_rel_ns") or 0) + (r.get("dur_ns") or 0) for r in rows)
    axis = (f'<div class="tl-axis"><span>{_e(caption)}</span><span>0</span>'
            f'<span>{_num(hi - lo, MEASURED)}</span></div>')
    return f'{axis}<div class="tl">{_tl_rows(rows, lo, hi)}</div>'


def _timeline(model: dict) -> str:
    tl = model.get("timeline") or {}
    rows = [r for r in (tl.get("rows") or []) if isinstance(r, dict)]
    host_rows = [r for r in rows if not r.get("pass")]
    by_pass: dict = {}
    for r in rows:
        p = r.get("pass")
        if p:
            by_pass.setdefault(p, []).append(r)
    body = []
    trunc = tl.get("truncated") or 0
    if trunc:
        body.append(f'<div class="banner warn" role="status"><strong>Timeline truncated.</strong> '
                    f'{_num(trunc, STATIC, "count")} further spans are not drawn (row cap '
                    f'{_num(tl.get("max_rows"), STATIC, "count")}). All analysis tables still use every span.</div>')
    kinds = sorted({r.get("kind") for r in rows if r.get("kind") in KIND_CSS},
                   key=lambda k: PHASE_ORDER.index(k) if k in PHASE_ORDER else -1)
    if kinds:
        body.append('<div class="kinds">' + "".join(
            f'<span><span class="dot {KIND_CSS[k]}"></span>{_e(KIND_LABELS.get(k, k))}</span>' for k in kinds)
            + "</div>")
    body.append("<h3>Host</h3>" + _tl_block(host_rows, "whole run"))
    passes = sorted(by_pass, key=lambda p: (not isinstance(p, int), p if isinstance(p, int) else str(p)))
    if passes:
        opts = "".join(f'<option value="{_e(p)}"{" selected" if i == 0 else ""}>Pass {_e(p)}</option>'
                       for i, p in enumerate(passes))
        body.append(f'<h3>Pass <label class="sel">detail for <select id="tl-pass-select" '
                    f'aria-label="Pass to show">{opts}</select></label></h3>')
        for i, p in enumerate(passes):
            body.append(f'<div class="tl-pass" data-pass="{_e(p)}"{"" if i == 0 else " hidden"}>'
                        + _tl_block(by_pass[p], f"pass {p}") + "</div>")
    else:
        body.append('<p class="muted">No pass spans in the timeline.</p>')
    return _section("timeline", "Timeline", "".join(body),
                    "Each bar is one span between its begin and end events, indented by nesting depth. "
                    "VBA spans are placed on the host timebase via the clock offset, so their position "
                    "relative to host spans is only as good as the offset uncertainty.")


def _summary_cells(summary: Any, cls: str = MEASURED) -> str:
    n = summary.get("n") if isinstance(summary, dict) else None
    return (f'<td data-v="{_sv(_med(summary))}">{_num(_med(summary), cls)}</td>'
            f'<td>{_range(summary, cls)}</td><td data-v="{_sv(n)}">{_num(n, STATIC, "count")}</td>')


def _semantic_inline(row: dict) -> str:
    annotations = [item for item in (row.get("annotations") or []) if isinstance(item, dict)]
    if not annotations:
        return ""
    parts = []
    for item in annotations[:3]:
        category = (f'<span class="semantic-category">{_e(item.get("category"))}</span>'
                    if item.get("category") else "")
        parts.append(f'<div class="semantic-inline"><strong>{_e(item.get("label"))}</strong>{category}'
                     f'<div>{_e(item.get("intent"))}</div></div>')
    if len(annotations) > 3:
        parts.append(f'<div class="muted">+{len(annotations) - 3} more annotations</div>')
    return "".join(parts)


def _drilldown(model: dict) -> str:
    dd = model.get("drilldown") or {}
    wb = dd.get("workbook") or {}
    full, recalc = wb.get("full"), wb.get("recalc")
    cards = []
    card_specs = [("Full calculation", full, "Application.CalculateFull — worst case"),
                  ("Recalculation", recalc, "Application.Calculate immediately after — best case")]
    if isinstance(wb.get("fullrebuild"), dict):
        card_specs.insert(0, ("Full rebuild", wb.get("fullrebuild"),
                              "Application.CalculateFullRebuild — rebuilds dependencies, then calculates"))
    for title, s, sub in card_specs:
        n = s.get("n") if isinstance(s, dict) else None
        cards.append(f'<div class="card"><h3>{_e(title)}</h3><p class="big">{_num(_med(s), MEASURED)}</p>'
                     f'<p class="muted">{_e(sub)}. median; min–max {_range(s)}; '
                     f'n {_num(n, STATIC, "count")}</p></div>')
    cards.append(f'<div class="card"><h3>Volatility ratio</h3><p class="big">'
                 f'{_num(wb.get("volatility_ratio"), DERIVED, "ratio", F_VOLATILITY)}</p>'
                 f'<p class="formula">{_e(F_VOLATILITY)}</p><p class="muted">Near 1: most of the workbook '
                 'recalculates every time (volatile functions or dirty cells). Near 0: recalcs are cheap.</p></div>')
    parts = [f'<div class="cards">{"".join(cards)}</div>']

    sheets = [s for s in (dd.get("sheets") or []) if isinstance(s, dict)]
    if sheets:
        parts.append('<h3>Sheets</h3><p class="muted">Worksheet.Calculate per sheet. The share bar is '
                     f'{_badge(DERIVED)} <span class="formula">{_e(F_SHARE)}</span>. Each sheet was timed as '
                     'its own call, so shares compare sheets with each other; they do not add up to the '
                     'workbook recalc time.</p>')
        parts.append('<div class="dd-head" aria-hidden="true"><span>Sheet</span><span>median recalc</span>'
                     '<span>min – max</span><span>n</span><span>share</span></div>')
        for s in sheets:
            summ = s.get("recalc")
            share = s.get("share_of_recalc_sum")
            pct = max(0.0, min(1.0, float(share))) * 100 if _is_num(share) else 0.0
            kids = [c for c in (s.get("children") or []) if isinstance(c, dict)]
            head = (f'<span class="dd-name">{_e(s.get("name"))}</span>'
                    f'<span>{_num(_med(summ), MEASURED)}</span><span>{_range(summ)}</span>'
                    f'<span>{_num(summ.get("n") if isinstance(summ, dict) else None, STATIC, "count")}</span>'
                    f'<span class="share"><span class="bar"><span style="width:{pct:.1f}%"></span></span>'
                    f'{_num(share, DERIVED, "pct", F_SHARE)}</span>')
            more = s.get("children_truncated")
            if not kids and not (_is_num(more) and more > 0):
                parts.append(f'<div class="dd-row leaf">{head}</div>')
                continue
            rows = []
            for c in kids:
                ck = c.get("kind")
                cs = c.get("summary")
                caveat = (f'<div class="caveat">{_badge(MEASURED, "isolated Range.Calculate")} '
                          f'{_e(GROUP_CAVEAT)}</div>' if ck == "calc.group" else "")
                cm = c.get("measurement")
                if ck != "calc.group" and isinstance(cm, str) and cm not in ("measured", ""):
                    caveat = f'<div class="caveat">{_e(cm)}</div>'

                rows.append(f'<tr><td>{_kind(ck)}</td><td>{_semantic_inline(c)}'
                            f'<span class="technical-name">{_e(c.get("name"))}</span>{caveat}</td>'
                            f'<td><code>{_e(c.get("address"))}</code></td>{_summary_cells(cs)}</tr>')
            more = s.get("children_truncated")
            if _is_num(more) and more > 0:
                rows.append(f'<tr class="more"><td colspan="6">{_num(more, STATIC, "count")} more not shown '
                            '(children capped to keep the report bounded)</td></tr>')
            has_group = any(c.get("kind") == "calc.group" for c in kids)
            gnote = ('<p class="muted">Group timings are separate isolated operations. They are not summed '
                     'and not divided by cell count; per-formula time is ' + _na("no per-formula timer")
                     + '.</p>') if has_group else ""
            parts.append(f'<details class="dd-row"><summary>{head}</summary><div class="dd-kids">'
                         '<table><thead><tr><th>kind</th><th>name</th><th>address</th><th>median</th>'
                         f'<th>min – max</th><th>n</th></tr></thead><tbody>{"".join(rows)}</tbody></table>'
                         f'{gnote}</div></details>')
        if _is_num(dd.get("sheets_truncated")) and dd.get("sheets_truncated") > 0:
            parts.append(f'<p class="muted">{_num(dd.get("sheets_truncated"), STATIC, "count")} more sheets '
                         'not shown.</p>')
    else:
        parts.append('<p class="muted">No sheet timings in this trace.</p>')
    return _section("drilldown", "Drill-down", "".join(parts),
                    "Microsoft's order: whole workbook, then each sheet, then ranges, defined names and "
                    "(opt-in) formula groups inside it. Expand a sheet to see its children.")


def _hotspots(model: dict) -> str:
    hs = [h for h in (model.get("hotspots") or []) if isinstance(h, dict)]
    if not hs:
        return _section("hotspots", "Hotspots", '<p class="muted">No hotspots.</p>')
    rows = []
    for i, h in enumerate(hs, 1):
        # Total, min and max are direct timings of repeats of one operation. Self time
        # equals the total for a leaf span (measured) and is derived otherwise.
        mcls = MEASURED
        self_cell = (_num(h.get("median_self_ns"), MEASURED) if h.get("measurement") == "measured"
                     else _num(h.get("median_self_ns"), DERIVED, "ns", F_SELF))
        note = h.get("note") if isinstance(h.get("note"), str) else None
        if h.get("kind") == "calc.group" and not note:
            note = "measured (isolated Range.Calculate)"
        caveat = f'<div class="caveat">{_e(note)}</div>' if note else ""
        rows.append(
            f'<tr><td data-v="{i}">{i}</td><td data-v="{_sv(h.get("kind"))}">{_kind(h.get("kind"))}</td>'
            f'<td data-v="{_sv(h.get("name"))}">{_semantic_inline(h)}'
            f'<span class="technical-name">{_e(h.get("name"))}</span>'
            f'{("<code class=technical-address>" + _e(h.get("address")) + "</code>") if h.get("address") else ""}{caveat}</td>'
            f'<td data-v="{_sv(h.get("sheet"))}">{_e(h.get("sheet") or "")}</td>'
            f'<td data-v="{_sv(h.get("median_self_ns"))}">{self_cell}</td>'
            f'<td data-v="{_sv(h.get("median_ns"))}">{_num(h.get("median_ns"), mcls)}</td>'
            f'<td data-v="{_sv(h.get("n"))}">{_num(h.get("n"), STATIC, "count")}</td>'
            f'<td data-v="{_sv(h.get("min_ns"))}">{_num(h.get("min_ns"), mcls)} – {_num(h.get("max_ns"), mcls)}</td></tr>')
    table = ('<div class="scroll"><table class="sortable"><thead><tr>'
             '<th data-sort-type="num">#</th><th data-sort-type="str">kind</th><th data-sort-type="str">name</th>'
             '<th data-sort-type="str">sheet</th>'
             '<th data-sort-type="num" aria-sort="descending">median self</th>'
             f'<th data-sort-type="num">median total {_badge(MEASURED)}</th>'
             '<th data-sort-type="num">n</th><th data-sort-type="num">min – max</th></tr></thead>'
             f'<tbody>{"".join(rows)}</tbody></table></div>')
    note = (f'<p class="formula">{_e(F_SELF)}</p><p class="muted">Median self time is measured for spans '
            'with no child spans (self = total) and derived otherwise. Ranked by median self time. Click a '
            'column heading to sort. A range and a name covering the same cells are listed separately '
            'and never added together.</p>')
    return _section("hotspots", "Hotspots", table + note)


def _proc_list(nodes: list) -> str:
    if not nodes:
        return ""
    items = []
    for n in nodes:
        s = n.get("summary")
        ss = n.get("self_summary")
        items.append(f'<li><div class="proc">{_kind(n.get("kind"))} {_semantic_inline(n)}'
                     f'<strong class="technical-name">{_e(n.get("name"))}</strong> '
                     f'median {_num(_med(s), MEASURED)}, self {_num(_med(ss), DERIVED, "ns", F_SELF)}, '
                     f'min–max {_range(s)}, n {_num(s.get("n") if isinstance(s, dict) else None, STATIC, "count")}'
                     f'</div>{_proc_list(n.get("children") or [])}</li>')
    return f'<ul class="tree">{"".join(items)}</ul>'


def _phases(model: dict) -> str:
    ph = [p for p in (model.get("phases") or []) if isinstance(p, dict)]
    ph.sort(key=lambda p: PHASE_ORDER.index(p.get("kind")) if p.get("kind") in PHASE_ORDER else 99)
    parts = ['<p class="note"><strong>These are unions, not sums.</strong> For each kind, the time is the '
             'length of the union of all its spans within a pass, so overlapping or nested spans are '
             'not double counted. Kinds nest (every calc runs inside its pass, and a macro may trigger calcs), '
             'so the column does not add up to the pass time.</p>']
    if ph:
        rows = "".join(
            f'<tr><td>{_kind(p.get("kind"))}</td><td>{_num(p.get("count"), STATIC, "count")}</td>'
            f'<td>{_num(p.get("passes"), STATIC, "count")}</td>'
            f'<td>{_num(p.get("union_ns_median_per_pass"), DERIVED, "ns", F_UNION)}</td>'
            f'<td>{_num(p.get("share_of_pass"), DERIVED, "pct", F_SHARE_PASS)}</td></tr>' for p in ph)
        parts.append('<div class="scroll"><table><thead><tr><th>kind</th><th>spans</th><th>in passes</th>'
                     f'<th>union per pass (median) {_badge(DERIVED)}</th><th>share of pass {_badge(DERIVED)}</th>'
                     f'</tr></thead><tbody>{rows}</tbody></table></div>'
                     f'<p class="formula">{_e(F_UNION)}; {_e(F_SHARE_PASS)}</p>')
    else:
        parts.append('<p class="muted">No phase data.</p>')
    passes = [p for p in (model.get("passes") or []) if isinstance(p, dict)]
    if passes:
        rows = "".join(f'<tr><td>{_e(p.get("pass"))}</td><td>{_num(p.get("dur_ns"), MEASURED)}</td>'
                       f'<td>{_e(p.get("status") or "")}</td></tr>' for p in passes)
        parts.append('<details><summary>Per-pass durations</summary><table><thead><tr><th>pass</th>'
                     f'<th>run.pass {_badge(MEASURED)}</th><th>status</th></tr></thead><tbody>{rows}</tbody>'
                     '</table></details>')
    procs = model.get("vba_procs") or []
    parts.append("<h3>VBA procedures</h3>")
    if procs:
        parts.append('<p class="muted">Procedures from <code>--macro</code> and their '
                     '<code>XSP_Begin</code>/<code>XSP_End</code> spans, merged across passes by name path.</p>'
                     + _proc_list(procs))
    else:
        parts.append('<p class="muted">No VBA procedures were profiled.</p>')
    return _section("phases", "VBA and calculation phases", "".join(parts))


def _overhead(model: dict) -> str:
    oh = model.get("overhead") or {}
    if not oh.get("available"):
        reason = oh.get("reason") or "Instrumentation overhead unavailable."
        return _section("overhead", "Instrumentation overhead",
                        f'<p>{_na("no trace-off run")} {_e(reason)}</p>')
    rows = []
    for r in (oh.get("rows") or []):
        if not isinstance(r, dict):
            continue
        head = f'<td>{_kind(r.get("kind"))}</td><td>{_e(r.get("name"))}</td>'
        if r.get("available") is False:
            reason = r.get("unavailable_reason") if isinstance(r.get("unavailable_reason"), str) else "no reason given"
            rows.append(f'<tr>{head}<td colspan="7" class="verdict">{_na("overhead unavailable")} unavailable: {_e(reason)}</td></tr>')
            continue
        if r.get("resolvable") is True:
            verdict = '<span class="flag on">resolvable</span>'
        else:
            verdict = '<span class="flag">not resolvable (below run-to-run noise)</span>'
            if _is_num(r.get("pairs")) and r.get("pairs") < 3:
                verdict += ' <span class="muted">(fewer than 3 pairs)</span>'
        for key in ("reason", "unavailable_reason"):
            if isinstance(r.get(key), str):
                verdict += f' <span class="muted">{_e(r.get(key))}</span>'
        dropped = _num(r.get("dropped_pairs"), STATIC, "count")
        dp = [str(x) for x in (r.get("dropped_passes") or []) if _is_num(x)] if isinstance(
            r.get("dropped_passes"), list) else []
        if dp:
            shown = ", ".join(dp[:20]) + (f" +{len(dp) - 20} more" if len(dp) > 20 else "")
            dropped += f' <span class="muted">(pass {_e(shown)})</span>'
        rows.append(f'<tr>{head}'
                    f'<td>{_num(r.get("median_on_ns"), MEASURED)}</td><td>{_num(r.get("median_off_ns"), MEASURED)}</td>'
                    f'<td>{_num(r.get("diff_ns"), DERIVED, "ns", F_OVERHEAD)}</td>'
                    f'<td>{_num(r.get("diff_min_ns"), DERIVED, "ns", F_OVERHEAD_RANGE)} – '
                    f'{_num(r.get("diff_max_ns"), DERIVED, "ns", F_OVERHEAD_RANGE)}</td>'
                    f'<td>{_num(r.get("pairs"), STATIC, "count")}</td><td>{dropped}</td>'
                    f'<td class="verdict">{verdict}<div class="muted">spread (IQR) on {_num(r.get("iqr_on_ns"), DERIVED, "ns", F_IQR)}'
                    f' / off {_num(r.get("iqr_off_ns"), DERIVED, "ns", F_IQR)}</div></td></tr>')
    oh_rows = [r for r in (oh.get("rows") or []) if isinstance(r, dict)]
    lead = ""
    if oh_rows and all(r.get("available") is False for r in oh_rows):
        reason = oh_rows[0].get("unavailable_reason")
        lead = (f'<p>{_na("overhead unavailable")} Instrumentation overhead is unavailable: '
                f'{_e(reason if isinstance(reason, str) else "no reason given")}</p>')
    table = (lead + '<div class="scroll"><table class="overhead"><thead><tr><th>kind</th><th>name</th>'
             f'<th>median on {_badge(MEASURED)}</th><th>median off {_badge(MEASURED)}</th>'
             f'<th>overhead {_badge(DERIVED)}</th><th>paired diff min – max {_badge(DERIVED)}</th>'
             f'<th>pairs</th><th>dropped pairs</th><th class="verdict">verdict</th></tr></thead><tbody>{"".join(rows)}</tbody>'
             '</table></div>')
    note = (f'<p class="formula">{_e(F_OVERHEAD)}; {_e(F_OVERHEAD_RANGE)}</p><p class="muted">Trace-off passes '
            'run the same calc calls in the same order but record only pass, full and recalc spans. Runs '
            'are interleaved off, on, off, on. Trace-off pass k is paired with trace-on pass k, and only '
            'when both traces carry the same run_id; passes that failed or are missing from either trace are '
            'dropped from pairing and listed. The overhead is the median of the paired differences. It is '
            'resolvable only with at least 3 pairs and when the min–max of the paired differences excludes 0; '
            'otherwise it is below run-to-run noise. Median on and median off are the medians of each run\'s '
            'paired samples, so their difference need not equal the median paired difference. The IQR spread under each verdict is for information only '
            'and is not part of the rule.</p>')
    return _section("overhead", "Instrumentation overhead", table + note)


def _formulas(model: dict) -> str:
    fm = model.get("formulas")
    if not fm:
        return _section("formulas", "Formula diagnostics",
                        f'<p>{_na("no formulas.json")} No formula inspection was supplied.</p>')
    t = fm.get("totals") or {}
    tiles = "".join(
        f'<div class="tile"><span class="tl-k">{_e(label)}</span>{_num(t.get(key), STATIC, "count")}</div>'
        for key, label in FORMULA_TOTALS.items() if key in t)
    parts = [f'<div class="tiles">{tiles}</div>',
             '<p class="muted">Read from the file, not timed. Formulas are grouped by a fingerprint '
             '(hash of the normalised relative R1C1 form); formula text is never stored or shown. '
             'Whether UDFs call <code>Application.Volatile</code> is '
             + _na("VBA source not parsed in v1") + '.</p>']
    sheets = fm.get("sheets") or []
    if sheets:
        rows = "".join(f'<tr><td>{_e(s.get("sheet"))}</td><td>{_num(s.get("formula_cells"), STATIC, "count")}</td>'
                       f'<td><code>{_e(s.get("used_range"))}</code></td>'
                       f'<td>{_num(s.get("data_tables"), STATIC, "count")}</td></tr>' for s in sheets)
        parts.append('<details><summary>Sheets</summary><div class="scroll"><table><thead><tr><th>sheet</th>'
                     '<th>formula cells</th><th>used range</th><th>data tables</th></tr></thead>'
                     f'<tbody>{rows}</tbody></table></div></details>')
    groups = fm.get("groups") or []
    if groups:
        rows = []
        for g in groups:
            areas = g.get("areas") or []
            n_areas = g.get("areas_total") if _is_num(g.get("areas_total")) else len(areas)
            area_txt = ", ".join(areas[:3]) + (f" +{int(n_areas) - min(len(areas), 3)} more"
                                               if n_areas > min(len(areas), 3) else "")
            other = [lbl for key, lbl in (("external_ref", "external"), ("whole_column_ref", "whole column"),
                                          ("whole_row_ref", "whole row")) if g.get(key)]
            if _is_num(g.get("lambda_calls")) and g.get("lambda_calls"):
                other.append(f"LAMBDA calls {_num(g.get('lambda_calls'), STATIC, 'count')}")
            vol = _yes(g.get("volatile")) + (
                f' <span class="muted">{_e(", ".join(g.get("volatile_functions") or []))}</span>'
                if g.get("volatile_functions") else "")
            st = _yes(g.get("single_threaded")) + (
                f' <span class="muted">{_e(", ".join(g.get("single_threaded_functions") or []))}</span>'
                if g.get("single_threaded_functions") else "")
            arr = "dynamic" if g.get("dynamic_array") else ("yes" if g.get("array") else "no")
            rows.append(
                f'<tr><td data-v="{_sv(g.get("group"))}"><code>{_e(g.get("group"))}</code></td>'
                f'<td>{_semantic_inline(g) or "<span class=muted>Unmapped</span>"}</td>'
                f'<td data-v="{_sv(g.get("fingerprint"))}"><code class="fp">{_e(g.get("fingerprint"))}</code></td>'
                f'<td data-v="{_sv(g.get("sheet"))}">{_e(g.get("sheet"))}</td>'
                f'<td data-v="{_sv(g.get("cells"))}">{_num(g.get("cells"), STATIC, "count")}</td>'
                f'<td data-v="{_sv(n_areas)}"><code>{_e(area_txt)}</code></td>'
                f'<td data-v="{_sv(" ".join(g.get("functions") or []))}">{_e(", ".join(g.get("functions") or []))}</td>'
                f'<td data-v="{_sv(g.get("volatile"))}">{vol}</td>'
                f'<td data-v="{_sv(g.get("udfs"))}">{_num(g.get("udfs"), STATIC, "count")}</td>'
                f'<td data-v="{_sv(g.get("single_threaded"))}">{st}</td>'
                f'<td data-v="{_sv(arr)}">{_e(arr)}</td>'
                f'<td data-v="{_sv(g.get("cross_sheet"))}">{_yes(g.get("cross_sheet"))}</td>'
                f'<td data-v="{_sv(len(other))}">{", ".join(o if o.startswith("LAMBDA") else _e(o) for o in other)}</td>'
                f'<td data-v="{_sv(g.get("length_bucket"))}">{_e(g.get("length_bucket"))}</td></tr>')
        shown, total = len(groups), fm.get("groups_total") or len(groups)
        cap = (f'<p class="muted">Showing the {_num(shown, STATIC, "count")} largest of '
               f'{_num(total, STATIC, "count")} groups.</p>' if _is_num(total) and total > shown else "")
        parts.append('<h3>Formula groups</h3><div class="scroll"><table class="sortable"><thead><tr>'
                     '<th data-sort-type="str">group</th><th>meaning / intent</th><th data-sort-type="str">fingerprint</th>'
                     '<th data-sort-type="str">sheet</th><th data-sort-type="num" aria-sort="descending">cells</th>'
                     '<th data-sort-type="num">areas</th><th data-sort-type="str">functions</th>'
                     '<th data-sort-type="num">volatile</th><th data-sort-type="num">UDFs</th>'
                     '<th data-sort-type="num">single-threaded</th><th data-sort-type="str">array</th>'
                     '<th data-sort-type="num">cross-sheet</th><th data-sort-type="num">other refs</th>'
                     '<th data-sort-type="str">length</th></tr></thead>'
                     f'<tbody>{"".join(rows)}</tbody></table></div>{cap}')
    names = fm.get("names") or []
    if names:
        rows = "".join(
            f'<tr><td>{_semantic_inline(n) or "<span class=muted>Unmapped</span>"}</td>'
            f'<td><code>{_e(n.get("name"))}</code></td><td>{_e(n.get("scope"))}</td>'
            f'<td><code>{_e(n.get("refers_to_range"))}</code></td>'
            f'<td>{_num(n.get("formula_cells"), STATIC, "count")}</td></tr>' for n in names[:500])
        parts.append('<details><summary>Defined names</summary><div class="scroll"><table><thead><tr>'
                     '<th>meaning / intent</th><th>name</th><th>scope</th><th>refers to</th>'
                     '<th>formula cells</th></tr></thead><tbody>' + rows + '</tbody></table></div></details>')
    if fm.get("redaction_names"):
        parts.append(f'<p class="muted">Identifier redaction in formula inspection: {_e(fm.get("redaction_names"))}.</p>')
    if _is_num(fm.get("names_count")):
        parts.append(f'<p class="muted">Defined names inspected: {_num(fm.get("names_count"), STATIC, "count")}.</p>')
    return _section("formulas", "Formula diagnostics", "".join(parts))


def _semantics(model: dict) -> str:
    semantic_map = ((model.get("formulas") or {}).get("semantics") or {})
    annotations = semantic_map.get("annotations") or []
    if not annotations:
        return _section("semantics", "Model meaning",
                        '<p class="muted">No semantic map was supplied. Technical names and cell references remain available in the report.</p>')
    rows = []
    for item in annotations:
        count = item.get("timed_match_count") or 0
        if item.get("binding") == "unmatched":
            observed = '<span class="flag">Target not found in this workbook</span>'
        elif count:
            observed = f'<span class="flag on">Matched {count} report rows</span>'
        else:
            observed = '<span class="flag">Mapped; no matching timed row in this run</span>'
        category = (f'<div class="semantic-category">{_e(item.get("category"))}</div>'
                    if item.get("category") else "")
        rows.append(f'<tr><td><strong>{_e(item.get("label"))}</strong>{category}</td>'
                    f'<td>{_e(item.get("intent"))}</td><td>{_e(item.get("source") or "Analyst annotation")}</td>'
                    f'<td>{observed}</td></tr>')
    warnings = []
    if semantic_map.get("unbound_count"):
        warnings.append(f'{_num(semantic_map.get("unbound_count"), STATIC, "count")} annotation(s) did not resolve to workbook structure.')
    if not semantic_map.get("workbook_bound"):
        warnings.append('This semantic map is not bound to a workbook SHA-256.')
    note = '<p class="muted">Descriptions come from the supplied map. They identify the purpose of measured regions; they do not assign per-formula runtime.</p>'
    if warnings:
        note += '<ul class="muted">' + ''.join(f'<li>{item}</li>' for item in warnings) + '</ul>'
    table = ('<div class="scroll"><table><thead><tr><th>semantic label</th><th>intent</th>'
             '<th>source</th><th>trace match</th></tr></thead><tbody>' + ''.join(rows) + '</tbody></table></div>')
    return _section("semantics", "Model meaning", note + table,
                    "Human-authored descriptions attached to workbook structure and timed spans.")


def _limits(model: dict) -> str:
    items = [
        ("No per-formula timer", "Excel has no per-formula execution timer. Time inside a normal recalc "
         "cannot be attributed to individual formulas; per-formula runtime is " + _na("per-formula runtime") + "."),
        ("Range.Calculate isolation", "Range, name and group timings run Range.Calculate on those cells with "
         "outside dependencies held at their current values. They exclude calc-chain scheduling and "
         "cross-range effects, and may thread differently from a normal recalc. Iteration is switched off "
         "during range timings."),
        ("Screen updating excluded", "ScreenUpdating is off while profiling, so Excel's final calc phase "
         "(updating visible windows) is not in any timing."),
        ("Multithreading", "Multithreaded calculation spreads work across threads. Wall time is what is "
         "measured; which thread ran what (thread placement) is " + _na("thread placement") + "."),
        ("First versus repeat calculation", "The first calculation after open can be slower (calc chain "
         "build, caches). A warmup runs first, and CalculateFull is timed on every pass. The cost of "
         "calc-chain reordering is " + _na("calc chain reorder cost") + "."),
        ("Operating system noise", "Other processes and power management add run-to-run noise. Microsoft's "
         "guidance is to repeat and average; this report repeats and uses medians with min–max."),
        ("Clock alignment", "VBA timings (MicroTimer) and host timings (perf_counter) are aligned by one "
         "clock probe. Positions across the two clocks are uncertain by the offset uncertainty in the header."),
        ("Not measurable here", "Per-formula runtime inside a normal recalc, thread placement, calc-chain "
         "reorder cost, whether UDFs call Application.Volatile, and database functions over PivotTables "
         "are all " + _na("not measurable") + "."),
    ]
    body = "".join(f"<li><strong>{_e(t)}.</strong> {d}</li>" for t, d in items)
    fm_limits = (model.get("formulas") or {}).get("limitations") or []
    if fm_limits:
        body += ('<li><strong>Formula inspection limits.</strong> Reported by the static inspector:<ul>'
                 + "".join(f"<li>{_e(x)}</li>" for x in fm_limits if isinstance(x, str)) + "</ul></li>")
    return _section("limits", "Measurement limits", f'<ul class="limits">{body}</ul>',
                    "Method: Microsoft's MicroTimer and workbook → worksheet → range drill-down, from the "
                    "article “Excel performance: Improving calculation performance”.")


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------

_CSS = r"""
:root{--bg:#fbfbfa;--fg:#1d1f23;--muted:#5d636d;--panel:#ffffff;--panel-2:#f1f2f4;--line:#d6d9de;
--accent:#1f5fbf;--measured:#1f6f43;--derived:#1f5fbf;--static:#6b4fa0;--na:#7a7f87;
--synth-bg:#fff1c2;--synth-fg:#5b4200;--bad-bg:#fde2e1;--bad-fg:#8a1c17;--warn-bg:#fff4e0;--warn-fg:#6b4500;
--k-host:#7d8793;--k-pass:#b9c0c9;--k-full:#c0392b;--k-recalc:#e08e2b;--k-sheet:#2e86c1;--k-range:#17a589;
--k-name:#8e6fc4;--k-group:#d4a017;--k-proc:#c2458c;--k-other:#95a0ab;color-scheme:light dark}
@media (prefers-color-scheme:dark){:root{--bg:#15171a;--fg:#e6e8eb;--muted:#9aa1ab;--panel:#1d2024;--panel-2:#262a2f;
--line:#353a41;--accent:#7fb0ff;--measured:#6fcf97;--derived:#7fb0ff;--static:#c3a6f5;--na:#9aa1ab;
--synth-bg:#4a3a00;--synth-fg:#ffe28a;--bad-bg:#4d1512;--bad-fg:#ffb4ae;--warn-bg:#3d2c0c;--warn-fg:#ffd48a;
--k-host:#8d97a3;--k-pass:#4a525c;--k-full:#e6655a;--k-recalc:#f0a852;--k-sheet:#5aa9e0;--k-range:#45c9ad;
--k-name:#ad93e0;--k-group:#e8c14a;--k-proc:#e174b0;--k-other:#7c8792}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
main{max-width:1180px;margin:0 auto;padding:24px 16px 64px}
h1{font-size:1.7rem;margin:.1rem 0 .8rem;overflow-wrap:anywhere}
h2{font-size:1.25rem;margin:0 0 .6rem}
h3{font-size:1rem;margin:1.2rem 0 .5rem}
section,header#run{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:18px 18px;margin:0 0 18px}
.eyebrow{margin:0;color:var(--muted);font-size:.8rem;text-transform:uppercase;letter-spacing:.06em}
.intro{color:var(--muted);margin-top:-.3rem}
.muted{color:var(--muted);font-size:.9rem}
code{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:.88em;overflow-wrap:anywhere}
.fp{font-size:.8em}
.banner{border-radius:8px;padding:12px 14px;margin:.6rem 0;font-size:.95rem}
.banner ul{margin:.4rem 0 0 1.1rem;padding:0}
.banner.synth{background:var(--synth-bg);color:var(--synth-fg);border:2px solid currentColor;font-size:1.05rem}
.banner.invalid{background:var(--bad-bg);color:var(--bad-fg);border:2px solid currentColor}
.banner.warn{background:var(--warn-bg);color:var(--warn-fg)}
.facts{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:8px 20px;margin:.8rem 0 0}
.facts div{min-width:0}.facts dt{color:var(--muted);font-size:.78rem;text-transform:uppercase;letter-spacing:.04em}
.facts dd{margin:0;overflow-wrap:anywhere}
.num{font-variant-numeric:tabular-nums;white-space:nowrap}
.num::after{content:"m";display:inline-block;margin-left:3px;padding:0 3px;border-radius:3px;font-size:.66em;
line-height:1.35;vertical-align:.15em;font-weight:600;color:var(--bg);background:var(--measured)}
.num[data-class=derived]::after{content:"d";background:var(--derived)}
.num[data-class=static]::after{content:"s";background:var(--static)}
.num[data-class=na]::after{content:"n/a";background:var(--na)}
.badge{display:inline-block;padding:0 6px;border-radius:9px;font-size:.72rem;font-weight:600;border:1px solid currentColor;
white-space:nowrap;vertical-align:.1em}
.badge[data-class=measured]{color:var(--measured)}.badge[data-class=derived]{color:var(--derived)}
.badge[data-class=static]{color:var(--static)}.badge[data-class=na]{color:var(--na)}
.legend{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:10px 18px}
.legend .leg p{margin:.3rem 0 0;font-size:.9rem}
.legend>p{grid-column:1/-1;margin:.2rem 0 0}
.formula{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:.82rem;color:var(--derived)}
.scroll{overflow-x:auto;-webkit-overflow-scrolling:touch}
table{border-collapse:collapse;width:100%;font-size:.9rem}
th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);vertical-align:top}
th{font-weight:600;color:var(--muted);font-size:.8rem;white-space:nowrap}
table.sortable th[data-sort-type]{cursor:pointer;user-select:none}
table.sortable th[aria-sort=descending]::after{content:" ▾"}table.sortable th[aria-sort=ascending]::after{content:" ▴"}
.flag{color:var(--muted)}.flag.on{color:var(--fg);font-weight:600}
.kind{font-size:.8rem;white-space:nowrap}
.kind::before,.dot{content:"";display:inline-block;width:.6em;height:.6em;border-radius:2px;margin-right:.35em;background:var(--k-other)}
.dot{margin-right:.4em}
.k-host.kind::before,.dot.k-host,.tl-bar.k-host{background:var(--k-host)}
.k-pass.kind::before,.dot.k-pass,.tl-bar.k-pass{background:var(--k-pass)}
.k-full.kind::before,.dot.k-full,.tl-bar.k-full{background:var(--k-full)}
.k-recalc.kind::before,.dot.k-recalc,.tl-bar.k-recalc{background:var(--k-recalc)}
.k-sheet.kind::before,.dot.k-sheet,.tl-bar.k-sheet{background:var(--k-sheet)}
.k-range.kind::before,.dot.k-range,.tl-bar.k-range{background:var(--k-range)}
.k-name.kind::before,.dot.k-name,.tl-bar.k-name{background:var(--k-name)}
.k-group.kind::before,.dot.k-group,.tl-bar.k-group{background:var(--k-group)}
.k-proc.kind::before,.dot.k-proc,.tl-bar.k-proc{background:var(--k-proc)}
svg.route{display:block;max-width:none;font-size:12px}
.rt-box{fill:var(--panel-2);stroke:var(--line);stroke-width:1.2}
.rt-box.rt-on{stroke:var(--accent);stroke-width:2}
.rt-box.rt-bad{stroke:var(--bad-fg);stroke-width:2}
.rt-box.rt-missing{fill:none;stroke-dasharray:4 3}
.rt-t{fill:var(--fg);font-weight:600}.rt-n{fill:var(--fg)}.rt-b{fill:var(--muted);font-size:9px}
.rt-cap{fill:var(--muted);font-size:11px}
.rt-line{stroke:var(--muted);stroke-width:1.3}.rt-arrow{fill:var(--muted)}
.rt-zoom{stroke:var(--accent);stroke-dasharray:3 3;fill:none;opacity:.7}
.kinds{display:flex;flex-wrap:wrap;gap:4px 14px;font-size:.82rem;color:var(--muted);margin-bottom:.4rem}
.tl-axis{display:grid;grid-template-columns:minmax(7rem,15rem) 1fr auto;gap:8px;font-size:.78rem;color:var(--muted);
border-bottom:1px solid var(--line);padding-bottom:2px}
.tl-axis span:nth-child(3){text-align:right}
.tl{max-height:560px;overflow-y:auto}
.tl-row{display:grid;grid-template-columns:minmax(7rem,15rem) 1fr 6.5rem;gap:8px;align-items:center;font-size:.8rem;
min-height:20px;border-bottom:1px solid var(--panel-2)}
.tl-row:hover{background:var(--panel-2)}
.tl-label{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tl-track{position:relative;height:12px}
.tl-bar{position:absolute;top:0;height:12px;min-width:2px;border-radius:2px;background:var(--k-other)}
.tl-bad .tl-bar{outline:2px solid var(--bad-fg)}.status{color:var(--bad-fg);font-weight:600}
.tl-dur{text-align:right}
label.sel{font-weight:400;font-size:.9rem;color:var(--muted)}
select{font:inherit;background:var(--panel);color:var(--fg);border:1px solid var(--line);border-radius:6px;padding:2px 6px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px}
.card{background:var(--panel-2);border-radius:8px;padding:12px 14px}.card h3{margin:0}
.card p{margin:.3rem 0}.big{font-size:1.5rem;font-weight:600}
.dd-head,.dd-row>summary,.dd-row.leaf{display:grid;grid-template-columns:minmax(8rem,2fr) 1fr 1.6fr .5fr 1.6fr;gap:8px;
align-items:center;padding:6px 8px}
.dd-head{font-size:.78rem;color:var(--muted);font-weight:600;border-bottom:1px solid var(--line)}
.dd-row{border-bottom:1px solid var(--line)}
.dd-row>summary{cursor:pointer;list-style:none}.dd-row>summary::-webkit-details-marker{display:none}
.dd-row>summary .dd-name::before{content:"▸ ";color:var(--muted)}.dd-row[open]>summary .dd-name::before{content:"▾ "}
.dd-row.leaf .dd-name{padding-left:1em}
.dd-name{font-weight:600;overflow-wrap:anywhere}
.share{display:flex;align-items:center;gap:6px}
.bar{flex:1;height:8px;background:var(--panel-2);border-radius:4px;overflow:hidden;min-width:40px}
.bar>span{display:block;height:100%;background:var(--derived);opacity:.75}
.dd-kids{padding:4px 8px 12px 1.6rem;overflow-x:auto}
.caveat{font-size:.8rem;color:var(--muted);margin-top:2px}.caveat-inline{font-size:.8rem;color:var(--muted)}
.semantic-inline{margin:2px 0 5px;padding:5px 7px;border-left:3px solid var(--accent);background:var(--panel-2);border-radius:3px;overflow-wrap:anywhere}
.semantic-inline>div{font-size:.82rem;color:var(--muted);font-weight:400;margin-top:2px}
.semantic-category{display:inline-block;margin-left:7px;padding:1px 6px;border-radius:9px;background:var(--panel-2);color:var(--muted);font-size:.72rem;font-weight:600}
.technical-name{display:block;font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:.78rem;color:var(--muted);overflow-wrap:anywhere}
.technical-address{display:block;margin-top:2px;color:var(--muted);font-size:.76rem;overflow-wrap:anywhere}
.note{background:var(--panel-2);border-left:3px solid var(--derived);padding:8px 12px;border-radius:4px}
.tree{list-style:none;padding-left:1.1rem;margin:.2rem 0}.tree .tree{border-left:1px solid var(--line)}
.proc{padding:3px 0;font-size:.9rem}
.tiles{display:flex;flex-wrap:wrap;gap:10px}.tile{background:var(--panel-2);border-radius:8px;padding:8px 12px;min-width:130px}
.tile .num{font-size:1.2rem;font-weight:600;display:block}.tl-k{font-size:.78rem;color:var(--muted)}
.limits li{margin:.35rem 0}
table.overhead th{white-space:normal}
table.overhead .verdict{min-width:16rem}
table.overhead .verdict .num{white-space:nowrap}
footer{color:var(--muted);font-size:.82rem;text-align:center}
@media (max-width:640px){main{padding:12px 16px 40px}section,header#run{padding:14px 12px}
.tl-row,.tl-axis{grid-template-columns:6.5rem 1fr 5.5rem}
.dd-head{display:none}.dd-row>summary,.dd-row.leaf{grid-template-columns:1fr 1fr}
.dd-row .dd-name,.dd-row .share{grid-column:1/-1}}
"""

_JS = r"""
(function(){
  var sel=document.getElementById('tl-pass-select');
  if(sel){sel.addEventListener('change',function(){
    var blocks=document.querySelectorAll('.tl-pass');
    for(var i=0;i<blocks.length;i++){blocks[i].hidden=blocks[i].getAttribute('data-pass')!==sel.value;}
  });}
  var tables=document.querySelectorAll('table.sortable');
  for(var t=0;t<tables.length;t++){(function(table){
    var ths=table.tHead.rows[0].cells;
    for(var c=0;c<ths.length;c++){(function(th,col){
      var type=th.getAttribute('data-sort-type'); if(!type) return;
      th.tabIndex=0;
      function sortBy(){
        var dir=th.getAttribute('aria-sort')==='descending'?'ascending':'descending';
        for(var k=0;k<ths.length;k++){ths[k].removeAttribute('aria-sort');}
        th.setAttribute('aria-sort',dir);
        var body=table.tBodies[0], rows=Array.prototype.slice.call(body.rows);
        rows.sort(function(a,b){
          var x=a.cells[col].getAttribute('data-v')||'', y=b.cells[col].getAttribute('data-v')||'', r;
          if(type==='num'){x=parseFloat(x);y=parseFloat(y);if(isNaN(x))x=-Infinity;if(isNaN(y))y=-Infinity;r=x<y?-1:x>y?1:0;}
          else{r=x<y?-1:x>y?1:0;}
          return dir==='ascending'?r:-r;
        });
        for(var i=0;i<rows.length;i++){body.appendChild(rows[i]);}
      }
      th.addEventListener('click',sortBy);
      th.addEventListener('keydown',function(e){if(e.key==='Enter'||e.key===' '){e.preventDefault();sortBy();}});
    })(ths[c],c);}
  })(tables[t]);}
})();
"""


def render_html(model: dict) -> str:
    """Render the view-model as one standalone HTML page. Pure and deterministic."""
    model = _apply_caps(model if isinstance(model, dict) else {})
    meta = model.get("meta") or {}
    val = model.get("validation") or {}
    show_timing = bool(val.get("ok")) or bool(val.get("allow_invalid"))
    title = f"XLSprint · {meta.get('label') or 'profile'}"
    sections = [_header(model), _legend()]
    if show_timing:
        sections += [_route(model), _timeline(model), _drilldown(model), _hotspots(model),
                     _phases(model), _overhead(model)]
        errs = model.get("section_errors") or {}
        if errs:
            items = "".join(f"<li><code>{_e(k)}</code>: {_e(v)}</li>" for k, v in sorted(errs.items()))
            sections.insert(2, f'<div class="banner invalid" role="alert"><strong>Some sections could not '
                               f'be analysed from this invalid trace.</strong><ul>{items}</ul></div>')
    else:
        sections.append('<section id="timing-withheld"><h2>Timing withheld</h2><p>The trace failed '
                        'validation, so no timing sections are shown. Fix the trace, or re-render with '
                        '<code>--allow-invalid</code> to inspect it anyway.</p></section>')
    sections += [_semantics(model), _formulas(model), _limits(model)]
    return (
        "<!DOCTYPE html>\n<html lang=\"en\"><head><meta charset=\"utf-8\">"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">"
        "<meta name=\"color-scheme\" content=\"light dark\">"
        f"<title>{_e(title)}</title><style>{_CSS}</style></head><body><main>"
        + "".join(sections)
        + f"<footer>Generated by XLSprint {_e(__version__)}. Report schema {REPORT_SCHEMA}.</footer></main>"
        f'<script type="application/json" id="xlsprint-model">{_json_embed(model)}</script>'
        f"<script>{_JS}</script></body></html>\n"
    )


# ---------------------------------------------------------------------------
# Files
# ---------------------------------------------------------------------------

def _load_run(path: Path) -> Any:
    return _analysis().load_run(path)


def render(trace_paths: Iterable, formulas_path: Any, out_html: Any, *, allow_invalid: bool = False) -> Path:
    """Analyse one trace-on (and optionally one trace-off) trace and write the report.

    Fails closed: raises ReportError when a trace is invalid unless ``allow_invalid``.
    """
    runs = [_load_run(Path(p)) for p in trace_paths]
    on = [r for r in runs if _get(r, "mode") == "trace-on"]
    off = [r for r in runs if _get(r, "mode") == "trace-off"]
    other = [r for r in runs if _get(r, "mode") not in ("trace-on", "trace-off")]
    if len(on) != 1:
        raise ReportError(f"expected exactly one trace-on trace, got {len(on)}")
    if len(off) > 1:
        raise ReportError(f"expected at most one trace-off trace, got {len(off)}")
    if other:
        raise ReportError("trace with unknown mode (expected trace-on or trace-off)")

    errors = _validation_dict(on[0], "trace-on")["errors"]
    if off:
        errors += _validation_dict(off[0], "trace-off")["errors"]
        if not _get(_get(off[0], "validation"), "ok", False) and not errors:
            errors.append("trace-off: validation failed")
    if not _get(_get(on[0], "validation"), "ok", False) and not errors:
        errors.append("trace-on: validation failed")
    if errors and not allow_invalid:
        raise ReportError("refusing to render timing from an invalid trace "
                          "(use --allow-invalid to inspect it):\n  " + "\n  ".join(errors))

    formulas = None
    if formulas_path is not None:
        with open(formulas_path, "r", encoding="utf-8") as f:
            formulas = json.load(f)
        if not isinstance(formulas, dict) or formulas.get("schema") != "xlsprint.formulas/1":
            raise ReportError("formulas file has an unknown schema (expected xlsprint.formulas/1)")

    model = build_model(on[0], off[0] if off else None, formulas)
    model["validation"]["allow_invalid"] = bool(allow_invalid)
    out = Path(out_html)
    out.write_text(render_html(model), encoding="utf-8")
    return out
