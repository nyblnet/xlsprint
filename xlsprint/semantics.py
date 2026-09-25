"""Bind human-authored semantic annotations to value-free XLSprint targets.

The sidecar is deliberately JSON-only and optional. Its selectors are resolved
against the workbook's in-memory structural plan, then clear workbook names
are discarded before the formulas inventory or HTML report is written.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

SCHEMA = "xlsprint.semantics/1"
MAX_ANNOTATIONS = 500
_ID = re.compile(r"^[a-z][a-z0-9_.-]{0,79}$")
_SHA256 = re.compile(r"^[0-9a-f]{64}$")
_ADDR = re.compile(r"^[A-Za-z0-9_:$! .'-]{1,256}$")
_CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}"
_A1_RANGE = re.compile(rf"^{_CELL}(?::{_CELL})?$", re.IGNORECASE)
_TEXT_LIMITS = {"label": 120, "intent": 600, "category": 80, "source": 120}
_SPAN_KINDS = {
    "host.stage", "run.pass", "calc.full", "calc.fullrebuild", "calc.recalc",
    "calc.sheet", "calc.range", "calc.name", "calc.group", "vba.proc", "marker",
}


class SemanticsError(ValueError):
    """A semantic sidecar is malformed or belongs to a different workbook."""


def _text(value: Any, field: str, maximum: int, *, required: bool = True) -> str | None:
    if value is None and not required:
        return None
    if not isinstance(value, str) or not value.strip() or len(value) > maximum:
        raise SemanticsError(f"{field} must be non-empty text of at most {maximum} characters")
    return value.strip()


def normalize_address(value: str) -> str:
    if not isinstance(value, str) or not _ADDR.fullmatch(value):
        raise SemanticsError("range address contains unsupported characters")
    value = re.sub(r"\s+", "", value)
    if not _A1_RANGE.fullmatch(value):
        raise SemanticsError("range address must be an A1 cell or rectangular area")
    return value.replace("$", "").upper()


def load_manifest(path: str | Path) -> dict:
    """Load and strictly validate an ``xlsprint.semantics/1`` JSON sidecar."""
    try:
        with open(path, "r", encoding="utf-8") as stream:
            raw = json.load(stream)
    except (OSError, json.JSONDecodeError) as exc:
        raise SemanticsError(f"could not read semantic map: {exc}") from exc
    if not isinstance(raw, dict) or raw.get("schema") != SCHEMA:
        raise SemanticsError(f"semantic map schema must be {SCHEMA}")
    workbook_sha256 = raw.get("workbook_sha256")
    if workbook_sha256 is not None and (
        not isinstance(workbook_sha256, str) or not _SHA256.fullmatch(workbook_sha256)
    ):
        raise SemanticsError("workbook_sha256 must be 64 lowercase hexadecimal characters")
    annotations = raw.get("annotations")
    if not isinstance(annotations, list) or len(annotations) > MAX_ANNOTATIONS:
        raise SemanticsError(f"annotations must be a list of at most {MAX_ANNOTATIONS} entries")
    out = {"schema": SCHEMA, "workbook_sha256": workbook_sha256, "annotations": []}
    seen: set[str] = set()
    for index, item in enumerate(annotations):
        label = f"annotations[{index}]"
        if not isinstance(item, dict):
            raise SemanticsError(f"{label} must be an object")
        ident = item.get("id")
        if not isinstance(ident, str) or not _ID.fullmatch(ident) or ident in seen:
            raise SemanticsError(f"{label}.id must be unique and match {_ID.pattern}")
        seen.add(ident)
        fields = {key: _text(item.get(key), f"{label}.{key}", limit, required=(key in ("label", "intent")))
                  for key, limit in _TEXT_LIMITS.items()}
        match = item.get("match")
        if not isinstance(match, dict) or len(set(match) & {"defined_name", "range", "formula_group", "span"}) != 1:
            raise SemanticsError(f"{label}.match must select exactly one supported target")
        kind = next(key for key in ("defined_name", "range", "formula_group", "span") if key in match)
        selector = match[kind]
        if kind == "defined_name":
            if isinstance(selector, str):
                selector = {"name": selector}
            if not isinstance(selector, dict):
                raise SemanticsError(f"{label}.match.defined_name must be a name or object")
            selector = {"name": _text(selector.get("name"), f"{label}.match.defined_name.name", 256)}
            if selector.get("name") and ("=" in selector["name"] or '"' in selector["name"]):
                raise SemanticsError(f"{label}.match.defined_name.name contains unsupported characters")
            raw_name_match = match["defined_name"]
            if isinstance(raw_name_match, dict) and "scope" in raw_name_match:
                scope = _text(raw_name_match.get("scope"), f"{label}.match.defined_name.scope", 256)
                selector["scope"] = scope
        elif kind in ("range", "formula_group"):
            if not isinstance(selector, dict):
                raise SemanticsError(f"{label}.match.{kind} must be an object")
            selector = {
                "sheet": _text(selector.get("sheet"), f"{label}.match.{kind}.sheet", 256),
                "address": normalize_address(_text(selector.get("address"), f"{label}.match.{kind}.address", 256)),
            }
        else:
            if not isinstance(selector, dict):
                raise SemanticsError(f"{label}.match.span must be an object")
            span_kind = _text(selector.get("kind"), f"{label}.match.span.kind", 40)
            if span_kind not in _SPAN_KINDS:
                raise SemanticsError(f"{label}.match.span.kind is not a supported XLSprint span kind")
            selector = {"kind": span_kind}
            for key in ("name", "sheet"):
                if key in match["span"]:
                    selector[key] = _text(match["span"].get(key), f"{label}.match.span.{key}", 256)
            if "address" in match["span"]:
                selector["address"] = normalize_address(_text(match["span"].get("address"), f"{label}.match.span.address", 256))
        out["annotations"].append({"id": ident, **fields, "match": {kind: selector}})
    if workbook_sha256 is None and any(
        next(iter(item["match"])) != "span" for item in out["annotations"]
    ):
        raise SemanticsError("workbook_sha256 is required for defined-name, range, and formula-group selectors")
    return out


def _redact(value: str | None, names: str, salt: str) -> str | None:
    if value is None or names == "clear":
        return value
    from .trace import redact_name
    return redact_name(value, salt)


def bind_manifest(manifest: dict, *, workbook_sha256: str, formulas: dict,
                  plan: dict, names: str, salt: str) -> dict:
    """Resolve clear selectors to the run's redacted structural identifiers."""
    expected = manifest.get("workbook_sha256")
    if expected and expected != workbook_sha256:
        raise SemanticsError("semantic map workbook_sha256 does not match the profiled workbook")
    if formulas.get("workbook_sha256") not in (None, workbook_sha256):
        raise SemanticsError("formula inventory workbook hash does not match the profiled workbook")

    resolved = []
    redaction_map = plan.get("redaction_map") if isinstance(plan.get("redaction_map"), dict) else {}
    clear_by_token = {str(token): str(clear) for token, clear in redaction_map.items()}
    raw_groups = list(plan.get("groups") or [])
    for annotation in manifest["annotations"]:
        selector_kind, selector = next(iter(annotation["match"].items()))
        matches: list[dict] = []
        if selector_kind == "defined_name":
            for item in (formulas.get("names") or []):
                if not isinstance(item, dict) or clear_by_token.get(item.get("name"), item.get("name")) != selector["name"]:
                    continue
                clear_scope = clear_by_token.get(item.get("scope"), item.get("scope", "workbook"))
                if selector.get("scope") is not None and clear_scope != selector["scope"]:
                    continue
                range_ref = item.get("refers_to_range")
                address = range_ref.rsplit("!", 1)[-1] if isinstance(range_ref, str) and "!" in range_ref else None
                matches.append({"target": "defined_name", "name": item.get("name"),
                                "scope": item.get("scope", "workbook"),
                                "sheet": item.get("sheet"),
                                "address": normalize_address(address) if address else None,
                                "span_kind": "calc.name" if address else None})
        elif selector_kind == "range":
            sheets = {item.get("sheet") for item in (plan.get("sheets") or []) if isinstance(item, dict)}
            if selector["sheet"] in sheets:
                matches.append({"target": "range", "sheet": _redact(selector["sheet"], names, salt),
                                "address": selector["address"],
                                "span_kinds": ["calc.range", "calc.name", "calc.group"]})
        elif selector_kind == "formula_group":
            for item in raw_groups:
                if item.get("sheet") != selector["sheet"]:
                    continue
                areas = [normalize_address(area) for area in (item.get("areas") or [])]
                if selector["address"] not in areas:
                    continue
                matches.append({"target": "formula_group", "group": item.get("group"),
                                "sheet": _redact(item.get("sheet"), names, salt),
                                "address": selector["address"], "span_kind": "calc.group"})
        else:
            item = dict(selector)
            item["target"] = "span"
            if "name" in item:
                item["name"] = _redact(item["name"], names, salt)
            if "sheet" in item:
                item["sheet"] = _redact(item["sheet"], names, salt)
            matches.append(item)
        resolved.append({key: annotation[key] for key in ("id", "label", "intent", "category", "source") if annotation.get(key)
                         } | {"targets": matches, "binding": "bound" if matches else "unmatched"})
    return {"schema": SCHEMA, "workbook_sha256": workbook_sha256, "workbook_bound": bool(expected),
            "annotations": resolved,
            "unbound_count": sum(item["binding"] == "unmatched" for item in resolved)}
