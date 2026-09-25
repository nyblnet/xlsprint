"""Trace schema, bounded writer, reader, and fail-closed validator.

A trace is JSON Lines: a header line, then B/E/M events, then a footer line
(see docs/DESIGN.md, "Trace format"). Every number in a trace is a raw clock
reading in integer nanoseconds; nothing here is derived.

Privacy is enforced here, not by callers: attribute keys come from a closed
allow-list and attribute string values must match ``ATTR_VALUE_RE``, which
excludes ``=`` and ``"`` so formula text and string literals cannot pass.
"""

from __future__ import annotations

import contextlib
import hashlib
import json
import math
import os
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Tuple

SCHEMA = "xlsprint.trace/1"

SPAN_KINDS = frozenset(
    {
        "host.stage",
        "run.pass",
        "calc.full",
        "calc.fullrebuild",
        "calc.recalc",
        "calc.sheet",
        "calc.range",
        "calc.name",
        "calc.group",
        "vba.proc",
        "marker",
    }
)

# Clock each kind must use. "marker" may use either clock.
KIND_CLOCK = {
    "host.stage": "host",
    "run.pass": "vba",
    "calc.full": "vba",
    "calc.fullrebuild": "vba",
    "calc.recalc": "vba",
    "calc.sheet": "vba",
    "calc.range": "vba",
    "calc.name": "vba",
    "calc.group": "vba",
    "vba.proc": "vba",
}

HOST_STAGES = frozenset(
    {
        "prepare_copy",
        "launch_excel",
        "open_workbook",
        "instrument",
        "warmup",
        "profile_trace_off",
        "profile_trace_on",
        "collect_trace",
        "close_excel",
        "verify_output",
        "render_report",
    }
)

# Kinds a trace-off run records (plus host stages and markers).
TRACE_OFF_KINDS = frozenset({"run.pass", "calc.full", "calc.recalc", "host.stage", "marker"})

ALLOWED_ATTR_KEYS = frozenset(
    {
        "method",
        "repeat",
        "cells",
        "areas",
        "address",
        "sheet",
        "group",
        "calc_mode",
        "threads",
        "multithreaded",
        "iteration",
        "error_code",
        "note",
        "count",
        "overhead_mode",
        "bytes",
        "events",
        "expected",
        "found",
        "sha256_prefix",
    }
)

# DESIGN.md's class. \w is Unicode-aware so non-ASCII sheet names pass; '='
# and '"' stay excluded. \Z (not $) so a trailing newline cannot slip through
# callers that use .match(); check_attr uses fullmatch anyway.
ATTR_VALUE_RE = re.compile(r"^[\w .:$!'#\-/()\[\],@]*\Z")
ATTR_MAX_LEN = 256

# Rule 10: structural identifiers only. Excludes '=', '"', '*', '+', '^', '<',
# '>' and control characters, so formula text cannot pass as a name. The
# runner hashes any clear identifier that fails this (see check_name).
NAME_RE = re.compile(r"^[\w .:$!'#\-/()\[\],@&%]*\Z")
NAME_MAX_LEN = 256

NS_MAX = 2**63  # ns must satisfy 0 <= ns < NS_MAX

CLOCKS = frozenset({"host", "vba"})
STATUSES = frozenset({"ok", "error", "aborted"})
MODES = frozenset({"trace-on", "trace-off"})
SOURCE_KINDS = frozenset({"synthetic", "real"})

B_KEYS = frozenset({"type", "id", "parent", "pass", "depth", "ns", "clock", "kind", "name", "attrs"})
M_KEYS = frozenset({"type", "id", "parent", "pass", "depth", "ns", "clock", "kind", "name", "attrs"})
E_KEYS = frozenset({"type", "id", "ns", "clock", "status", "attrs"})

DEFAULT_MAX_EVENTS = 200_000
DEFAULT_MAX_BYTES = 64 * 2**20
# Bytes held back for the footer line when enforcing max_bytes; a footer with
# 20-digit counters is under 200 bytes.
FOOTER_RESERVE_BYTES = 256

HEADER_OBJECTS = ("host", "clock", "source", "redaction", "limits")

_MAX_REPORTED_ERRORS = 200


class TraceError(Exception):
    """A trace could not be written or read without violating the schema."""


def redact_name(name: str, salt: str) -> str:
    """Return ``"h:" + sha256(salt + name)[:10]`` (UTF-8)."""
    digest = hashlib.sha256((salt + name).encode("utf-8")).hexdigest()
    return "h:" + digest[:10]


# --------------------------------------------------------------------------
# Field checks shared by the writer and the validator. Each returns an error
# string or None.


def _is_int(v: Any) -> bool:
    return isinstance(v, int) and not isinstance(v, bool)


def _in(v: Any, vocab: frozenset) -> bool:
    # Membership that cannot raise on unhashable JSON values (lists, objects).
    return isinstance(v, str) and v in vocab


def _is_ns(v: Any) -> bool:
    return _is_int(v) and 0 <= v < NS_MAX


def check_attr(key: Any, value: Any) -> Optional[str]:
    if not isinstance(key, str) or key not in ALLOWED_ATTR_KEYS:
        return "attribute key %r is not allow-listed" % (key,)
    if isinstance(value, bool) or _is_int(value):
        return None
    if isinstance(value, float):
        return None if math.isfinite(value) else "attribute %r is not a finite number" % key
    if isinstance(value, str):
        if len(value) > ATTR_MAX_LEN:
            return "attribute %r value longer than %d characters" % (key, ATTR_MAX_LEN)
        if not ATTR_VALUE_RE.fullmatch(value):
            return "attribute %r value fails the allowed pattern" % key
        return None
    return "attribute %r has disallowed type %s" % (key, type(value).__name__)


def check_attrs(attrs: Any) -> List[str]:
    if attrs is None:
        return []
    if not isinstance(attrs, dict):
        return ["attrs must be an object"]
    return [e for e in (check_attr(k, v) for k, v in attrs.items()) if e]


def check_name(name: Any) -> Optional[str]:
    """Rule 10: return None if ``name`` may appear in a trace, else the reason.

    Names are structural identifiers only (sheet, address, defined name,
    procedure, stage, group id): a string of at most NAME_MAX_LEN characters
    that fully matches NAME_RE. Callers use this to decide whether a clear
    identifier must be hashed with ``redact_name`` (whose output always passes).
    """
    if not isinstance(name, str):
        return "name must be a string"
    if len(name) > NAME_MAX_LEN:
        return "name longer than %d characters" % NAME_MAX_LEN
    if not NAME_RE.fullmatch(name):
        return "name fails the allowed pattern"
    return None


def _check_kind_clock(kind: Any, clock: Any) -> Optional[str]:
    if not _in(kind, SPAN_KINDS):
        return "unknown kind %r" % (kind,)
    if not _in(clock, CLOCKS):
        return "unknown clock %r" % (clock,)
    want = KIND_CLOCK.get(kind)
    if want is not None and clock != want:
        return "kind %s must use clock %r, not %r" % (kind, want, clock)
    return None


def _check_header(header: Any) -> List[str]:
    """Header errors, each prefixed with its DESIGN.md rule number."""
    if not isinstance(header, dict) or header.get("type") != "header":
        return ["[1] header missing"]
    errs = []
    if header.get("schema") != SCHEMA:
        errs.append("[1] unknown schema %r" % (header.get("schema"),))
    if not _in(header.get("mode"), MODES):
        errs.append("[1] header mode must be one of %s" % sorted(MODES))
    run_id = header.get("run_id")
    if not isinstance(run_id, str) or not run_id:
        errs.append("[9] header run_id must be a non-empty string")
    for key in HEADER_OBJECTS:
        if key in header and not isinstance(header[key], dict):
            errs.append("[9] header %s must be an object" % key)
    source = header.get("source")
    if not isinstance(source, dict) or not _in(source.get("kind"), SOURCE_KINDS):
        errs.append("[1] header source.kind is mandatory and must be 'synthetic' or 'real'")
    elif "label" in source and not isinstance(source["label"], str):
        errs.append("[9] header source.label must be a string")
    clock = header.get("clock")
    if isinstance(clock, dict):
        off = clock.get("vba_offset_ns")
        if off is not None and not (_is_int(off) and -NS_MAX < off < NS_MAX):
            errs.append("[9] header clock.vba_offset_ns must be an integer")
        unc = clock.get("vba_offset_uncertainty_ns")
        if unc is not None and not _is_ns(unc):
            errs.append("[9] header clock.vba_offset_uncertainty_ns must be a non-negative integer")
    limits = header.get("limits")
    if isinstance(limits, dict):
        for k in ("max_events", "max_bytes"):
            if k in limits and not (_is_int(limits[k]) and limits[k] > 0):
                errs.append("[9] header limits.%s must be a positive integer" % k)
    red = header.get("redaction")
    if isinstance(red, dict) and "names" in red and not _in(red["names"], frozenset({"clear", "hashed"})):
        errs.append("[9] header redaction.names must be 'clear' or 'hashed'")
    expect = _header_expect(header)
    if expect is not None:
        ok = isinstance(expect, dict) and all(
            _in(k, SPAN_KINDS) and isinstance(v, list) and all(isinstance(n, str) for n in v) for k, v in expect.items()
        )
        if not ok:
            errs.append("[9] header expect must map span kinds to lists of strings")
    return errs


def _header_expect(header: dict) -> Any:
    # Older DESIGN.md text said "expected"; accept it as an alias of "expect".
    if "expect" in header:
        return header["expect"]
    return header.get("expected")


def _dump(obj: dict) -> str:
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"), allow_nan=False)


# --------------------------------------------------------------------------
# Writer


class TraceWriter:
    """Bounded writer that keeps begin/end balanced and nesting consistent.

    Events stream to ``<path>.part`` as they arrive, so a crash leaves
    diagnostics behind. ``close()`` writes the final file: header (which may be
    amended until then with ``update_header``, e.g. once the clock probe has
    produced ``vba_offset_ns``), all written events, and the footer.

    Bounds: the final file (header + events + footer) never exceeds
    ``max_bytes`` and holds at most ``max_events`` events. Once either would
    be exceeded, every further event is dropped and counted, and the footer
    says ``truncated: true``. If the header grows after events were written
    (``update_header``), close() drops trailing events to stay in bounds.
    Nesting is still tracked for dropped events so callers see the same
    begin/end contract either way. The header's ``limits`` always records the
    bounds this writer enforced.
    """

    def __init__(self, path, header: dict, *, max_events=DEFAULT_MAX_EVENTS, max_bytes=DEFAULT_MAX_BYTES):
        self.path = Path(path)
        self.max_events = int(max_events)
        self.max_bytes = int(max_bytes)
        hdr = {"type": "header", "schema": SCHEMA}
        hdr.update(json.loads(json.dumps(header)))  # private deep copy
        hdr["type"], hdr["schema"] = "header", SCHEMA
        if self.max_events < 1 or self.max_bytes < 1:
            raise TraceError("max_events and max_bytes must be positive")
        hdr["limits"] = {"max_events": self.max_events, "max_bytes": self.max_bytes}
        hdr.setdefault("redaction", {"names": "hashed"})
        self._header = hdr
        self._set_header_bytes()
        self._next_id = 1
        self._used_ids: set = set()
        # Open-span stack entries: (id, kind, clock, begin_ns, pass).
        self._stack: List[Tuple[int, str, str, int, int]] = []
        self._last_ns: Dict[str, int] = {}
        self._written = 0
        self._dropped = 0
        self._truncated = False
        self._event_bytes = 0
        self._closed = False
        self._footer: Optional[dict] = None
        self._part_path = self.path.with_name(self.path.name + ".part")
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._part = open(self._part_path, "w", encoding="utf-8", newline="\n")

    # -- properties --------------------------------------------------------

    @property
    def header(self) -> dict:
        return self._header

    @property
    def open_span_id(self) -> int:
        """Innermost open span id, or 0."""
        return self._stack[-1][0] if self._stack else 0

    @property
    def open_depth(self) -> int:
        """Number of open spans (the depth the next begin() will get)."""
        return len(self._stack)

    @property
    def truncated(self) -> bool:
        return self._truncated

    def update_header(self, fields: dict) -> None:
        """Merge ``fields`` into the header (one level deep for nested objects).

        The header is serialized at close(), so values known only mid-run
        (``clock.vba_offset_ns``, Excel version) can be filled in late.
        """
        self._require_open()
        fields = json.loads(json.dumps(fields))
        for k in fields:
            if k in ("type", "schema", "limits"):
                raise TraceError("header field %r is set by the writer" % k)
        before = json.loads(json.dumps(self._header))
        for k, v in fields.items():
            if isinstance(v, dict) and isinstance(self._header.get(k), dict):
                self._header[k].update(v)
            else:
                self._header[k] = v
        try:
            self._set_header_bytes()
        except TraceError:
            self._header = before
            raise

    def _set_header_bytes(self) -> None:
        errs = _check_header(self._header)
        if errs:
            raise TraceError("; ".join(errs))
        size = len(_dump(self._header).encode("utf-8")) + 1
        if size + FOOTER_RESERVE_BYTES > self.max_bytes:
            raise TraceError("header alone exceeds max_bytes")
        self._header_bytes = size

    # -- id management -----------------------------------------------------

    def reserve_ids(self, n: int) -> int:
        """Reserve ``n`` consecutive ids (e.g. for VBA) and return the first."""
        if not _is_int(n) or n < 1:
            raise TraceError("reserve_ids needs a positive count")
        self._require_open()
        first = self._next_id
        self._next_id += n
        return first

    def _alloc_id(self) -> int:
        while self._next_id in self._used_ids:
            self._next_id += 1
        i = self._next_id
        self._next_id += 1
        self._used_ids.add(i)
        return i

    # -- primitive checks --------------------------------------------------

    def _require_open(self) -> None:
        if self._closed:
            raise TraceError("trace writer is closed")

    def _resolve_ns(self, clock: str, ns: Optional[int]) -> int:
        if ns is None:
            if clock != "host":
                raise TraceError("ns is mandatory for clock=%r" % clock)
            ns = time.perf_counter_ns()
        if not _is_ns(ns):
            raise TraceError("ns must be an integer with 0 <= ns < 2**63")
        last = self._last_ns.get(clock)
        if last is not None and ns < last:
            raise TraceError("ns went backwards on clock %r (%d < %d)" % (clock, ns, last))
        return ns

    @staticmethod
    def _check_attrs_or_raise(attrs: Any) -> dict:
        errs = check_attrs(attrs)
        if errs:
            raise TraceError("; ".join(errs))
        return dict(attrs or {})

    def _emit(self, ev: dict) -> None:
        self._last_ns[ev["clock"]] = ev["ns"]
        if self._truncated:
            self._dropped += 1
            return
        line = _dump(ev) + "\n"
        size = len(line.encode("utf-8"))
        budget = self.max_bytes - self._header_bytes - FOOTER_RESERVE_BYTES
        if self._written + 1 > self.max_events or self._event_bytes + size > budget:
            self._truncated = True
            self._dropped += 1
            return
        self._part.write(line)
        self._written += 1
        self._event_bytes += size

    # -- public event API --------------------------------------------------

    def begin(self, kind, name, *, clock="host", pass_=0, attrs=None, ns=None) -> int:
        """Open a span and return its id. VBA-clock spans require ``ns``."""
        self._require_open()
        err = _check_kind_clock(kind, clock) or check_name(name)
        if kind == "marker":
            err = "use marker() for kind 'marker'"
        if err:
            raise TraceError(err)
        if not _is_int(pass_) or pass_ < 0:
            raise TraceError("pass must be a non-negative integer")
        attrs = self._check_attrs_or_raise(attrs)
        ns = self._resolve_ns(clock, ns)
        span_id = self._alloc_id()
        ev = {
            "type": "B",
            "id": span_id,
            "parent": self.open_span_id,
            "pass": pass_,
            "depth": self.open_depth,
            "ns": ns,
            "clock": clock,
            "kind": kind,
            "name": name,
            "attrs": attrs,
        }
        self._stack.append((span_id, kind, clock, ns, pass_))
        self._emit(ev)
        return span_id

    def end(self, span_id, *, status="ok", attrs=None, ns=None) -> None:
        """Close the innermost open span. Raises TraceError on anything else."""
        self._require_open()
        if not self._stack:
            raise TraceError("end(%r) with no open span" % (span_id,))
        top_id, _kind, clock, begin_ns, _pass = self._stack[-1]
        if span_id != top_id:
            raise TraceError("end(%r) is not the innermost open span (%d)" % (span_id, top_id))
        if not _in(status, STATUSES):
            raise TraceError("unknown status %r" % (status,))
        attrs = self._check_attrs_or_raise(attrs)
        ns = self._resolve_ns(clock, ns)
        if ns < begin_ns:
            raise TraceError("span %d ends before it begins" % span_id)
        ev = {"type": "E", "id": span_id, "ns": ns, "clock": clock, "status": status}
        if attrs:
            ev["attrs"] = attrs
        self._stack.pop()
        self._emit(ev)

    def marker(self, name, *, clock="host", pass_=0, attrs=None, ns=None) -> int:
        """Write an instantaneous marker and return its id."""
        self._require_open()
        err = _check_kind_clock("marker", clock) or check_name(name)
        if err:
            raise TraceError(err)
        if not _is_int(pass_) or pass_ < 0:
            raise TraceError("pass must be a non-negative integer")
        attrs = self._check_attrs_or_raise(attrs)
        ns = self._resolve_ns(clock, ns)
        mid = self._alloc_id()
        ev = {
            "type": "M",
            "id": mid,
            "parent": self.open_span_id,
            "pass": pass_,
            "ns": ns,
            "clock": clock,
            "kind": "marker",
            "name": name,
            "attrs": attrs,
        }
        self._emit(ev)
        return mid

    @contextlib.contextmanager
    def span(self, kind, name, **kw) -> Iterator[int]:
        """Context manager around begin/end.

        Status is ``"error"`` if the body raises an Exception and ``"aborted"``
        for other BaseExceptions (e.g. KeyboardInterrupt); the exception is
        always re-raised. ``ns`` is not accepted: both ends read the host clock.
        """
        if kw.get("clock", "host") != "host" or "ns" in kw:
            raise TraceError("span() is for host-clock spans; use begin/end for explicit ns")
        span_id = self.begin(kind, name, **kw)
        try:
            yield span_id
        except Exception:
            self.end(span_id, status="error")
            raise
        except BaseException:
            self.end(span_id, status="aborted")
            raise
        self.end(span_id)

    def append_event(self, ev: dict) -> None:
        """Import a pre-built B/E/M event (e.g. from the VBA event file).

        Checks keys, id uniqueness, kind/clock, attrs, monotonic ns, and
        nesting. ``parent`` in the imported event is relative to its producer:
        0 (or absent) means "no parent there" and is rebased onto the writer's
        innermost open span; a non-zero parent must equal that span. ``depth``
        is always recomputed from the writer's stack.
        """
        self._require_open()
        if not isinstance(ev, dict):
            raise TraceError("event must be an object")
        typ = ev.get("type")
        if typ in ("B", "M"):
            extra = set(ev) - (B_KEYS if typ == "B" else M_KEYS)
            if extra:
                raise TraceError("event has disallowed keys %s" % sorted(extra))
            kind = ev.get("kind", "marker" if typ == "M" else None)
            clock = ev.get("clock")
            err = _check_kind_clock(kind, clock) or check_name(ev.get("name"))
            if not err and (typ == "M") != (kind == "marker"):
                err = "M events must have kind 'marker' and B events must not"
            if err:
                raise TraceError(err)
            eid = ev.get("id")
            if not _is_int(eid) or eid < 1:
                raise TraceError("id must be a positive integer")
            if eid in self._used_ids:
                raise TraceError("duplicate id %d" % eid)
            pass_ = ev.get("pass", 0)
            if not _is_int(pass_) or pass_ < 0:
                raise TraceError("pass must be a non-negative integer")
            parent = ev.get("parent", 0)
            if not _is_int(parent) or (parent != 0 and parent != self.open_span_id):
                raise TraceError("parent %r is not the innermost open span %d" % (parent, self.open_span_id))
            attrs = self._check_attrs_or_raise(ev.get("attrs"))
            ns = ev.get("ns")
            if ns is None:
                raise TraceError("imported events must carry ns")
            ns = self._resolve_ns(clock, ns)
            self._used_ids.add(eid)
            if eid >= self._next_id:
                self._next_id = eid + 1
            out = {"type": typ, "id": eid, "parent": self.open_span_id, "pass": pass_}
            if typ == "B":
                out["depth"] = self.open_depth
            out.update({"ns": ns, "clock": clock, "kind": kind, "name": ev["name"], "attrs": attrs})
            if typ == "B":
                self._stack.append((eid, kind, clock, ns, pass_))
            self._emit(out)
        elif typ == "E":
            extra = set(ev) - E_KEYS
            if extra:
                raise TraceError("E event has disallowed keys %s" % sorted(extra))
            if not _in(ev.get("clock"), CLOCKS):
                raise TraceError("unknown clock %r" % (ev.get("clock"),))
            if self._stack and self._stack[-1][2] != ev.get("clock"):
                raise TraceError("E clock differs from its B clock")
            if ev.get("ns") is None:
                raise TraceError("imported events must carry ns")
            self.end(ev.get("id"), status=ev.get("status", "ok"), attrs=ev.get("attrs"), ns=ev["ns"])
        else:
            raise TraceError("unknown event type %r" % (typ,))

    def close(self) -> dict:
        """Write the final file and return the footer.

        Open spans are reported as ``open_spans_at_close``; no E events are
        fabricated for them, so such a trace fails validation (rule 2).
        """
        if self._closed:
            return dict(self._footer or {})
        self._closed = True
        self._part.close()
        header_line = _dump(self._header) + "\n"
        header_bytes = len(header_line.encode("utf-8"))
        budget = self.max_bytes - header_bytes - FOOTER_RESERVE_BYTES
        tmp = self.path.with_name(self.path.name + ".tmp")
        kept = kept_bytes = 0
        with open(tmp, "w", encoding="utf-8", newline="\n") as out, open(self._part_path, "r", encoding="utf-8") as part:
            out.write(header_line)
            for line in part:
                size = len(line.encode("utf-8"))
                if kept_bytes + size > budget:
                    # Only reachable when the header grew after events were written.
                    self._truncated = True
                    self._dropped += self._written - kept
                    break
                out.write(line)
                kept += 1
                kept_bytes += size
            footer = {
                "type": "footer",
                "events": kept,
                "dropped": self._dropped,
                "truncated": self._truncated,
                "open_spans_at_close": len(self._stack),
                "bytes": header_bytes + kept_bytes,
            }
            out.write(_dump(footer) + "\n")
        os.replace(tmp, self.path)
        os.remove(self._part_path)
        self._footer = footer
        return dict(footer)

    def __enter__(self) -> "TraceWriter":
        return self

    def __exit__(self, *exc) -> None:
        self.close()


# --------------------------------------------------------------------------
# Readers


def _parse_lines(path) -> Tuple[List[Tuple[int, Any]], List[str], int]:
    """Return ([(line_no, obj)], parse_errors, file size in bytes)."""
    raw = Path(path).read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]
    out, errs = [], []
    for no, line in enumerate(raw.decode("utf-8", errors="replace").splitlines(), 1):
        if not line.strip():
            continue
        try:
            out.append((no, json.loads(line)))
        except ValueError:
            errs.append("line %d: not valid JSON" % no)
    return out, errs, len(raw)


def _split(objs: List[Tuple[int, Any]]):
    header = footer = None
    body = list(objs)
    if body and isinstance(body[0][1], dict) and body[0][1].get("type") == "header":
        header = body.pop(0)[1]
    if body and isinstance(body[-1][1], dict) and body[-1][1].get("type") == "footer":
        footer = body.pop()[1]
    return header, body, footer


def read_trace(path) -> Tuple[dict, List[dict], Optional[dict]]:
    """Return (header, events, footer). Raises TraceError without a header or on bad JSON.

    This does not validate; use ``validate`` for that.
    """
    objs, errs, _ = _parse_lines(path)
    if errs:
        raise TraceError(errs[0])
    header, body, footer = _split(objs)
    if header is None:
        raise TraceError("trace has no header")
    return header, [o for _, o in body], footer


def read_vba_events(path) -> Tuple[List[dict], dict]:
    """Read the VBA-side event file: JSON Lines of B/E/M events, last line a vba_footer.

    Returns (events, {"type": "vba_footer", "events", "dropped", "truncated",
    "open_spans"}). Tolerates a UTF-8 BOM, CRLF, and integral float ``ns``
    values (VBA's CDec formatting). Raises TraceError without a vba_footer:
    a missing footer means the VBA side did not finish, which is fail-closed.
    """
    objs, errs, _ = _parse_lines(path)
    if errs:
        raise TraceError(errs[0])
    if not objs or not isinstance(objs[-1][1], dict) or objs[-1][1].get("type") != "vba_footer":
        raise TraceError("VBA event file has no vba_footer (incomplete or truncated)")
    footer_raw = objs.pop()[1]
    # Extra VBA-side keys (faults, attr_rejected, mode, pass, ...) are kept
    # for the caller; the five contract keys are normalized.
    footer = dict(footer_raw)
    footer.update({
        "type": "vba_footer",
        "events": footer_raw.get("events"),
        "dropped": footer_raw.get("dropped", 0),
        "truncated": bool(footer_raw.get("truncated", False)),
        "open_spans": footer_raw.get("open_spans", 0),
    })
    events = []
    for no, ev in objs:
        if not isinstance(ev, dict) or ev.get("type") not in ("B", "E", "M"):
            raise TraceError("line %d: not a B/E/M event" % no)
        ns = ev.get("ns")
        if isinstance(ns, float):
            if not math.isfinite(ns) or ns != int(ns):
                raise TraceError("line %d: ns is not an integer" % no)
            ev["ns"] = int(ns)
        events.append(ev)
    if footer["events"] is not None and footer["events"] != len(events):
        raise TraceError("vba_footer events=%r but %d events were read" % (footer["events"], len(events)))
    return events, footer


# --------------------------------------------------------------------------
# Validator


@dataclass
class ValidationResult:
    ok: bool
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    stats: dict = field(default_factory=dict)


def validate(path) -> ValidationResult:
    """Apply the fail-closed rules 1-10 of DESIGN.md to a trace file.

    ``ok`` is True only when ``errors`` is empty; each error starts with its
    rule number, e.g. ``[5]``. ``stats`` summarizes what was observed (event
    counts per type/kind/clock, passes, open spans). Never raises on a
    malformed trace: anything unexpected becomes an error.
    """
    try:
        return _validate(path)
    except Exception as e:  # fail closed; callers must always get a result
        return ValidationResult(False, ["[1] trace could not be validated: %s: %s" % (type(e).__name__, e)], [], {})


_EVENT_TYPES = frozenset({"B", "E", "M"})
_SPAN_KINDS_IN_PASS = frozenset(k for k in SPAN_KINDS if k.startswith("calc.")) | {"vba.proc"}


def _validate(path) -> ValidationResult:
    errors: List[str] = []
    warnings: List[str] = []
    stats: Dict[str, Any] = {}
    err = errors.append

    try:
        objs, parse_errs, total_bytes = _parse_lines(path)
    except OSError as e:
        return ValidationResult(False, ["[1] cannot read trace: %s" % e], [], {})
    for e in parse_errs:
        err("[1] " + e)
    header, body, footer = _split(objs)

    # Rules 1 and 9 (header part).
    if header is None:
        err("[1] header missing")
        header = {}
        header_errs = []
    else:
        header_errs = _check_header(header)
        errors.extend(header_errs)
    if footer is None:
        err("[1] footer missing")
    mode = header.get("mode") if _in(header.get("mode"), MODES) else None
    source = header.get("source")
    stats["mode"] = mode
    stats["source_kind"] = source.get("kind") if isinstance(source, dict) and _in(source.get("kind"), SOURCE_KINDS) else None
    clock_hdr = header.get("clock") if isinstance(header.get("clock"), dict) else {}
    offset = clock_hdr.get("vba_offset_ns")
    offset = offset if _is_int(offset) and -NS_MAX < offset < NS_MAX else None
    unc = clock_hdr.get("vba_offset_uncertainty_ns")
    unc = unc if _is_ns(unc) else 0

    # Walk events: rules 2, 3, 4, 6, 9, 10.
    stack: List[dict] = []
    seen_ids: set = set()
    last_ns: Dict[str, Tuple[int, int]] = {}
    spans: Dict[int, dict] = {}  # B events by id
    by_type = {"B": 0, "E": 0, "M": 0}
    by_kind: Dict[str, int] = {}
    by_clock: Dict[str, int] = {"host": 0, "vba": 0}
    run_pass_of: Dict[int, int] = {}  # span id -> enclosing run.pass span id (or 0)
    pass_numbers: Dict[int, int] = {}  # run.pass span id -> pass number

    for no, ev in body:
        where = "line %d" % no
        if not isinstance(ev, dict):
            err("[1] %s: event is not an object" % where)
            continue
        typ = ev.get("type")
        if not _in(typ, _EVENT_TYPES):
            err("[1] %s: unexpected record type %r" % (where, typ))
            continue
        by_type[typ] += 1

        extra = set(ev) - (E_KEYS if typ == "E" else B_KEYS)
        if extra:
            err("[6] %s: %s event has disallowed keys %s" % (where, typ, sorted(extra)))

        clock = ev.get("clock") if _in(ev.get("clock"), CLOCKS) else None
        if clock is None:
            err("[3] %s: unknown clock %r" % (where, ev.get("clock")))
        else:
            by_clock[clock] += 1
        ns = ev.get("ns")
        if not _is_ns(ns):
            err("[9] %s: ns must be an integer with 0 <= ns < 2**63" % where)
            ns = None
        elif clock is not None:
            prev = last_ns.get(clock)
            if prev is not None and ns < prev[0]:
                err("[3] %s: ns went backwards on clock %s (%d < %d at line %d)" % (where, clock, ns, prev[0], prev[1]))
            last_ns[clock] = (ns, no)

        for e in check_attrs(ev.get("attrs")):
            err("[6] %s: %s" % (where, e))

        eid = ev.get("id")
        if not _is_int(eid) or eid < 1:
            err("[4] %s: id must be a positive integer" % where)
            continue

        if typ == "E":
            if not _in(ev.get("status"), STATUSES):
                err("[2] %s: E has unknown status %r" % (where, ev.get("status")))
            if not stack:
                err("[2] %s: E id=%d has no matching open B" % (where, eid))
                continue
            if stack[-1]["id"] != eid:
                if any(s["id"] == eid for s in stack):
                    err("[2] %s: E id=%d closes a span that is not innermost (innermost is %d)" % (where, eid, stack[-1]["id"]))
                    # Recover: pop through the named span so later lines are judged sanely.
                    while stack and stack[-1]["id"] != eid:
                        stack.pop()
                    stack.pop()
                else:
                    err("[2] %s: E id=%d has no matching open B" % (where, eid))
                continue
            b = stack.pop()
            if clock is not None and clock != b["clock"]:
                err("[3] %s: E id=%d clock %r differs from its B clock %r" % (where, eid, clock, b["clock"]))
            if ns is not None and b["ns"] is not None and ns < b["ns"]:
                err("[3] %s: E id=%d ends before its B (%d < %d)" % (where, eid, ns, b["ns"]))
            b["end_ns"] = ns
            hp = b.get("host_parent")
            if hp and ns is not None and b["ns"] is not None:
                spans[hp]["vba_kids"].append((b["id"], b["ns"], ns))
            if b["clock"] == "host" and b["vba_kids"] and offset is not None and ns is not None and b["ns"] is not None:
                # Rule 9: aligned VBA children must fall inside their host parent.
                lo, hi = b["ns"] - unc, ns + unc
                bad = [k for k, s0, s1 in b["vba_kids"] if s0 - offset < lo or s1 - offset > hi]
                if bad:
                    err("[9] %s: %d VBA span(s) (first id=%d) lie outside host span id=%d after clock alignment "
                        "(offset %d +/- %d ns); the offset or its sign is wrong" % (where, len(bad), bad[0], eid, offset, unc))
            continue

        # B or M
        if eid in seen_ids:
            err("[4] %s: duplicate id %d" % (where, eid))
        seen_ids.add(eid)
        kind = ev.get("kind", "marker" if typ == "M" else None)
        kind_key = kind if isinstance(kind, str) else "<invalid>"
        e = _check_kind_clock(kind, clock) if clock is not None else (None if _in(kind, SPAN_KINDS) else "unknown kind %r" % (kind,))
        if e:
            err("[1] %s: %s" % (where, e))
        if (typ == "M") != (kind == "marker"):
            err("[1] %s: M events must have kind 'marker' and B events must not" % where)
        name = ev.get("name")
        ne = check_name(name)
        if ne:
            err("[10] %s: %s" % (where, ne))
            name = None
        by_kind[kind_key] = by_kind.get(kind_key, 0) + 1

        expected_parent = stack[-1]["id"] if stack else 0
        parent = ev.get("parent")
        if not _is_int(parent) or parent != expected_parent:
            err("[4] %s: id=%d parent %r inconsistent with open-span stack (expected %d)" % (where, eid, parent, expected_parent))
        depth = ev.get("depth")
        if typ == "B" and (not _is_int(depth) or depth != len(stack)):
            err("[4] %s: id=%d depth %r inconsistent with open-span stack (expected %d)" % (where, eid, depth, len(stack)))
        if typ == "M" and depth is not None and (not _is_int(depth) or depth != len(stack)):
            err("[4] %s: marker id=%d depth %r inconsistent with open-span stack" % (where, eid, depth))

        pass_ = ev.get("pass")
        if not _is_int(pass_) or pass_ < 0:
            err("[4] %s: id=%d pass must be a non-negative integer" % (where, eid))
            pass_ = None
        enclosing_rp = 0
        for s in reversed(stack):
            if s["kind"] == "run.pass":
                enclosing_rp = s["id"]
                break
        if kind == "run.pass":
            if enclosing_rp:
                err("[4] %s: run.pass id=%d nested inside another run.pass" % (where, eid))
            if pass_ is not None and pass_ < 1:
                err("[4] %s: run.pass id=%d must have pass >= 1" % (where, eid))
            elif pass_ is not None and pass_ in pass_numbers.values():
                err("[4] %s: run.pass pass=%r appears twice" % (where, pass_))
            pass_numbers[eid] = pass_
        elif enclosing_rp and pass_ != pass_numbers.get(enclosing_rp):
            err("[9] %s: id=%d pass %r differs from enclosing run.pass pass %r" % (where, eid, pass_, pass_numbers.get(enclosing_rp)))
        elif not enclosing_rp and _in(kind, _SPAN_KINDS_IN_PASS) and typ == "B":
            err("[9] %s: %s id=%d is not inside any run.pass" % (where, kind, eid))
        run_pass_of[eid] = enclosing_rp

        if typ == "B":
            rec = {"id": eid, "kind": kind_key, "name": name, "clock": clock, "ns": ns, "end_ns": None, "vba_kids": []}
            parent_rec = spans.get(expected_parent)
            if clock == "vba" and parent_rec is not None and parent_rec["clock"] == "host":
                rec["host_parent"] = expected_parent
            spans[eid] = rec
            stack.append(rec)

    for s in stack:
        err("[2] span id=%d (%s %r) still open at footer" % (s["id"], s["kind"], s["name"]))

    if by_clock["vba"] and offset is None:
        err("[1] header clock.vba_offset_ns is required when VBA events are present")

    # Rule 5: truncation, drops, footer count. Rule 9: header.limits.
    if footer is not None:
        if footer.get("truncated") is not False:
            err("[5] footer reports truncated=%r" % (footer.get("truncated"),))
        dropped = footer.get("dropped")
        if not _is_int(dropped) or dropped != 0:
            err("[5] footer reports dropped=%r events" % (dropped,))
        fev = footer.get("events")
        if not _is_int(fev) or fev != len(body):
            err("[5] footer events=%r but %d events were observed" % (fev, len(body)))
        osc = footer.get("open_spans_at_close")
        if osc is not None and (not _is_int(osc) or osc != len(stack)):
            warnings.append("footer open_spans_at_close=%r but %d spans are open" % (osc, len(stack)))
        if _is_int(footer.get("bytes")):
            footer_line = len((_dump(footer) + "\n").encode("utf-8"))
            if footer["bytes"] != total_bytes - footer_line:
                warnings.append("footer bytes=%d differs from observed %d" % (footer["bytes"], total_bytes - footer_line))
    limits = header.get("limits") if isinstance(header.get("limits"), dict) else {}
    max_events, max_bytes = limits.get("max_events"), limits.get("max_bytes")
    if _is_int(max_events) and len(body) > max_events:
        err("[9] %d events exceed header limits.max_events=%d" % (len(body), max_events))
    if _is_int(max_bytes) and total_bytes > max_bytes:
        err("[9] file size %d bytes exceeds header limits.max_bytes=%d" % (total_bytes, max_bytes))

    # Rule 7: expected instrumentation (only when `expect` is well formed).
    expect = _header_expect(header)
    if isinstance(expect, dict) and not any("expect" in e for e in header_errs):
        for kind, names in expect.items():
            present = {s["name"] for s in spans.values() if s["kind"] == kind and s["name"] is not None}
            for name in names:
                if name == "*":
                    if not present:
                        err("[7] expected instrumentation missing: no %s spans" % kind)
                elif name not in present:
                    err("[7] expected instrumentation missing: %s %r" % (kind, name))

    # Rule 8: trace-on needs run.pass spans each containing calc.full and calc.recalc.
    rp_ids = sorted(pass_numbers)
    contents: Dict[int, set] = {i: set() for i in rp_ids}
    for sid, rp in run_pass_of.items():
        if rp and sid in spans:
            contents[rp].add(spans[sid]["kind"])
    missing_calc = []
    for rp in rp_ids:
        for need in ("calc.full", "calc.recalc"):
            if need not in contents[rp]:
                missing_calc.append("pass %r lacks %s" % (pass_numbers[rp], need))
    if mode == "trace-on":
        if not rp_ids:
            err("[8] trace-on run has no run.pass spans")
        for m in missing_calc:
            err("[8] " + m)
        if rp_ids and not set(by_kind) - TRACE_OFF_KINDS:
            warnings.append("trace-on run records only run.pass/calc.full/calc.recalc; it may be a trace-off run")
    elif mode == "trace-off":
        warnings.extend(missing_calc)
        extra_kinds = set(by_kind) - TRACE_OFF_KINDS
        if extra_kinds:
            warnings.append("trace-off run records kinds beyond run.pass/calc.full/calc.recalc: %s" % sorted(extra_kinds))

    if stats.get("source_kind") == "synthetic":
        warnings.append("synthetic trace: not a real-workbook measurement")

    stats.update(
        {
            "events": len(body),
            "by_type": by_type,
            "by_kind": by_kind,
            "by_clock": by_clock,
            "passes": sorted(p for p in pass_numbers.values() if _is_int(p)),
            "open_spans": len(stack),
        }
    )
    if len(errors) > _MAX_REPORTED_ERRORS:
        extra = len(errors) - _MAX_REPORTED_ERRORS
        errors = errors[:_MAX_REPORTED_ERRORS] + ["... %d more errors" % extra]
    return ValidationResult(ok=not errors, errors=errors, warnings=warnings, stats=stats)
