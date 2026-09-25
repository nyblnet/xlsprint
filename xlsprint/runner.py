"""Windows COM host orchestration for ``xlsprint profile``.

Drives an isolated Excel instance over COM (pywin32, imported lazily inside
``profile``), imports ``vba/XLSprintTimer.bas`` into a disposable copy of the
workbook, and runs matched trace-off / trace-on passes with
``XSP_RunPass``. See docs/DESIGN.md for the contract.

Host stages are timed with ``time.perf_counter_ns`` into memory while the run
happens and are written to ``trace-on.jsonl`` only after ``verify_output``,
together with the VBA events of each pass nested under the
``profile_trace_on`` stage that produced them. ``trace.validate`` then runs on
the written files, and its result, plus the ``render_report`` timing, goes
into ``run.json``.

Every blocking COM call runs under a watchdog: when a stage passes its
deadline (a modal dialog, a hung macro), the watchdog thread terminates the
Excel process, which makes the blocked call fail; that becomes
``RunnerError("timed out in <stage> ...")``. The watchdog makes no COM calls.

Everything above the COM layer (plan building, array extension, path checks,
event merging, redaction, argv reduction, the watchdog) is pure and tested
on any OS.
"""

from __future__ import annotations

import contextlib
import datetime as _dt
import gc
import hashlib
import json
import os
import platform
import re
import secrets
import shutil
import sys
import threading
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple

from xlsprint import __version__

RS = "\x1e"  # step separator (ChrW(30))
US = "\x1f"  # field separator (ChrW(31))

VBA_MODULE = "XLSprintTimer"
VBA_VERSION = "xlsprint-vba/1"
BAS_PATH = Path(__file__).with_name("vba") / (VBA_MODULE + ".bas")
PLAN_FILE = "plan.txt"

MAX_GROUP_AREAS = 32
CLOCK_PROBES = 7
QUIT_WAIT_S = 15.0
DRIFT_FAIL_FACTOR = 10
DRIFT_FAIL_MIN_NS = 1_000_000

# Spans whose failure makes the whole run fail (the workbook-level baseline).
CRITICAL_KINDS = frozenset({"run.pass", "calc.full", "calc.recalc"})

XL_CALC = {-4105: "automatic", -4135: "manual", 2: "semiautomatic"}
MSO_AUTOMATION_SECURITY_LOW = 1
VBEXT_CT_STD_MODULE = 1

_A1_CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}"
_A1_AREA_RE = re.compile(r"^(%s)(?::(%s))?$" % (_A1_CELL, _A1_CELL))
_CELL_PARTS_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")
_WHOLE_COLS_RE = re.compile(r"^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$")
_WHOLE_ROWS_RE = re.compile(r"^\$?([0-9]{1,7}):\$?([0-9]{1,7})$")
_STEP_KEY_RE = re.compile(r"^[RNG][0-9]{4,}$")
_MAX_COL = 16_384
_MAX_ROW = 1_048_576
_WORKBOOK_SUFFIXES = (".xlsx", ".xlsm", ".xlsb", ".xls")

Rect = Tuple[int, int, int, int]


class RunnerError(Exception):
    """A profiling run failed a check. The run is fail-closed."""


@dataclass
class ProfileOptions:
    workbook: Path
    out_dir: Path
    repeats: int = 5
    blocks: int = 4
    ranges: List[str] = field(default_factory=list)
    macros: List[str] = field(default_factory=list)
    names_timing: bool = True
    group_timing: bool = False
    full_rebuild: bool = False
    names: str = "hashed"
    keep_copy: bool = False
    label: Optional[str] = None
    max_events: int = 100_000
    # Extensions to the DESIGN.md signature (all optional).
    timeout_s: float = 900.0  # deadline per pass and per open/instrument/warmup/close stage
    enable_events: bool = False  # Application.EnableEvents in the profiled instance
    salt: Optional[str] = None  # redaction salt; random per run when None; never written out
    argv: Optional[List[str]] = None  # recorded in run.json, paths reduced, identifiers redacted
    semantics: Optional[Path] = None  # optional human-authored xlsprint.semantics/1 sidecar


# --------------------------------------------------------------------------
# Pure helpers: hashing, paths


def sha256_file(path, *, chunk: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


def verify_unchanged(path, expected_sha256: str) -> None:
    """Raise RunnerError if ``path`` no longer hashes to ``expected_sha256``."""
    try:
        actual = sha256_file(path)
    except OSError as exc:
        raise RunnerError(f"original workbook could not be re-read for verification: {exc}") from exc
    if actual != expected_sha256:
        raise RunnerError(
            "original workbook changed during the run (sha256 %s... -> %s...); results discarded"
            % (expected_sha256[:10], actual[:10])
        )


def _is_unc(path: str) -> bool:
    return path.replace("/", "\\").startswith("\\\\")


def _is_network_drive(path: Path) -> bool:
    """True for a mapped network drive on Windows; False elsewhere."""
    if sys.platform != "win32":
        return False
    try:
        import ctypes

        drive = os.path.splitdrive(str(path))[0]
        if not drive:
            return False
        DRIVE_REMOTE = 4
        return ctypes.windll.kernel32.GetDriveTypeW(drive + "\\") == DRIVE_REMOTE  # type: ignore[attr-defined]
    except Exception:
        return False


def check_paths(
    workbook, out_dir, *, is_network: Callable[[Path], bool] = _is_network_drive
) -> Tuple[Path, Path]:
    """Validate the workbook and output directory. Returns resolved (workbook, out_dir).

    Refuses UNC and network paths, a missing workbook, an output directory
    whose parent does not exist, and an output directory that is the
    workbook's own folder (nothing is ever written next to the original).
    """
    wb_raw, out_raw = str(workbook), str(out_dir)
    for label, raw in (("workbook", wb_raw), ("output directory", out_raw)):
        if _is_unc(raw):
            raise RunnerError(f"{label} is a UNC/network path; copy it to a local disk first")
    wb = Path(workbook).expanduser().resolve()
    out = Path(out_dir).expanduser().resolve()
    for label, p in (("workbook", wb), ("output directory", out)):
        if _is_unc(str(p)) or is_network(p):
            raise RunnerError(f"{label} is on a network location; use a local disk")
    if not wb.is_file():
        raise RunnerError(f"workbook not found: {wb}")
    if wb.suffix.lower() not in (".xlsx", ".xlsm"):
        # The plan comes from openpyxl inspection, which reads only OOXML workbooks.
        raise RunnerError(f"unsupported workbook type {wb.suffix!r} (expected .xlsx or .xlsm; save .xlsb/.xls as .xlsm first)")
    if not out.parent.is_dir():
        raise RunnerError(f"parent of the output directory does not exist: {out.parent}")
    if out.exists() and not out.is_dir():
        raise RunnerError(f"output path exists and is not a directory: {out}")
    if out == wb.parent:
        raise RunnerError("output directory must not be the workbook's own folder; choose a separate directory")
    return wb, out


# --------------------------------------------------------------------------
# Pure helpers: redaction


def make_ident(names_mode: str, salt: str, *, check_name: Optional[Callable[[str], Optional[str]]] = None,
               redact: Optional[Callable[[str, str], str]] = None) -> Callable[[str], str]:
    """Identifier transform for everything the run persists.

    ``hashed``: every identifier becomes ``h:<10 hex>``. ``clear``: an
    identifier stays clear unless it fails ``trace.check_name`` (DESIGN rule
    10), in which case it is hashed too; the span is never dropped.
    """
    if check_name is None or redact is None:
        from xlsprint import trace

        check_name = check_name or trace.check_name
        redact = redact or trace.redact_name

    def ident(name: str) -> str:
        if names_mode == "hashed" or check_name(name) is not None:
            return redact(name, salt)
        return name

    return ident


def _range_spec_ident(spec: str, ident: Callable[[str], str]) -> str:
    try:
        sheet, addr = parse_user_range(spec)
    except RunnerError:
        return ident(spec)
    return ident(sheet) + "!" + addr


def reduce_argv(argv: Sequence[str], *, ident: Optional[Callable[[str], str]] = None,
                hash_files: bool = False) -> List[str]:
    """argv for run.json.

    Path-like values become basenames (``--out=C:\\x\\y`` -> ``--out=y``);
    with ``hash_files`` a workbook basename is redacted too (extension kept).
    With ``ident``, the values of ``--range`` and ``--macro`` are redacted.
    """

    def base(v: str) -> str:
        b = v
        if "/" in v or "\\" in v:
            b = re.split(r"[\\/]", v.rstrip("\\/"))[-1] or v
        if hash_files and ident is not None and b.lower().endswith(_WORKBOOK_SUFFIXES):
            stem, dot, ext = b.rpartition(".")
            b = ident(stem) + dot + ext
        return b

    def value(opt: str, v: str) -> str:
        if ident is not None and opt == "--range":
            return _range_spec_ident(v, ident)
        if ident is not None and opt == "--macro":
            return ident(v)
        return base(v)

    out: List[str] = []
    pending: Optional[str] = None
    for i, a in enumerate(argv):
        if i == 0:
            out.append(re.split(r"[\\/]", a.rstrip("\\/"))[-1] or a)
        elif pending is not None:
            out.append(value(pending, a))
            pending = None
        elif a.startswith("-") and "=" in a:
            k, v = a.split("=", 1)
            out.append(k + "=" + value(k, v))
        elif a in ("--range", "--macro"):
            out.append(a)
            pending = a
        else:
            out.append(base(a))
    return out


_GENERIC_PATH_RES = [
    re.compile(r"\\\\[^\s'\"]+"),  # UNC
    re.compile(r"\b[A-Za-z]:[\\/][^'\"\r\n]*?(?=['\"]|\s-|\s\(|$|\)\s|\)$)"),  # drive paths (may contain spaces)
    re.compile(r"(?<![\w.])/(?:[^\s'\"/]+/)+[^\s'\"]*"),  # POSIX absolute
]


def scrub_text(text: str, *, paths: Iterable[str] = (), idents: Iterable[str] = (),
               ident: Optional[Callable[[str], str]] = None) -> str:
    """Remove paths and (with ``ident``) identifiers from text that will be persisted."""
    out = text
    for p in sorted({str(p) for p in paths if p}, key=len, reverse=True):
        for variant in {p, p.replace("\\", "/"), p.replace("/", "\\")}:
            out = out.replace(variant, "<path>")
    for rx in _GENERIC_PATH_RES:
        out = rx.sub("<path>", out)
    if ident is not None:
        for name in sorted({n for n in idents if n and len(n) >= 3}, key=len, reverse=True):
            if name in out:
                out = out.replace(name, ident(name))
    return out


# --------------------------------------------------------------------------
# Pure helpers: A1 areas


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _col_letters(idx: int) -> str:
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_area(address: str) -> Rect:
    """``"A1:F20"`` / ``"$B$2"`` -> (row1, col1, row2, col2), normalized so 1 <= r1 <= r2."""
    m = _A1_AREA_RE.match(address.strip())
    if not m:
        raise RunnerError(f"not a single A1 area: {address!r}")
    a = _CELL_PARTS_RE.match(m.group(1))
    b = _CELL_PARTS_RE.match(m.group(2) or m.group(1))
    c1, r1 = _col_index(a.group(1)), int(a.group(2))
    c2, r2 = _col_index(b.group(1)), int(b.group(2))
    r1, r2 = sorted((r1, r2))
    c1, c2 = sorted((c1, c2))
    if not (1 <= r1 and r2 <= _MAX_ROW and 1 <= c1 and c2 <= _MAX_COL):
        raise RunnerError(f"address outside the worksheet grid: {address!r}")
    return r1, c1, r2, c2


def area_a1(r1: int, c1: int, r2: int, c2: int) -> str:
    first = f"${_col_letters(c1)}${r1}"
    if (r1, c1) == (r2, c2):
        return first
    return f"{first}:${_col_letters(c2)}${r2}"


def split_column_blocks(address: str, blocks: int) -> List[str]:
    """Split an area into at most ``blocks`` contiguous column blocks, as even as possible."""
    if blocks < 1:
        return []
    r1, c1, r2, c2 = parse_area(address)
    ncols = c2 - c1 + 1
    k = min(blocks, ncols)
    base, extra = divmod(ncols, k)
    out, start = [], c1
    for i in range(k):
        width = base + (1 if i < extra else 0)
        out.append(area_a1(r1, start, r2, start + width - 1))
        start += width
    return out


def parse_user_range(spec: str) -> Tuple[str, str]:
    """``Sheet1!A1:D10`` or ``'My Sheet'!A1:D10`` -> (sheet, "$A$1:$D$10")."""
    if "!" not in spec:
        raise RunnerError("--range needs SHEET!A1:B2")
    sheet, addr = spec.rsplit("!", 1)
    sheet = sheet.strip()
    if len(sheet) >= 2 and sheet[0] == sheet[-1] == "'":
        sheet = sheet[1:-1].replace("''", "'")
    if not sheet:
        raise RunnerError("--range has an empty sheet name")
    return sheet, area_a1(*parse_area(addr))


def clip_to_used(address: str, used_range: Optional[str]) -> Optional[str]:
    """Absolute A1 area for ``address``; whole columns/rows are intersected with ``used_range``.

    Returns None when a whole-column/row address has no used range or does
    not intersect it (nothing there to calculate).
    """
    a = address.strip()
    cols, rows = _WHOLE_COLS_RE.match(a), _WHOLE_ROWS_RE.match(a)
    if not cols and not rows:
        return area_a1(*parse_area(a))
    if not used_range:
        return None
    ur1, uc1, ur2, uc2 = parse_area(used_range)
    if cols:
        c1, c2 = sorted((_col_index(cols.group(1)), _col_index(cols.group(2))))
        r1, r2 = ur1, ur2
        c1, c2 = max(c1, uc1), min(c2, uc2)
    else:
        r1, r2 = sorted((int(rows.group(1)), int(rows.group(2))))
        c1, c2 = uc1, uc2
        r1, r2 = max(r1, ur1), min(r2, ur2)
    if r1 > r2 or c1 > c2:
        return None
    return area_a1(r1, c1, r2, c2)


def _intersects(a: Rect, b: Rect) -> bool:
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


def _subtract(a: Rect, b: Rect) -> List[Rect]:
    """a minus b as disjoint rectangles."""
    if not _intersects(a, b):
        return [a]
    r1, c1, r2, c2 = a
    out = []
    if b[0] > r1:
        out.append((r1, c1, b[0] - 1, c2))
    if b[2] < r2:
        out.append((b[2] + 1, c1, r2, c2))
    mr1, mr2 = max(r1, b[0]), min(r2, b[2])
    if b[1] > c1:
        out.append((mr1, c1, mr2, b[1] - 1))
    if b[3] < c2:
        out.append((mr1, b[3] + 1, mr2, c2))
    return out


def _add_disjoint(result: List[Rect], rect: Rect) -> None:
    pieces = [rect]
    for q in result:
        pieces = [p for x in pieces for p in _subtract(x, q)]
    result.extend(pieces)


def extend_with_arrays(areas: Sequence[str], array_areas: Sequence[str]) -> List[str]:
    """Areas plus every array-formula area that intersects them, as disjoint A1 areas.

    Microsoft's RangeTimer extends a range to whole array formulas
    (HasArray -> CurrentArray); this does the same from static inspection so
    Excel never calculates a cell twice and no per-cell scan runs in VBA.
    """
    base = [parse_area(a) for a in areas]
    result: List[Rect] = []
    for r in base:
        _add_disjoint(result, r)
    for arr_a1 in array_areas:
        arr = parse_area(arr_a1)
        if any(_intersects(arr, r) for r in base):
            _add_disjoint(result, arr)
    return [area_a1(*r) for r in result]


# --------------------------------------------------------------------------
# Pure helpers: plan building


def adapt_plan_info(info: dict) -> dict:
    """Normalize ``formulas.inspect_workbook_and_plan`` plan output (the only place its shape is read).

    Returns {"sheets": [{"sheet", "formula_cells", "formula_area", "array_areas"}],
             "names": [{"name", "scope", "sheet", "address"}],     # only names over formula cells
             "groups": [{"group", "sheet", "areas", "areas_total"}],
             "skipped_names": {reason: count}}.
    Whole-column/row name addresses are clipped to the sheet's used range.
    """
    sheets, used = [], {}
    for s in info.get("sheets") or []:
        used[s["sheet"]] = s.get("used_range")
        sheets.append({
            "sheet": s["sheet"],
            "formula_cells": int(s.get("formula_cells") or 0),
            "formula_area": s.get("formula_area"),
            "array_areas": list(s.get("array_areas") or []),
        })
    names, skipped = [], {}

    def skip(reason):
        skipped[reason] = skipped.get(reason, 0) + 1

    for n in info.get("names") or []:
        if n.get("is_range") is False:
            skip("not_a_range")
            continue
        if not n.get("has_formulas", int(n.get("formula_cells") or 0) > 0) or n.get("formula_cells") == 0:
            skip("no_formulas")
            continue
        sheet, address = n.get("sheet"), n.get("address")
        if (sheet is None or address is None) and n.get("refers_to_range"):
            sheet, address = n["refers_to_range"].rsplit("!", 1)
            sheet = sheet[1:-1].replace("''", "'") if sheet[:1] == sheet[-1:] == "'" else sheet
        if not sheet or not address:
            skip("unresolved")
            continue
        try:
            clipped = clip_to_used(address, used.get(sheet))
        except RunnerError:
            skip("unsupported_address")
            continue
        if clipped is None:
            skip("outside_used_range")
            continue
        names.append({"name": n["name"], "scope": n.get("scope") or "workbook", "sheet": sheet, "address": clipped})
    groups = []
    for g in info.get("groups") or []:
        areas = list(g.get("areas") or [])
        groups.append({
            "group": g["group"],
            "sheet": g["sheet"],
            "areas": areas,
            "areas_total": int(g.get("areas_total", len(areas))),
        })
    return {"sheets": sheets, "names": names, "groups": groups, "skipped_names": skipped}


def build_plan(plan_info: dict, opts: ProfileOptions) -> List[Tuple[str, ...]]:
    """Logical drill-down plan in Microsoft's order (DESIGN.md "Drill-down plan").

    Steps: ("fullrebuild",) ("full",) ("recalc",) ("sheet", sheet)
    ("range", sheet, addr, label) ("name", sheet, addr, name, scope)
    ("group", sheet, "a1,a2,...", gid) ("macro", name).
    ``compile_plan`` turns them into what the VBA receives.
    """
    steps: List[Tuple[str, ...]] = []
    if opts.full_rebuild:
        steps.append(("fullrebuild",))
    steps.append(("full",))
    steps.append(("recalc",))
    formula_sheets = [s for s in plan_info["sheets"] if s["formula_cells"] > 0]
    for s in formula_sheets:
        steps.append(("sheet", s["sheet"]))
    for spec in opts.ranges:
        sheet, addr = parse_user_range(spec)
        steps.append(("range", sheet, addr, "user"))
    for s in formula_sheets:
        if not s["formula_area"]:
            continue
        blocks = split_column_blocks(s["formula_area"], opts.blocks)
        for i, addr in enumerate(blocks, 1):
            steps.append(("range", s["sheet"], addr, f"auto {i}/{len(blocks)}"))
    if opts.names_timing:
        for n in plan_info["names"]:
            steps.append(("name", n["sheet"], n["address"], n["name"], n.get("scope") or "workbook"))
    if opts.group_timing:
        for g in plan_info["groups"]:
            if not g["areas"] or g["areas_total"] > MAX_GROUP_AREAS or len(g["areas"]) != g["areas_total"]:
                continue
            steps.append(("group", g["sheet"], ",".join(g["areas"]), g["group"]))
    for m in opts.macros:
        steps.append(("macro", m))
    return steps


def compile_plan(steps: Sequence[Sequence[str]], plan_info: dict,
                 ident: Callable[[str], str]) -> Tuple[List[Tuple[str, ...]], Dict[str, dict]]:
    """Logical steps -> (VBA steps, meta).

    Range, name and group steps get their addresses extended to whole array
    formulas and a short key (R0001 / N0001 / the group id) that the VBA uses
    as span name. ``meta[key]`` holds the trace name (already redacted), the
    clear sheet (for the ``sheet`` attribute), and the tool-generated note or
    group id.
    """
    arrays = {s["sheet"]: s.get("array_areas") or [] for s in plan_info.get("sheets", [])}
    vba: List[Tuple[str, ...]] = []
    meta: Dict[str, dict] = {}
    nr = nn = 0
    for st in steps:
        kind = st[0]
        if kind == "range":
            nr += 1
            key = "R%04d" % nr
            sheet, addr, label = st[1], st[2], st[3]
            meta[key] = {"name": ident(sheet) + "!" + addr, "sheet": sheet, "note": label}
        elif kind == "name":
            nn += 1
            key = "N%04d" % nn
            sheet, addr, nm, scope = st[1], st[2], st[3], st[4]
            trace_name = ident(nm) if scope == "workbook" else ident(scope) + "!" + ident(nm)
            meta[key] = {"name": trace_name, "sheet": sheet}
        elif kind == "group":
            key = st[3]
            if not _STEP_KEY_RE.match(key) or key in meta:
                raise RunnerError(f"bad or duplicate group id {key!r}")
            sheet, addr = st[1], st[2]
            meta[key] = {"name": key, "sheet": sheet, "group": key}
        else:
            vba.append(tuple(st))
            continue
        full = extend_with_arrays(addr.split(","), arrays.get(sheet, []))
        vba.append((kind, sheet, ",".join(full), key))
    return vba, meta


def encode_plan(steps: Iterable[Sequence[str]]) -> str:
    """Encode VBA steps as fields joined by ChrW(31) and steps joined by ChrW(30)."""
    arity = {"full": 1, "fullrebuild": 1, "recalc": 1, "sheet": 2, "range": 4, "name": 4, "group": 4, "macro": 2}
    out = []
    for st in steps:
        st = tuple(st)
        if not st or st[0] not in arity or len(st) != arity[st[0]]:
            raise RunnerError(f"malformed plan step {st[:1]!r} with {len(st)} fields")
        for f in st:
            if not isinstance(f, str) or not f or RS in f or US in f:
                raise RunnerError(f"plan step {st[0]!r} has an empty or separator-containing field")
        if st[0] in ("range", "name", "group") and not _STEP_KEY_RE.match(st[3]):
            raise RunnerError(f"plan step {st[0]!r} needs a key like R0001")
        out.append(US.join(st))
    return RS.join(out)


def write_plan_file(path: Path, encoded: str) -> Path:
    """UTF-16LE with BOM, as XSP_RunPass reads it (non-ASCII sheet names survive)."""
    path.write_bytes(b"\xff\xfe" + encoded.encode("utf-16-le"))
    return path


# --------------------------------------------------------------------------
# Pure helpers: VBA events -> trace


def check_vba_footer(footer: dict) -> List[str]:
    """Fail-closed checks on a vba_footer (as returned by trace.read_vba_events)."""
    probs = []
    if footer.get("truncated"):
        probs.append("VBA buffer truncated")
    if footer.get("dropped"):
        probs.append(f"VBA dropped {footer['dropped']} events")
    if footer.get("open_spans"):
        probs.append(f"{footer['open_spans']} VBA spans left open")
    if footer.get("faults"):
        probs.append(f"{footer['faults']} VBA instrumentation faults (first: {footer.get('first_fault') or '?'})")
    return probs


def merge_vba_events(
    writer,
    events: Sequence[dict],
    *,
    ident: Callable[[str], str],
    meta: Optional[Dict[str, dict]] = None,
    attr_ok: Optional[Callable[[str], bool]] = None,
) -> int:
    """Append one pass of VBA events under the writer's innermost open span.

    VBA ids are remapped into a block from ``writer.reserve_ids``; parents are
    remapped the same way (root parent 0 is rebased by ``append_event`` onto
    the open host stage) and depth is recomputed by the writer. Every
    identifier goes through ``ident``; range/name/group keys are replaced by
    their trace names from ``meta``. ``sheet``, ``note`` and ``group``
    attributes are added from ``meta`` (``sheet`` only if ``attr_ok``).
    Returns the number of events appended.
    """
    if not events:
        return 0
    ids = [e["id"] for e in events if e.get("type") in ("B", "M")]
    lo, hi = min(ids), max(ids)
    first = writer.reserve_ids(hi - lo + 1)
    meta = meta or {}
    n = 0
    for ev in events:
        ev = dict(ev)
        ev["id"] = first + (ev["id"] - lo)
        if ev.get("type") in ("B", "M"):
            ev["parent"] = first + (ev["parent"] - lo) if ev.get("parent") else 0
            ev.pop("depth", None)
            kind, name = ev.get("kind"), ev.get("name", "")
            attrs = dict(ev.get("attrs") or {})
            sheet = None
            if kind == "calc.sheet":
                sheet = name
                ev["name"] = ident(name)
            elif kind in ("calc.range", "calc.name", "calc.group"):
                m = meta.get(name)
                if m is None:
                    ev["name"] = ident(name)
                else:
                    ev["name"] = m["name"]
                    sheet = m.get("sheet")
                    for k in ("note", "group"):
                        if m.get(k):
                            attrs[k] = m[k]
            elif kind in ("run.pass", "calc.full", "calc.recalc", "calc.fullrebuild"):
                pass  # fixed names written by the driver ("pass", "workbook")
            else:  # vba.proc, marker: user-chosen identifiers
                ev["name"] = ident(name)
            if sheet is not None and "sheet" not in attrs:
                value = ident(sheet)
                if attr_ok is None or attr_ok(value):
                    attrs["sheet"] = value
            ev["attrs"] = attrs
        writer.append_event(ev)
        n += 1
    return n


def check_expected(
    passes: Sequence[Tuple[int, str, Sequence[dict]]],
    plan: Sequence[Sequence[str]],
) -> List[str]:
    """Missing-instrumentation check on raw (clear-name) VBA events.

    Every pass needs run.pass, calc.full and calc.recalc. Every trace-on pass
    also needs a calc.sheet for each planned sheet and a vba.proc for each
    requested macro. Trace-off passes must record nothing else. Messages
    never contain identifiers.
    """
    probs = []
    want_sheets = [s[1] for s in plan if s[0] == "sheet"]
    want_macros = [s[1] for s in plan if s[0] == "macro"]
    off_kinds = {"run.pass", "calc.full", "calc.recalc"}
    for pass_id, mode, events in passes:
        begins = [e for e in events if e.get("type") == "B"]
        kinds = {e["kind"] for e in begins}
        for k in sorted(off_kinds - kinds):
            probs.append(f"{mode} pass {pass_id}: no {k} span")
        if mode == "off":
            extra = kinds - off_kinds
            if extra:
                probs.append(f"off pass {pass_id}: unexpected kinds {sorted(extra)}")
            continue
        sheets = {e["name"] for e in begins if e["kind"] == "calc.sheet"}
        procs = {e["name"] for e in begins if e["kind"] == "vba.proc"}
        missing = [i for i, s in enumerate(want_sheets, 1) if s not in sheets]
        if missing:
            probs.append(f"on pass {pass_id}: {len(missing)} planned sheet(s) have no calc.sheet span")
        for i, m in enumerate(want_macros, 1):
            if m not in procs:
                probs.append(f"on pass {pass_id}: requested macro #{i} has no vba.proc span")
    return probs


def span_failures(passes: Sequence[Tuple[int, str, Sequence[dict]]]) -> Tuple[List[dict], List[dict]]:
    """(critical, other) spans that did not end "ok". Records carry no identifiers."""
    critical, other = [], []
    for pass_id, mode, events in passes:
        kinds = {e["id"]: e.get("kind") for e in events if e.get("type") == "B"}
        for e in events:
            if e.get("type") == "E" and e.get("status") != "ok":
                kind = kinds.get(e["id"], "?")
                rec = {"pass": pass_id, "mode": mode, "kind": kind, "status": e.get("status"),
                       "error_code": (e.get("attrs") or {}).get("error_code")}
                (critical if kind in CRITICAL_KINDS else other).append(rec)
    return critical, other


def normalized_bas(dest: Path) -> Path:
    """Copy the .bas with CRLF line endings (the VBE importer expects them)."""
    text = BAS_PATH.read_text(encoding="ascii")
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    dest.write_bytes(text.encode("ascii"))
    return dest


def choose_probe(probes: Sequence[Tuple[int, int, int]]) -> Tuple[int, int]:
    """(h0, vba_ns, h1) samples -> (offset_ns, uncertainty_ns) of the tightest one."""
    if not probes:
        raise RunnerError("no clock probes")
    h0, v, h1 = min(probes, key=lambda p: p[2] - p[0])
    if h1 < h0:
        raise RunnerError("host clock went backwards during clock probe")
    return v - (h0 + h1) // 2, (h1 - h0 + 1) // 2


def check_drift(drift_ns: int, uncertainty_ns: int) -> str:
    """"ok", "warn" (|drift| > uncertainty) or "fail" (> 10x uncertainty and > 1 ms)."""
    d = abs(drift_ns)
    if d > DRIFT_FAIL_FACTOR * uncertainty_ns and d > DRIFT_FAIL_MIN_NS:
        return "fail"
    if d > uncertainty_ns:
        return "warn"
    return "ok"


# --------------------------------------------------------------------------
# Host stage recorder and watchdog


class StageRecorder:
    """Times host stages into memory; they are written to the trace later."""

    def __init__(self) -> None:
        self.stages: List[dict] = []

    @contextlib.contextmanager
    def stage(self, name: str, **attrs: Any) -> Iterator[dict]:
        rec = {"name": name, "start": time.perf_counter_ns(), "end": None, "status": "ok", "attrs": dict(attrs)}
        self.stages.append(rec)
        try:
            yield rec
        except BaseException:
            rec["status"] = "error"
            raise
        finally:
            rec["end"] = time.perf_counter_ns()

    def summary(self) -> List[dict]:
        return [
            {"name": s["name"], "dur_ns": (s["end"] or s["start"]) - s["start"], "status": s["status"]}
            for s in self.stages
        ]


class Watchdog:
    """Deadline thread. On expiry it calls ``kill`` (never a COM call) once.

    ``guard(stage, seconds)`` arms the deadline around a blocking call. If
    the deadline fired, whatever the guarded call raised (or its return) is
    turned into ``RunnerError("timed out in <stage> ...")``. Guards do not nest.
    """

    def __init__(self) -> None:
        self._cond = threading.Condition()
        self._deadline: Optional[float] = None
        self._stage: Optional[str] = None
        self._kill: Optional[Callable[[], None]] = None
        self._stop = False
        self.fired: Optional[str] = None
        self._thread = threading.Thread(target=self._run, name="xlsprint-watchdog", daemon=True)
        self._thread.start()

    def set_kill(self, kill: Optional[Callable[[], None]]) -> None:
        with self._cond:
            self._kill = kill

    @contextlib.contextmanager
    def guard(self, stage: str, seconds: float) -> Iterator[None]:
        with self._cond:
            self.fired = None
            self._stage = stage
            self._deadline = time.monotonic() + float(seconds)
            self._cond.notify_all()
        try:
            yield
        except Exception as exc:
            if self.fired == stage:
                raise RunnerError(self._message(stage, seconds)) from exc
            raise
        finally:
            with self._cond:
                self._deadline = None
                self._stage = None
        if self.fired == stage:
            raise RunnerError(self._message(stage, seconds))

    @staticmethod
    def _message(stage: str, seconds: float) -> str:
        return (f"timed out in {stage} after {seconds:g}s; Excel was terminated "
                "(a modal dialog, a hung macro or a very long calculation?)")

    def _run(self) -> None:
        with self._cond:
            while not self._stop:
                if self._deadline is None:
                    self._cond.wait()
                    continue
                remaining = self._deadline - time.monotonic()
                if remaining > 0:
                    self._cond.wait(remaining)
                    continue
                self.fired = self._stage
                self._deadline = None
                if self._kill is not None:
                    try:
                        self._kill()
                    except Exception:
                        pass

    def stop(self) -> None:
        with self._cond:
            self._stop = True
            self._cond.notify_all()
        self._thread.join(timeout=5)


def write_trace(path, header: dict, stages: Sequence[dict], *, max_events: int, include, merge) -> dict:
    """Write host stages (those ``include(stage)`` selects) with VBA events merged in.

    ``merge(writer, stage)`` appends the VBA events belonging to ``stage``.
    Returns the footer.
    """
    from xlsprint import trace

    w = trace.TraceWriter(path, header, max_events=max_events)
    try:
        for st in stages:
            if not include(st):
                continue
            sid = w.begin("host.stage", st["name"], clock="host", pass_=0, attrs=st["attrs"] or None, ns=st["start"])
            merge(w, st)
            w.end(sid, status=st["status"], ns=st["end"])
    finally:
        footer = w.close()
    return footer


def attr_value_ok(value: str) -> bool:
    """True if ``value`` may be stored as a string attribute (the contract's pattern, via trace)."""
    from xlsprint import trace

    return trace.check_attr("sheet", value) is None


# --------------------------------------------------------------------------
# COM layer (Windows only)


def _qual(wb_name: str, proc: str) -> str:
    return "'%s'!%s.%s" % (wb_name.replace("'", "''"), VBA_MODULE, proc)


_RETRYABLE_EXCEL_BUSY_HRESULTS = frozenset({0x80010001, 0x8001010A})
_MAX_EXCEL_BUSY_RETRIES = 120
_EXCEL_BUSY_RETRY_MAX_DELAY_S = 0.25


def _is_excel_busy_rejection(exc: BaseException) -> bool:
    """Return whether COM reports Excel temporarily rejecting an automation call."""

    hresult = getattr(exc, "hresult", None)
    if hresult is None:
        args = getattr(exc, "args", ())
        hresult = args[0] if args else None
    try:
        code = int(hresult) & 0xFFFFFFFF
    except (TypeError, ValueError, OverflowError):
        return False
    return code in _RETRYABLE_EXCEL_BUSY_HRESULTS


def _com_err(exc: BaseException) -> str:
    return f"{type(exc).__name__}: {exc}"


def profile(opts: ProfileOptions) -> dict:
    """Run a profiling session and return the run.json content. Raises RunnerError on any failure."""
    if sys.platform != "win32":
        raise RunnerError(
            "profile needs Windows with Excel Desktop 2013+ (it drives Excel over COM). "
            "Offline commands (inspect, report, validate, selftest) work on any OS."
        )
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RunnerError("pywin32 is required on Windows: pip install 'xlsprint[windows]'") from exc

    from xlsprint import formulas, report, trace

    if opts.repeats < 1:
        raise RunnerError("repeats must be >= 1")
    if opts.timeout_s <= 0:
        raise RunnerError("timeout must be positive")
    if opts.names not in ("hashed", "clear"):
        raise RunnerError("names must be 'hashed' or 'clear'")
    for m in opts.macros:
        if not re.match(r"^[A-Za-z][A-Za-z0-9_]*(\.[A-Za-z][A-Za-z0-9_]*)?$", m):
            raise RunnerError("--macro must be a VBA procedure name (Module.Proc or Proc)")

    wb_path, out_dir = check_paths(opts.workbook, opts.out_dir)
    salt = opts.salt or secrets.token_hex(8)
    ident = make_ident(opts.names, salt)
    hashed = opts.names == "hashed"

    run_id = uuid.uuid4().hex
    out_dir.mkdir(parents=True, exist_ok=True)
    work = out_dir / "work" / run_id
    rec = StageRecorder()
    run: Dict[str, Any] = {
        "schema": "xlsprint.run/1",
        "ok": False,
        "run_id": run_id,
        "argv": reduce_argv(opts.argv if opts.argv is not None else sys.argv, ident=ident, hash_files=hashed),
        "tool_version": __version__,
        "python": platform.python_version(),
        "os": platform.platform(),
        "options": {
            "repeats": opts.repeats, "blocks": opts.blocks, "ranges": len(opts.ranges),
            "macros": len(opts.macros), "names_timing": opts.names_timing,
            "group_timing": opts.group_timing, "full_rebuild": opts.full_rebuild,
            "names": opts.names, "keep_copy": opts.keep_copy, "max_events": opts.max_events,
            "enable_events": opts.enable_events, "timeout_s": opts.timeout_s,
        },
        "warnings": [],
        "step_errors": [],
        "error": None,
    }
    warnings: List[str] = run["warnings"]
    # Everything persisted in run.json is scrubbed of these.
    secret_paths = [str(wb_path), str(wb_path.parent), str(out_dir), str(work)]
    secret_idents = [wb_path.name, wb_path.stem] + [parse_user_range(r)[0] for r in opts.ranges] + list(opts.macros)

    wd = Watchdog()
    xl = wb = None
    excel_pid: Optional[int] = None
    co_init = False
    orig_sha = None
    pass_results: List[dict] = []
    pass_events: Dict[Tuple[int, str], List[dict]] = {}
    passes_for_check: List[Tuple[int, str, List[dict]]] = []
    probes_before: List[Tuple[int, int, int]] = []
    probes_after: List[Tuple[int, int, int]] = []
    calc_original = "automatic"
    excel_info: Dict[str, Any] = {}
    meta: Dict[str, dict] = {}

    def run_vba(proc: str, *args):
        qualified = _qual(wb.Name, proc)
        for attempt in range(_MAX_EXCEL_BUSY_RETRIES + 1):
            try:
                return xl.Run(qualified, *args)
            except Exception as exc:
                if not _is_excel_busy_rejection(exc) or attempt >= _MAX_EXCEL_BUSY_RETRIES:
                    raise
                with contextlib.suppress(Exception):
                    pythoncom.PumpWaitingMessages()
                delay = min(0.05 * (attempt + 1), _EXCEL_BUSY_RETRY_MAX_DELAY_S)
                time.sleep(delay)

    def probe_clock(into: list) -> None:
        for _ in range(CLOCK_PROBES):
            h0 = time.perf_counter_ns()
            v = run_vba("XSP_ClockProbe")
            h1 = time.perf_counter_ns()
            into.append((h0, int(str(v).strip()), h1))

    def release_excel() -> List[str]:
        """Close without saving and Quit, then drop every COM reference.

        Settings are not restored: the instance is private and the copy is
        discarded, so switching back to automatic would only force a recalc.
        """
        nonlocal xl, wb
        errs = []
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception as exc:
                errs.append(f"close workbook: {_com_err(exc)}")
        if xl is not None:
            try:
                xl.Quit()
            except Exception as exc:
                errs.append(f"quit Excel: {_com_err(exc)}")
        wb = None
        xl = None
        gc.collect()
        return errs

    def ensure_exited() -> None:
        if excel_pid and not _wait_pid_exit(excel_pid, QUIT_WAIT_S):
            _terminate_pid(excel_pid)
            warnings.append("Excel did not exit after Quit; its process was terminated")

    try:
        pythoncom.CoInitialize()
        co_init = True

        # -- prepare_copy ---------------------------------------------------
        with rec.stage("prepare_copy") as st:
            orig_sha = sha256_file(wb_path)
            work.mkdir(parents=True, exist_ok=False)
            copy_path = work / wb_path.name
            secret_paths.append(str(copy_path))
            # Byte copy; alternate data streams (mark of the web) are not copied.
            shutil.copyfile(wb_path, copy_path)
            if sha256_file(copy_path) != orig_sha:
                raise RunnerError("work copy does not match the original")
            st["attrs"] = {"sha256_prefix": orig_sha[:10], "bytes": copy_path.stat().st_size}
            try:
                formulas_obj, plan_raw = formulas.inspect_workbook_and_plan(copy_path, names=opts.names, salt=salt)
            except Exception as exc:
                raise RunnerError(f"static inspection of the copy failed: {_com_err(exc)}") from exc
            if opts.semantics is not None:
                from xlsprint import semantics
                try:
                    manifest = semantics.load_manifest(opts.semantics)
                    formulas_obj["semantics"] = semantics.bind_manifest(
                        manifest, workbook_sha256=orig_sha, formulas=formulas_obj,
                        plan=plan_raw, names=opts.names, salt=salt)
                except (OSError, ValueError) as exc:
                    raise RunnerError(f"semantic map could not be bound: {_com_err(exc)}") from exc
            plan_info = adapt_plan_info(plan_raw)
            secret_idents += [s["sheet"] for s in plan_info["sheets"]] + [n["name"] for n in plan_info["names"]]
            plan = build_plan(plan_info, opts)
            vba_plan, meta = compile_plan(plan, plan_info, ident)
            plan_path = write_plan_file(work / PLAN_FILE, encode_plan(vba_plan))
            formulas.write_formulas_json(formulas_obj, out_dir / "formulas.json")
            bas = normalized_bas(work / (VBA_MODULE + ".bas"))
        run["plan"] = {"steps": len(plan), "by_kind": _count_kinds(plan),
                       "skipped_names": plan_info.get("skipped_names", {})}

        # -- launch_excel ---------------------------------------------------
        with rec.stage("launch_excel"):
            try:
                xl = win32com.client.DispatchEx("Excel.Application")
            except Exception as exc:
                raise RunnerError(f"could not start Excel over COM: {_com_err(exc)}") from exc
            excel_pid = _excel_pid(xl)
            if excel_pid is None:
                warnings.append("Excel process id unknown; timeouts cannot terminate it")
            else:
                pid = excel_pid
                wd.set_kill(lambda: _terminate_pid(pid))
            with wd.guard("launch_excel", opts.timeout_s):
                xl.Visible = False
                xl.DisplayAlerts = False
                xl.ScreenUpdating = False
                xl.EnableEvents = bool(opts.enable_events)
                xl.AskToUpdateLinks = False
                # Macros in files opened by this instance must run (the imported
                # module and --macro). The instance is private to this run.
                xl.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW

        # -- open_workbook --------------------------------------------------
        with rec.stage("open_workbook"), wd.guard("open_workbook", opts.timeout_s):
            try:
                wb = xl.Workbooks.Open(
                    str(copy_path), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True,
                    Notify=False, AddToMru=False,
                )
            except Exception as exc:
                if wd.fired:
                    raise
                raise RunnerError(f"Excel could not open the copy (password-protected or corrupt?): {_com_err(exc)}") from exc
            calc_original = XL_CALC.get(int(xl.Calculation), "automatic")
            excel_info = {
                "version": str(xl.Version),
                "build": str(xl.Build),
                "operating_system": str(xl.OperatingSystem),
            }
            mtc = xl.MultiThreadedCalculation
            excel_info["threads"] = int(mtc.ThreadCount)
            excel_info["multithreaded"] = bool(mtc.Enabled)
            del mtc  # plain values only; no COM reference may outlive the instance

        # -- instrument -----------------------------------------------------
        with rec.stage("instrument") as st, wd.guard("instrument", opts.timeout_s):
            _import_module(wb, bas)
            try:
                ver = run_vba("XSP_Version")
            except Exception as exc:
                if wd.fired:
                    raise
                raise RunnerError(
                    "the imported XLSprint module could not run. The workbook's VBA project may not "
                    f"compile, or macros are blocked by policy: {_com_err(exc)}"
                ) from exc
            if ver != VBA_VERSION:
                raise RunnerError(f"unexpected VBA module version {ver!r}")
            excel_info["bitness"] = int(run_vba("XSP_Bitness"))
            saved = run_vba("XSP_SaveSettings")
            if saved != calc_original:
                warnings.append(f"calculation mode read as {calc_original} over COM but {saved} in VBA")
            st["attrs"] = {"method": "VBComponents.Import"}

        # -- warmup (also probes the clock) --------------------------------
        with rec.stage("warmup", method="Application.CalculateFull"), wd.guard("warmup", opts.timeout_s):
            xl.CalculateFull()
            probe_clock(probes_before)
        offset_ns, uncert_ns = choose_probe(probes_before)

        # -- profile passes: off, on, off, on, ... -------------------------
        for k in range(1, opts.repeats + 1):
            for mode in ("off", "on"):
                path = work / f"vba-{mode}-{k:03d}.jsonl"
                stage = f"profile_trace_{mode}"
                with rec.stage(stage, repeat=k) as st, wd.guard(f"{stage} pass {k}", opts.timeout_s):
                    st["pass"] = k
                    st["mode"] = mode
                    r = run_vba("XSP_Init", str(path), mode, 1, 0, 0, int(opts.max_events))
                    if r != "ok":
                        raise RunnerError(f"XSP_Init failed ({r}) for {mode} pass {k}")
                    r = run_vba("XSP_RunPass", k, mode, str(plan_path))
                    pass_results.append({"repeat": k, "mode": mode, "result": str(r)})
                    if r != "ok":
                        raise RunnerError(f"{mode} pass {k} did not complete: {r}")
                    st["vba_path"] = path
        with wd.guard("clock_probe", opts.timeout_s):
            probe_clock(probes_after)

        # -- collect_trace --------------------------------------------------
        with rec.stage("collect_trace") as st:
            total = 0
            for s in rec.stages:
                if "vba_path" not in s:
                    continue
                try:
                    events, footer = trace.read_vba_events(s["vba_path"])
                except trace.TraceError as exc:
                    raise RunnerError(f"{s['mode']} pass {s['pass']}: {exc}") from exc
                probs = check_vba_footer(footer)
                if probs:
                    raise RunnerError(f"{s['mode']} pass {s['pass']}: " + "; ".join(probs))
                pass_events[(s["pass"], s["mode"])] = events
                passes_for_check.append((s["pass"], s["mode"], events))
                total += len(events)
            st["attrs"] = {"events": total}

        # -- close_excel ----------------------------------------------------
        with rec.stage("close_excel"):
            with wd.guard("close_excel", opts.timeout_s):
                warnings.extend(release_excel())
            ensure_exited()

        # -- verify_output --------------------------------------------------
        with rec.stage("verify_output") as st:
            verify_unchanged(wb_path, orig_sha)
            probs = check_expected(passes_for_check, plan)
            if probs:
                raise RunnerError("missing instrumentation: " + "; ".join(probs[:10]))
            critical, other = span_failures(passes_for_check)
            run["step_errors"] = critical + other
            if critical:
                c = critical[0]
                raise RunnerError(
                    f"{len(critical)} workbook-level span(s) did not complete; first: {c['kind']} "
                    f"in {c['mode']} pass {c['pass']} ({c['status']}, error {c['error_code']})")
            if other:
                warnings.append(f"{len(other)} step span(s) ended with an error (see step_errors)")
            offset_after, uncert_after = choose_probe(probes_after)
            drift = offset_after - offset_ns
            drift_status = check_drift(drift, max(uncert_ns, uncert_after))
            run["clock"] = {
                "vba_offset_ns": offset_ns, "vba_offset_uncertainty_ns": uncert_ns,
                "probes": len(probes_before), "offset_after_ns": offset_after,
                "offset_after_uncertainty_ns": uncert_after, "drift_ns": drift, "drift_status": drift_status,
            }
            if drift_status == "fail":
                raise RunnerError(f"VBA/host clock drift {drift} ns exceeds {DRIFT_FAIL_FACTOR}x the probe uncertainty")
            if drift_status == "warn":
                warnings.append(f"VBA/host clock drift {drift} ns exceeds the probe uncertainty")
            st["attrs"] = {"count": len(passes_for_check)}

        # -- write traces ---------------------------------------------------
        run["excel"] = excel_info
        run["passes"] = pass_results
        header = {
            "run_id": run_id,
            "created_utc": _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "tool_version": __version__,
            "host": {
                "os": platform.platform(), "python": platform.python_version(),
                "excel_version": excel_info.get("version"), "excel_build": excel_info.get("build"),
                "bitness": excel_info.get("bitness"), "threads": excel_info.get("threads"),
                "calc_mode_original": calc_original,
            },
            "clock": {"host": "perf_counter_ns", "vba": "MicroTimer/QPC",
                      "vba_offset_ns": offset_ns, "vba_offset_uncertainty_ns": uncert_ns},
            "source": {"kind": "real", "label": opts.label or f"workbook:{orig_sha[:10]}"},
            "redaction": {"names": opts.names},
        }
        per_trace_max = max(trace.DEFAULT_MAX_EVENTS, 2 * opts.repeats * (opts.max_events + 4) + 64)

        def merger(mode):
            def merge(w, st):
                if st.get("mode") == mode:
                    merge_vba_events(w, pass_events[(st["pass"], mode)], ident=ident, meta=meta,
                                     attr_ok=attr_value_ok)
            return merge

        on_path, off_path = out_dir / "trace-on.jsonl", out_dir / "trace-off.jsonl"
        expect: Dict[str, List[str]] = {}
        planned_sheets = [s[1] for s in plan if s[0] == "sheet"]
        if planned_sheets:
            expect["calc.sheet"] = sorted({ident(s) for s in planned_sheets})
        if opts.macros:
            expect["vba.proc"] = sorted({ident(m) for m in opts.macros})
        write_trace(on_path, dict(header, mode="trace-on", expect=expect), rec.stages,
                    max_events=per_trace_max, include=lambda s: True, merge=merger("on"))
        write_trace(off_path, dict(header, mode="trace-off"), rec.stages,
                    max_events=per_trace_max, include=lambda s: s.get("mode") == "off", merge=merger("off"))

        validation = {}
        for key, p in (("trace-on", on_path), ("trace-off", off_path)):
            res = trace.validate(p)
            validation[key] = {"ok": res.ok, "errors": res.errors[:50], "warnings": res.warnings[:50]}
        run["validation"] = validation
        run["traces"] = {"trace_on": on_path.name, "trace_off": off_path.name}
        run["formulas"] = "formulas.json"
        if not all(v["ok"] for v in validation.values()):
            raise RunnerError("written traces failed validation: " + "; ".join(
                f"{k}: {v['errors'][:3]}" for k, v in validation.items() if not v["ok"]))

        # -- render_report (timing recorded in run.json only) ---------------
        t0 = time.perf_counter_ns()
        run["render_report"] = {"dur_ns": None, "status": "error"}
        try:
            html = report.render([on_path, off_path], out_dir / "formulas.json", out_dir / "report.html")
        except Exception as exc:
            raise RunnerError(f"report rendering failed: {_com_err(exc)}") from exc
        finally:
            run["render_report"]["dur_ns"] = time.perf_counter_ns() - t0
        run["render_report"]["status"] = "ok"
        run["report"] = Path(html).name
        run["ok"] = True
    except RunnerError as exc:
        run["error"] = str(exc)
        raise
    except Exception as exc:
        run["error"] = f"unexpected: {_com_err(exc)}"
        raise RunnerError(run["error"]) from exc
    finally:
        if xl is not None or wb is not None:
            try:
                with wd.guard("cleanup", min(opts.timeout_s, 120.0)):
                    warnings.extend(release_excel())
            except Exception as exc:
                warnings.append(f"cleanup: {_com_err(exc)}")
            xl = wb = None
        gc.collect()
        if co_init:
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()
        if not run["ok"]:
            ensure_exited()
        wd.stop()
        if orig_sha is not None and not run["ok"]:
            try:
                verify_unchanged(wb_path, orig_sha)
            except RunnerError as exc:
                warnings.append(str(exc))
        run["stages"] = rec.summary()
        if not opts.keep_copy and work.exists():
            shutil.rmtree(work, ignore_errors=True)
            if work.exists():
                warnings.append("work copy could not be deleted")
            with contextlib.suppress(OSError):
                (out_dir / "work").rmdir()
        elif opts.keep_copy:
            run["work_copy"] = "work/" + run_id + "/" + (ident(wb_path.stem) + wb_path.suffix if hashed else wb_path.name)
        scrub_ident = ident if hashed else None
        if run["error"]:
            run["error"] = scrub_text(run["error"], paths=secret_paths, idents=secret_idents, ident=scrub_ident)
        run["warnings"] = [scrub_text(w, paths=secret_paths, idents=secret_idents, ident=scrub_ident)
                           for w in warnings]
        _write_json(out_dir / "run.json", run)
    return run


def _count_kinds(plan: Sequence[Sequence[str]]) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for s in plan:
        out[s[0]] = out.get(s[0], 0) + 1
    return out


def _write_json(path: Path, obj: dict) -> None:
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(obj, indent=1, default=str) + "\n", encoding="utf-8")
    os.replace(tmp, path)


def _import_module(wb, bas: Path) -> None:
    """Import the .bas into the copy's VBA project, failing closed on access problems."""
    try:
        project = wb.VBProject
        comps = project.VBComponents
    except Exception as exc:
        raise RunnerError(
            "Excel refused programmatic access to the VBA project. Enable it for this run: "
            "File > Options > Trust Center > Trust Center Settings > Macro Settings > "
            "'Trust access to the VBA project object model'. It can be switched off again afterwards. "
            f"({_com_err(exc)})"
        ) from exc
    try:
        if int(project.Protection) == 1:
            raise RunnerError("the workbook's VBA project is password-protected; the timer module cannot be imported")
    except RunnerError:
        raise
    except Exception:
        pass
    for i in range(int(comps.Count), 0, -1):
        c = comps.Item(i)
        if str(c.Name).lower() == VBA_MODULE.lower():
            comps.Remove(c)  # an older copy shipped with the workbook; the copy is never saved
        del c
    try:
        comps.Import(str(bas))
    except Exception as exc:
        raise RunnerError(f"importing {VBA_MODULE}.bas failed: {_com_err(exc)}") from exc
    names = [str(comps.Item(i).Name) for i in range(1, int(comps.Count) + 1)]
    if VBA_MODULE not in names:
        raise RunnerError(f"{VBA_MODULE} module missing after import (got a renamed copy?)")
    if int(comps.Item(VBA_MODULE).Type) != VBEXT_CT_STD_MODULE:
        raise RunnerError(f"{VBA_MODULE} is not a standard module after import")


def _excel_pid(xl) -> Optional[int]:
    try:
        import win32process  # type: ignore

        return int(win32process.GetWindowThreadProcessId(int(xl.Hwnd))[1])
    except Exception:
        return None


def _wait_pid_exit(pid: int, timeout_s: float) -> bool:
    """True once the process has exited (or cannot be opened any more)."""
    try:
        import win32api  # type: ignore
        import win32con  # type: ignore
        import win32event  # type: ignore
    except ImportError:
        return True
    try:
        h = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return True  # already gone
    try:
        return win32event.WaitForSingleObject(h, int(timeout_s * 1000)) == win32event.WAIT_OBJECT_0
    finally:
        win32api.CloseHandle(h)


def _terminate_pid(pid: int) -> None:
    """Terminate our own isolated Excel instance (timeout or failed Quit). No COM."""
    try:
        import win32api  # type: ignore
        import win32con  # type: ignore

        h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
        win32api.TerminateProcess(h, 1)
        win32api.CloseHandle(h)
    except Exception:
        pass
