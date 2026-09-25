"""Command-line entry point for XLSprint.

Commands:
  profile   Time a disposable copy of a workbook in Excel Desktop (Windows only).
  inspect   Static formula diagnostics (any OS; never launches Excel).
  report    Render the standalone HTML report from existing traces.
  validate  Fail-closed validation of a trace file.
  selftest  Synthetic end-to-end pipeline check (no Excel, no real data).
"""

from __future__ import annotations

import argparse
import json
import os
import re
import secrets
import sys
from pathlib import Path

from xlsprint import __version__


def _out_dir(value: str) -> Path:
    """Resolve a user-selected output directory, refusing network paths."""
    if value.startswith(("\\\\", "//")):
        raise argparse.ArgumentTypeError("output directory must be local, not a network (UNC) path")
    path = Path(value).expanduser().resolve()
    path.mkdir(parents=True, exist_ok=True)
    return path


def _cmd_inspect(args: argparse.Namespace) -> int:
    from xlsprint import formulas

    salt = args.salt or secrets.token_hex(8)
    result = formulas.inspect_workbook(args.workbook, names=args.names, salt=salt)
    out = args.out / "formulas.json"
    formulas.write_formulas_json(result, out)
    totals = result.get("totals", {})
    print(f"formulas.json written: {out}")
    print(f"  formula cells {totals.get('formula_cells', 0)}, groups {totals.get('groups', 0)}, "
          f"volatile cells {totals.get('volatile_cells', 0)}, UDF cells {totals.get('udf_cells', 0)}, "
          f"single-thread cells {totals.get('single_thread_cells', 0)}  [static counts, not timings]")
    return 0


MAX_LISTED_ERRORS = 20
_RULE_RE = re.compile(r"\[(\d+)\]")


def _by_rule(items: list[str]) -> list[str]:
    """Stable-sort validator messages by rule number: low rules are root causes."""
    def key(item: str) -> int:
        m = _RULE_RE.search(item)
        return int(m.group(1)) if m else 99
    return sorted(items, key=key)


def _print_limited(label: str, items: list[str]) -> None:
    """Print at most MAX_LISTED_ERRORS items; one structural fault can cascade."""
    for item in _by_rule(items)[:MAX_LISTED_ERRORS]:
        print(f"  {label}: {item}")
    if len(items) > MAX_LISTED_ERRORS:
        print(f"  ... {len(items) - MAX_LISTED_ERRORS} more {label}s (first ones listed are usually the cause)")


def _cmd_validate(args: argparse.Namespace) -> int:
    from xlsprint import trace

    worst = 0
    for path in args.traces:
        result = trace.validate(path)
        status = "VALID" if result.ok else "INVALID"
        print(f"{status}: {path}")
        _print_limited("error", result.errors)
        _print_limited("warning", result.warnings)
        if not result.ok:
            worst = 1
    return worst


def _cmd_report(args: argparse.Namespace) -> int:
    from xlsprint import report

    traces = [args.trace_on] + ([args.trace_off] if args.trace_off else [])
    out_html = args.out / "report.html"
    report.render(traces, args.formulas, out_html, allow_invalid=args.allow_invalid)
    print(f"report written: {out_html}")
    return 0


def _cmd_selftest(args: argparse.Namespace) -> int:
    """Run the offline pipeline on generated data and label it SYNTHETIC."""
    from xlsprint import formulas, report, synth, synthbook, trace

    out = args.out
    book = synthbook.make_synthetic_workbook(out / "synthetic-workbook.xlsx", seed=args.seed)
    inspected = formulas.inspect_workbook(book, names="clear", salt="selftest")
    formulas_path = out / "formulas.json"
    formulas.write_formulas_json(inspected, formulas_path)
    on_path, off_path = synth.synthetic_traces(out, seed=args.seed)
    failures = 0
    for path in (on_path, off_path):
        result = trace.validate(path)
        print(f"{'VALID' if result.ok else 'INVALID'}: {path.name}")
        _print_limited("error", result.errors)
        failures += not result.ok
    out_html = report.render([on_path, off_path], formulas_path, out / "report.html")
    print(f"report written: {out_html}")
    print("NOTE: every number in this output is SYNTHETIC; no workbook was timed.")
    return 1 if failures else 0


def _cmd_profile(args: argparse.Namespace) -> int:
    from xlsprint import runner

    opts = runner.ProfileOptions(
        workbook=Path(args.workbook).resolve(),
        out_dir=args.out,
        repeats=args.repeats,
        blocks=args.blocks,
        ranges=args.range or [],
        macros=args.macro or [],
        names_timing=not args.no_names_timing,
        group_timing=args.group_timing,
        full_rebuild=args.full_rebuild,
        names=args.names,
        keep_copy=args.keep_copy,
        label=args.label,
        max_events=args.max_events,
        timeout_s=args.timeout,
        enable_events=args.enable_events,
        argv=sys.argv[1:],
    )
    try:
        result = runner.profile(opts)
    except runner.RunnerError as exc:
        print(f"profile failed (fail-closed): {exc}", file=sys.stderr)
        return 2
    print(json.dumps({k: result.get(k) for k in ("run_id", "report", "traces") if k in result}, indent=2))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="xlsprint", description=__doc__.splitlines()[0])
    parser.add_argument("--version", action="version", version=f"xlsprint {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    def add_out(p: argparse.ArgumentParser) -> None:
        p.add_argument("--out", type=_out_dir, required=True, help="local output directory (created if missing)")

    p = sub.add_parser("profile", help="profile a disposable copy in Excel Desktop (Windows)")
    p.add_argument("workbook")
    add_out(p)
    p.add_argument("--repeats", type=int, default=5, help="matched trace-off/trace-on pairs (default 5)")
    p.add_argument("--blocks", type=int, default=4, help="auto column blocks per sheet (default 4)")
    p.add_argument("--range", action="append", metavar="SHEET!A1:B2", help="extra range to time (repeatable)")
    p.add_argument("--macro", action="append", metavar="NAME", help="instrumented macro to run and require in trace")
    p.add_argument("--no-names-timing", action="store_true", help="skip timing defined-name ranges")
    p.add_argument("--group-timing", action="store_true", help="opt-in isolated Range.Calculate per formula group")
    p.add_argument("--full-rebuild", action="store_true", help="also time CalculateFullRebuild")
    p.add_argument("--names", choices=["hashed", "clear"], default="hashed", help="identifier redaction (default hashed)")
    p.add_argument("--keep-copy", action="store_true", help="keep the disposable workbook copy")
    p.add_argument("--label", help="run label shown in the report")
    p.add_argument("--max-events", type=int, default=100_000)
    p.add_argument("--timeout", type=float, default=900.0, metavar="SECONDS",
                   help="deadline per pass and per Excel stage; Excel is terminated when it expires (default 900)")
    p.add_argument("--enable-events", action="store_true", help="leave Application.EnableEvents on while profiling")
    p.set_defaults(func=_cmd_profile)

    p = sub.add_parser("inspect", help="static formula diagnostics")
    p.add_argument("workbook")
    add_out(p)
    p.add_argument("--names", choices=["hashed", "clear"], default="hashed")
    p.add_argument("--salt", help="redaction salt (default: random per run)")
    p.set_defaults(func=_cmd_inspect)

    p = sub.add_parser("report", help="render report.html from traces")
    p.add_argument("trace_on", type=Path)
    p.add_argument("--trace-off", type=Path)
    p.add_argument("--formulas", type=Path)
    add_out(p)
    p.add_argument("--allow-invalid", action="store_true", help="render an invalid trace with an INVALID banner")
    p.set_defaults(func=_cmd_report)

    p = sub.add_parser("validate", help="validate trace files (exit 1 if any invalid)")
    p.add_argument("traces", nargs="+", type=Path)
    p.set_defaults(func=_cmd_validate)

    p = sub.add_parser("selftest", help="synthetic offline pipeline check")
    add_out(p)
    p.add_argument("--seed", type=int, default=0)
    p.set_defaults(func=_cmd_selftest)
    return parser


def _tool_errors() -> tuple[type[BaseException], ...]:
    from xlsprint.report import ReportError
    from xlsprint.trace import TraceError

    return (OSError, ValueError, TraceError, ReportError)


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if os.environ.get("XLSPRINT_DEBUG"):
        return args.func(args)
    try:
        return args.func(args)
    except _tool_errors() as exc:
        lines = str(exc).splitlines() or [type(exc).__name__]
        print(f"xlsprint: {lines[0]}", file=sys.stderr)
        _print_limited_err(lines[1:])
        return 1


def _print_limited_err(lines: list[str]) -> None:
    for line in _by_rule(lines)[:MAX_LISTED_ERRORS]:
        print(line, file=sys.stderr)
    if len(lines) > MAX_LISTED_ERRORS:
        print(f"  ... {len(lines) - MAX_LISTED_ERRORS} more", file=sys.stderr)
