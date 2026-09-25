# XLSprint design contract

This file is the integration contract between modules. Change it only
deliberately; every module and test depends on it.

XLSprint is an independent, open-source Excel calculation profiler. It follows
Microsoft's published calculation-performance guidance (MicroTimer, and the
workbook → worksheet → range drill-down) from
<https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance>.
It is not affiliated with FastExcel. It does not reproduce FastExcel's code or
its proprietary profiling method.

## Supported platform

- **Profiling runs:** Windows 10/11 with Excel Desktop 2013 or later (32- or
  64-bit, VBA7). The run is driven over COM (pywin32). MicroTimer calls
  `kernel32` `QueryPerformanceCounter`, which makes the method Windows-only.
- **Offline commands** (`inspect`, `report`, `validate`, `selftest`): any OS
  with Python ≥ 3.9 and openpyxl. They never launch Excel.
- Excel for Mac and Excel for the web are not supported for timing.

## Package layout (file ownership)

| Path | Owner | Purpose |
|---|---|---|
| `xlsprint/trace.py` | trace | Event schema, bounded writer, reader, fail-closed validator |
| `xlsprint/analysis.py` | trace | Span tree, self time, interval-union totals, hotspots, stats, overhead |
| `xlsprint/formulas.py` | formulas | Static formula inspection and grouping (openpyxl) |
| `xlsprint/report.py` | report | Standalone HTML report |
| `xlsprint/vba/XLSprintTimer.bas` | excel | VBA MicroTimer, trace API, drill-down driver |
| `xlsprint/runner.py` | excel | Windows COM host orchestration and host-stage spans |
| `xlsprint/cli.py`, `xlsprint/__main__.py` | lead | Command-line entry point |
| `tests/fixtures/` | each owner | Synthetic fixtures only; no real workbook data |

Only stdlib and openpyxl are required. pywin32 is the optional `[windows]` extra
and may be imported only inside `runner.py`, lazily.

## Clock

- Host (Python): `time.perf_counter_ns()`. It is monotonic. On Windows it is
  backed by QueryPerformanceCounter.
- VBA: `MicroTimer()` as published by Microsoft (QPC ticks / QPC frequency,
  in seconds). The trace records `ns = floor(QPC ticks × 1e9 / QPC frequency)`,
  computed exactly in VBA Decimal from the raw ticks, outside timed regions.
  It is an integer nanosecond value from an arbitrary epoch.
- The two clock sources are aligned by a **clock probe**. The host records
  `h0 = perf_counter_ns()`, calls the VBA function `XSP_ClockProbe()` (which
  returns MicroTimer ns), then records `h1`. The offset is `vba_ns − (h0+h1)/2`,
  with uncertainty `±(h1−h0)/2`. The trace header stores the offset and
  uncertainty. A VBA span is never compared to a host span more finely than
  that uncertainty.
- Every event has `clock: "host"` or `clock: "vba"`. The analyser converts VBA
  times to the host timebase using the offset.

## Trace format (`trace.jsonl`)

JSON Lines, UTF-8. Line 1 is the header, the last line is the footer, and
everything between is an event.

### Header
```json
{"type":"header","schema":"xlsprint.trace/1","run_id":"<uuid4 hex>","mode":"trace-on|trace-off",
 "created_utc":"2026-09-25T12:00:00Z","tool_version":"0.1.0",
 "host":{"os":"Windows-10...","python":"3.12.1","excel_version":"16.0","excel_build":"...","bitness":64,"threads":8,"calc_mode_original":"automatic"},
 "clock":{"host":"perf_counter_ns","vba":"MicroTimer/QPC","vba_offset_ns":123,"vba_offset_uncertainty_ns":4000},
 "limits":{"max_events":200000,"max_bytes":67108864},
 "source":{"kind":"synthetic|real","label":"<user label or 'synthetic:<fixture>'>"},
 "redaction":{"names":"clear|hashed"}}
```
`source.kind` is mandatory. The report shows synthetic runs with a
"SYNTHETIC — not a real-workbook measurement" banner.

### Events
```json
{"type":"B","id":17,"parent":3,"pass":2,"depth":2,"ns":123456789,"clock":"vba","kind":"calc.sheet","name":"Sheet2","attrs":{"method":"Worksheet.Calculate","repeat":1}}
{"type":"E","id":17,"ns":123460000,"clock":"vba","status":"ok"}
{"type":"M","id":18,"parent":3,"pass":2,"ns":123460100,"clock":"vba","kind":"marker","name":"AfterCalculate","attrs":{}}
```
- `B`/`E` are balanced begin and end events. `id` is unique per run (a positive
  int). An `E` carries only `id`, `ns`, `clock`, and `status`
  (`ok|error|aborted`), plus optional `attrs`.
- `parent` is the id of the enclosing open span, or 0 for none. `depth` is the
  nesting depth, where the root is 0.
- `pass` is the calculation-pass identifier. It is 0 for host stages outside
  any profiling pass. Each repeat of the drill-down sequence is its own pass
  (1..N).
- `M` is an instantaneous marker with no duration.
- `name` holds only structural identifiers: a sheet name, range address,
  defined-name name, procedure name, stage name, or formula-group id. With
  `redaction.names == "hashed"`, sheet, name, and procedure identifiers are
  replaced by `h:<first 10 hex of sha256(salt+name)>`. Range addresses keep
  their A1 coordinates but take a hashed sheet prefix.
- **Never** in a trace: cell values, formula text, string constants, file
  contents. Enforce this with an allow-list of attribute keys (see
  `trace.ALLOWED_ATTR_KEYS`). Unknown keys are rejected by both the writer
  and the validator.

### Span kinds (closed vocabulary)
| kind | clock | meaning |
|---|---|---|
| `host.stage` | host | End-to-end stage: `prepare_copy`, `launch_excel`, `open_workbook`, `instrument`, `warmup` (includes the clock probe), `profile_trace_off`, `profile_trace_on`, `collect_trace`, `close_excel`, `verify_output`, `render_report` |
| `run.pass` | vba | One repeat of the drill-down sequence |
| `calc.full` | vba | `Application.CalculateFull` |
| `calc.fullrebuild` | vba | `Application.CalculateFullRebuild` (opt-in) |
| `calc.recalc` | vba | `Application.Calculate` |
| `calc.sheet` | vba | `Worksheet.Calculate` |
| `calc.range` | vba | `Range.Calculate` on a user block or an auto block |
| `calc.name` | vba | `Range.Calculate` on the range a defined name refers to |
| `calc.group` | vba | `Range.Calculate` on the cells of one formula group (isolated) |
| `vba.proc` | vba | User-instrumented VBA procedure (`XSP_Begin`/`XSP_End`) |
| `marker` | either | Instant marker |

### Allowed attribute keys
`method, repeat, cells, areas, address, sheet, group, calc_mode, threads,
multithreaded, iteration, error_code, note, count, overhead_mode, bytes, events,
expected, found, sha256_prefix`.

Values must be numbers, booleans, or strings of ≤ 256 characters that match
`^[\w .:$!'#\-/()\[\],@]*$` (Python `re`, Unicode `\w`, so non-ASCII sheet
names are allowed). The validator rejects anything else.
`=` and `"` are excluded, so formula text and string literals cannot pass.
Allow-listed keys are never populated from cell contents.

### Footer
```json
{"type":"footer","events":1234,"dropped":0,"truncated":false,"open_spans_at_close":0,"bytes":98765}
```

## Fail-closed validation (`trace.validate(path) -> ValidationResult`)

A trace is **invalid**, and the report refuses to render timing sections
without `--allow-invalid` (which stamps an "INVALID TRACE" banner), when any of
the following holds:
1. The header or footer is missing, or the schema is unknown.
2. An `E` has no matching open `B`, an `E` closes a span that is not the
   innermost open span (improper nesting), or a `B` is still open at the footer.
3. `ns` goes backwards within one clock between consecutive events, or an
   `E.ns < B.ns`.
4. A duplicate `id` appears, or a `parent`/`depth` is inconsistent with the
   open-span stack.
5. The footer has `truncated == true` or `dropped > 0`, or the footer event
   count differs from the observed count.
6. An attribute key is not allow-listed, or an attribute value fails the
   pattern.
7. **Missing instrumentation:** the header `expect` map (optional
   `"expect":{"vba.proc":["Name",...],"calc.sheet":["*"]}`) names spans that
   never appeared.
8. A `trace-on` run lacks `run.pass` spans, or lacks a `calc.full` or
   `calc.recalc` inside every pass.

9. **Structure** (added after review):
   - Header sub-objects (`host`, `clock`, `source`, `redaction`, `limits`)
     must be objects, and `expect` must map kinds to lists of strings.
   - `ns` must be an int with 0 ≤ ns < 2**63.
   - Every `calc.*` and `vba.proc` span must lie inside a `run.pass`, with
     the matching `pass`.
   - A VBA span whose parent is a host span must lie inside that parent
     after alignment, within ±`vba_offset_uncertainty_ns`. This catches a
     wrong-sign offset.
   - The event count and file size must not exceed `header.limits`.
10. **Names** must match `NAME_RE` = `^[\w .:$!'#\-/()\[\],@&%]*$`
    (≤ 256 chars). That excludes `=`, `"`, `*`, `+`, `^`, `<`, `>`, and
    control characters. The runner hashes any clear identifier that fails
    this pattern, even with `--names clear`, and never drops the span.
    `attrs.note` holds tool-generated labels only.

`ValidationResult` = `{ok: bool, errors: [str], warnings: [str], stats: {...}}`.

## Bounded output

The writer holds `max_events` (default 200 000) and `max_bytes` (default 64 MiB).
On overflow it stops writing events, counts `dropped`, and sets
`truncated=true` in the footer. A truncated trace is invalid (fail closed). The
VBA side has its own bound (`XSP_MAX_EVENTS`, default 100 000) and reports
overflow in the footer it writes.

## Measurement classes (the report labels every number with one)

| Class | Badge | Examples |
|---|---|---|
| **DIRECT** | `measured` | Wall time of one Excel calc call or host stage between balanced events (MicroTimer or perf_counter). Includes all Excel work triggered by that call. The median and min–max over repeats of *the same operation* are still `measured`, and always show `n`. |
| **DERIVED** | `derived` | Arithmetic that combines different measurements: self time (span minus union of children), volatility ratio (recalc / full), share of pass time, overhead (on − off). Every formula is shown. |
| **STRUCTURAL** | `static` | Read, not timed: formula counts, group sizes, function features (volatile, UDF, single-thread list) from the file, and run settings (thread count, repeats, calc mode) from Excel or the command line. |
| **UNAVAILABLE** | `n/a` | Not measurable by this method: per-formula runtime inside a normal recalc, thread placement, calc-chain reorder time. |

### Formula-group drill-down decision
`calc.group` times `Range.Calculate` over the cells of one formula group with
dependencies outside the range held at their current values. That timing is
**DIRECT** for *that operation*. It is **not** the group's share of a normal
recalc: it excludes calc-chain scheduling and cross-group dependency effects,
and Range.Calculate may thread differently. The report therefore:
- shows group isolated timings as `measured (isolated Range.Calculate)`;
- never divides by cell count to show per-formula time;
- never sums group timings to "explain" sheet time;
- makes the feature opt-in (`--group-timing`), limited to groups whose cells
  form ≤ 32 rectangular areas, with ≥ 3 repeats, reporting median and range.

`share_of_recalc_sum` = sheet median ÷ Σ sheet medians. It is DERIVED and only
a ranking aid: sheet calculations are separate calls, and their times are not
additive parts of the workbook recalc time.

## Human-authored semantics

An optional `xlsprint.semantics/1` JSON sidecar labels workbook regions and
VBA procedure spans with owner-supplied business intent. The profile checks
its optional full workbook SHA-256, resolves clear selectors while the
workbook plan is in memory, then stores only the same redacted identifiers as
the trace plus the supplied label, category, intent and source in
`formulas.json`. Raw formulas and values are not used for annotation matching.
The report shows semantic descriptions before technical identifiers and cell
addresses, and marks selectors with no observed timing explicitly. Inferred
intent is outside this contract and must be labelled as a suggestion with its
own confidence and provenance if introduced later.

The map annotates measured regions; it does not split a range's elapsed time
among formulas or change the meaning of `Range.Calculate` measurements.

## Double-counting rule

A total over a set of spans is the length of the **union of their intervals**
on one timeline, per clock domain after alignment. It is never the sum of
their durations. Self time = duration − |union(children)|. Hotspot ranking uses
the median self time across passes, and within a pass, a range and a name
covering the same cells are listed separately and never added together.

## Overhead

A pass is one VBA call, `XSP_RunPass(passId, mode, planPath)`, that runs the
whole drill-down plan inside Excel, so no COM round-trips fall inside timed
regions. `profile` runs R interleaved matched pairs in the order
off, on, off, on, …:
- **trace-off** runs *exactly the same calc calls in the same order* and
  records only `run.pass`, `calc.full`, and `calc.recalc` spans.
- **trace-on** records every span in the plan.

Pairs are matched **by pass number**. Trace-off pass k runs immediately
before trace-on pass k, and both carry `pass` = k. Both files must carry the
same `run_id`, or overhead is unavailable.
Only passes with status `ok` are paired, and dropped pairs are reported.
Instrumentation overhead is DERIVED as the **median of the per-pair
differences** (on − off), together with their min–max, and likewise for
`calc.full` and `calc.recalc`. It is "resolvable" only with ≥ 3 pairs and when
the min–max of the paired differences excludes 0. Otherwise it is reported as
"not resolvable (below run-to-run noise)". Trace-off events are written to
`trace-off.jsonl` with header `mode: "trace-off"`.

### Drill-down plan (per pass, Microsoft's order)
1. `calc.full`: `Application.CalculateFull` (worst case)
2. `calc.recalc`: `Application.Calculate` immediately after (best case)
3. `calc.sheet` for each worksheet with formulas: `Worksheet.Calculate`
4. `calc.range` for user blocks (`--range`) and auto blocks. Auto blocks are
   the formula area of each sheet split into ≤ `--blocks` (default 4) column
   blocks. As in Microsoft's RangeTimer, the range is extended to cover whole
   CSE array formulas. The host does this from the plan's `array_areas`,
   producing disjoint areas so no cell is calculated twice. Iteration is
   switched off for range timings, then restored.
5. `calc.name` for each defined name that refers to a single-sheet range
   containing formulas (`--names-timing`, on by default)
6. `calc.group` (opt-in `--group-timing`)
7. `vba.proc`: each user macro given with `--macro` (run after the calc steps,
   in a `vba.proc` wrapper span; its internal `XSP_Begin`/`XSP_End` spans nest
   below it)

The plan is written to `work/<run_id>/plan.txt` (UTF-16LE) and read by VBA
before the `run.pass` span opens. VBA only ever sees keys (`R0001`, `N0001`,
`G0001`); trace names come from host-side metadata. Calculation mode is set to
manual for the session. It is *not* restored before close: the instance is
private and the copy is never saved, and restoring automatic mode would force
a recalculation. ScreenUpdating is off, so phase 4
of Excel's calc process (updating visible windows) is excluded. The report
states this.

## Formula inspection output (`formulas.json`)
```json
{"schema":"xlsprint.formulas/1","workbook_sha256_prefix":"ab12cd34ef","sheets":[{"sheet":"Sheet1","formula_cells":1200,"used_range":"A1:F2000"}],
 "groups":[{"group":"G0001","fingerprint":"<sha256[:16] of normalized R1C1 form>","sheet":"Sheet1","cells":1000,"areas":["B2:B1001"],
   "functions":["SUM","OFFSET"],"udfs":1,"udf_names":["h:..."],"volatile":true,"volatile_functions":["OFFSET"],
   "single_threaded":false,"single_threaded_functions":[],"array":false,"dynamic_array":false,"cross_sheet":false,"external_ref":false,
   "whole_column_ref":false,"references":2,"length_bucket":"<64"}],
 "names":[{"name":"InputBlock","scope":"workbook","refers_to_range":"Sheet1!$A$1:$A$100","sheet":"Sheet1","is_range":true}],
 "totals":{"formula_cells":1200,"groups":3,"volatile_cells":10,"udf_cells":0,"single_thread_cells":0}}
```
- **Normalization:** convert to relative R1C1, so `=A1+1` in B1 and `=A2+1` in
  B2 match. Collapse whitespace and uppercase function names. Replace string
  and number literals with placeholders for the fingerprint only; literals
  never appear in output. Formula text never appears in output. In hashed mode
  the fingerprint is `HMAC-SHA256(salt, normalized)[:16]`, so a guessed
  formula can't be confirmed offline, and fingerprints can't be linked across
  runs with different salts. In `--names clear` mode it is a plain
  `sha256(normalized)[:16]`.
- **Spilled dynamic arrays:** a spill anchor counts as 1 formula cell. The
  spilled range is reported separately as `spill_cells` and is not counted
  as formula cells.
- **Group features** are the OR over all member cells. Literal arguments that
  change a detected feature (the CELL info_type) are kept in the normalized
  form, so they split groups.
- Function names: built-ins are output as-is (they are Excel vocabulary, not
  data). Non-built-in function names (UDFs) are hashed by default, and clear
  only with `--names clear`.
- The single-threaded list comes from Microsoft's article: PHONETIC, CELL
  (format/address), INDIRECT, GETPIVOTDATA, the CUBE* functions, ADDRESS with a
  sheet_name argument, ERROR.TYPE, HYPERLINK, and VBA/COM UDFs. Detection is
  syntactic, so CELL and ADDRESS are flagged only when the argument check
  matches. Database functions over PivotTables cannot be detected statically;
  that is documented as unavailable.
- Volatile list: RAND, RANDBETWEEN, NOW, TODAY, OFFSET, INDIRECT, CELL, INFO,
  plus RANDARRAY. UDFs declared with `Application.Volatile` are detected only if
  the VBA source is readable (`.xlsm` vbaProject.bin is not parsed in v1, so
  this is reported as unavailable).

## Report inputs
`report.render(trace_paths: list[Path], formulas_path: Path|None, out_html: Path, *, allow_invalid=False) -> Path`.
It takes one trace-on and optionally one trace-off trace, analyses them with
`analysis`, and embeds everything inline: no network, no external assets.

## Output directory layout (`--out DIR`, user-selected, must be local)
```
DIR/
  work/<run_id>/<copy>.<ext>       disposable copy, same extension as the original (.xlsx/.xlsm); never saved; deleted unless --keep-copy
  trace-on.jsonl  trace-off.jsonl
  formulas.json
  report.html
  run.json                          argv (paths reduced to basenames), versions, validation results
```
The original workbook is opened read-only for hashing only. Its sha256 is
checked before and after, and a change fails the run.

## Python module APIs (binding)

Modules may return extra keys beyond those listed here, such as `iqr_*`,
`areas_total`, and `limitations`. Consumers must ignore keys they don't know,
and the report must never render unknown fields. Errors are raised as
`TraceError` (trace), `ReportError` (report), and `RunnerError` (runner).

### `xlsprint/trace.py`
```python
SCHEMA = "xlsprint.trace/1"
SPAN_KINDS: frozenset[str]            # the closed vocabulary above (+ "marker")
ALLOWED_ATTR_KEYS: frozenset[str]
ATTR_VALUE_RE: re.Pattern
class TraceError(Exception)
def redact_name(name: str, salt: str) -> str            # "h:" + sha256(salt+name)[:10]
class TraceWriter:
    def __init__(self, path, header: dict, *, max_events=200_000, max_bytes=64*2**20)
    def begin(self, kind, name, *, clock="host", pass_=0, attrs=None, ns=None) -> int
    def end(self, span_id, *, status="ok", attrs=None, ns=None) -> None   # raises TraceError on non-innermost
    def marker(self, name, *, clock="host", pass_=0, attrs=None, ns=None) -> int
    def span(self, kind, name, **kw)          # context manager; status="error" if body raises (re-raises)
    def append_event(self, ev: dict) -> None  # import a pre-built B/E/M event (e.g. from VBA); checks ids, nesting, attrs
    def reserve_ids(self, n: int) -> int      # returns first id of a reserved block, for VBA to use
    @property open_span_id -> int             # innermost open span id or 0
    @property open_depth -> int
    def close(self) -> dict                   # writes footer; never fabricates E events for open spans
def read_trace(path) -> tuple[dict, list[dict], dict|None]      # header, events, footer
def read_vba_events(path) -> tuple[list[dict], dict]            # events, {"type":"vba_footer","events","dropped","truncated","open_spans"}
@dataclass class ValidationResult: ok: bool; errors: list[str]; warnings: list[str]; stats: dict
def validate(path) -> ValidationResult
```
The writer uses `time.perf_counter_ns()` when `ns` is None (host clock only).
For `clock="vba"`, `ns` is mandatory.

### `xlsprint/analysis.py`
```python
@dataclass class Span: id, parent, pass_, depth, kind, name, clock, start_ns, end_ns, status, attrs, children: list[int]
    # start_ns/end_ns are in the HOST timebase (VBA times shifted by -vba_offset_ns)
    dur_ns: int (property)
@dataclass class Run: header: dict; spans: dict[int, Span]; markers: list[dict]; validation: ValidationResult
    source_kind -> "synthetic"|"real"; mode -> "trace-on"|"trace-off"
def load_run(path) -> Run
def union_ns(intervals: Iterable[tuple[int,int]]) -> int
def self_time_ns(run, span_id) -> int                      # dur - union(children)
def summarize(values: list[int]) -> dict                   # {n, median, min, max, p25, p75}
def host_stages(run) -> list[dict]                         # {name, start_rel_ns, dur_ns, status}
def timeline(run, *, max_rows=2000) -> list[dict]          # {id, kind, name, pass, depth, start_rel_ns, dur_ns, status}; truncation flagged via last row {"truncated": k}
def passes(run) -> list[dict]                              # {pass, dur_ns, by_kind_union_ns: {kind: ns}}
def drilldown(run) -> dict   # {"workbook": {"full": summary, "recalc": summary, "volatility_ratio": float|None},
                             #  "sheets": [{"name", "recalc": summary, "share_of_recalc_sum": float,
                             #              "children": [{"kind","name","summary","address"}]}]}
def phase_breakdown(run) -> list[dict]                     # {kind, count, union_ns_median_per_pass, share_of_pass}
def hotspots(run, top=25) -> list[dict]                    # {kind, name, sheet, median_ns, median_self_ns, n, min_ns, max_ns, measurement:"measured"|"derived"}
def overhead(on: Run, off: Run) -> list[dict]              # {kind, name, median_on_ns, median_off_ns, diff_ns, diff_min_ns, diff_max_ns, resolvable: bool, pairs}
```
Summaries always carry `n`. Every derived field is named in the report with
its formula.

### `xlsprint/formulas.py`
```python
def inspect_workbook(path, *, names="hashed", salt: str) -> dict   # the formulas.json object
def write_formulas_json(obj, path) -> None
def normalize_formula(formula: str, row: int, col: int) -> str     # internal; returns relative-R1C1 normalized form (never output)
VOLATILE_FUNCTIONS, SINGLE_THREADED_FUNCTIONS, BUILTIN_FUNCTIONS: frozenset[str]
```

### `xlsprint/synth.py` (trace owner)
```python
def synthetic_traces(out_dir, *, seed=0, sheets=3, passes=5) -> tuple[Path, Path]   # (trace-on, trace-off), source.kind="synthetic"
```
### `xlsprint/synthbook.py` (formulas owner)
```python
def make_synthetic_workbook(path, *, seed=0) -> Path   # openpyxl-built .xlsx with only generated numbers; covers volatile, single-threaded, UDF-call, array, cross-sheet, names
```

### `xlsprint/report.py`
```python
def build_model(on: Run, off: Run|None, formulas: dict|None) -> dict   # pure view-model from analysis
def render_html(model: dict) -> str                                    # pure; standalone HTML, inline CSS/JS/SVG
def render(trace_paths, formulas_path, out_html, *, allow_invalid=False) -> Path
```

### `xlsprint/runner.py` (Windows only)
```python
@dataclass class ProfileOptions: workbook: Path; out_dir: Path; repeats=5; blocks=4; ranges: list[str]; macros: list[str];
    names_timing=True; group_timing=False; full_rebuild=False; names="hashed"; keep_copy=False; label: str|None; max_events=100_000
def profile(opts: ProfileOptions) -> dict    # returns run.json content; raises RunnerError (fail-closed) on any check failure
```
