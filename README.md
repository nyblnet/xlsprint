# XLSprint

XLSprint is an open-source (MIT) profiler for Excel calculation. It times a
controlled Excel Desktop run at several levels:

- the end-to-end host stages, from copying the workbook to verifying the output;
- the workbook, each worksheet, and bounded ranges, following Microsoft's
  documented `MicroTimer` method and workbook → worksheet → range drill-down;
- defined names that refer to ranges, and optionally isolated formula groups;
- explicit Excel calculation calls and passes, and VBA procedures you
  instrument yourself.

It also reports static formula diagnostics: identical formulas grouped and
counted, and flags for volatile functions, UDFs, and single-threaded functions.

Everything goes into one standalone HTML report. The report labels every
number by what it is:

| Badge | Meaning |
|---|---|
| `measured` | A direct wall-clock measurement of one Excel call or host stage, between balanced begin/end events. A median or min–max over repeats of the same operation is still `measured`, and shows `n`. |
| `derived` | Arithmetic that combines measurements: self time, interval unions, volatility ratio, shares, instrumentation overhead. The formula is shown next to the number. |
| `static` | Counts read from the file, not timed. |
| `n/a` | Not measurable with this method. |

XLSprint is purpose-comparable to tools such as FastExcel's profiler, but it is
an independent implementation. It shares no code with FastExcel and does not
reproduce FastExcel's proprietary method. The method is Microsoft's published
guidance:
<https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance>.

## What the numbers do *not* mean

Excel has no per-formula execution timer, so XLSprint never shows per-formula
runtime.

- A `calc.range`, `calc.name`, or `calc.group` time is the wall time of
  `Range.Calculate` on those cells alone, with dependencies outside the range
  taken at their current values. It is a valid way to compare blocks and
  alternative formulas (Microsoft's RangeTimer method). It is **not** that
  block's share of a normal recalculation: it leaves out calc-chain scheduling
  and cross-block dependencies, and threading can differ.
- Isolated timings are never added up to "explain" a sheet, and never divided
  by cell count.
- Sheet shares rank sheets against each other. Separate `Worksheet.Calculate`
  calls are not additive parts of `Application.Calculate`.
- Totals over overlapping spans are interval unions, so nested or overlapping
  spans are never double-counted.
- A second calculation is often faster than the first. Timings also vary with
  OS scheduling. XLSprint repeats every pass (default 5) and reports the median
  and min–max.

## Requirements

| Use | Requirement |
|---|---|
| `profile` (timing) | Windows 10 or 11, Excel Desktop 2013 or later (32- or 64-bit, VBA7), Python ≥ 3.9, pywin32. In Excel, turn on *File → Options → Trust Center → Trust Center Settings → Macro Settings → Trust access to the VBA project object model*. |
| `inspect`, `report`, `validate`, `selftest` | Any OS with Python ≥ 3.9 and openpyxl. These commands never start Excel. |

Excel for Mac and Excel for the web are **not supported** for timing.
MicroTimer calls `kernel32` `QueryPerformanceCounter`.

## Setup

```sh
python -m venv .venv
.venv/bin/pip install -e .              # macOS/Linux: offline commands
.venv\Scripts\pip install -e .[windows] # Windows: adds pywin32 for profile
```

## Usage

```sh
# Offline pipeline check on generated data (labelled SYNTHETIC in the report)
xlsprint selftest --out ./xlsprint-out/selftest

# Static formula diagnostics only
xlsprint inspect Book.xlsx --out ./xlsprint-out/book

# Full profile of a disposable copy (Windows + Excel Desktop)
xlsprint profile Book.xlsm --out D:\xlsprint-out\book --repeats 5 --label "Book baseline"
    [--range "Model!B2:K5000"]...   # extra blocks to time
    [--macro RefreshModel]...       # instrumented macro to run; the run fails if its spans are missing
    [--group-timing]                # opt-in isolated Range.Calculate per formula group
    [--full-rebuild]                # also time CalculateFullRebuild
    [--names clear]                 # default is hashed identifiers
    [--keep-copy]
    [--timeout 900]                 # seconds per pass / Excel stage

# Re-render or check existing traces
xlsprint report out/trace-on.jsonl --trace-off out/trace-off.jsonl --formulas out/formulas.json --out out
xlsprint validate out/trace-on.jsonl out/trace-off.jsonl
```

To add human-readable business labels and intent, pass an optional JSON
`xlsprint.semantics/1` sidecar. The profile checks its workbook SHA-256,
resolves readable selectors in memory, and shows the supplied meaning beside
measured regions while preserving cell references as technical locators:

```sh
xlsprint profile Book.xlsm --out D:\xlsprint-out\book --semantics Book-meaning.json
```

See [`docs/SEMANTICS.md`](docs/SEMANTICS.md) for the sidecar contract and
privacy behavior. XLSprint displays owner-supplied intent; it does not claim to
infer financial meaning from formula text.

### What `profile` does

1. **prepare_copy.** Hashes the original (sha256) and copies it into
   `OUT/work/<run_id>/`. The original is never opened in Excel.
2. **launch_excel.** Starts a new, isolated, invisible Excel instance with
   alerts off.
3. **open_workbook.** Opens the copy with links not updated.
4. **instrument.** Imports `xlsprint/vba/XLSprintTimer.bas` into the copy. The
   copy is never saved.
5. **warmup.** Runs one untimed full calculation.
6. **clock probe.** Aligns the VBA MicroTimer clock with the host clock and
   records the uncertainty.
7. **profile passes.** Runs R matched pairs, alternating trace-off and
   trace-on. Each pass is a single VBA call, `XSP_RunPass`, which reads the
   plan from `work/<run_id>/plan.txt` before its timer starts. Range steps
   are extended in Python to cover whole CSE array formulas, without
   overlapping areas. The pass runs
   Microsoft's drill-down order:
   - full calc, then recalc;
   - each sheet;
   - range blocks;
   - names;
   - groups (opt-in);
   - macros.

   Calculation mode is manual. ScreenUpdating is off.
8. **collect_trace**, then **close_excel.** Closes the copy without saving
   and quits Excel. Settings are not restored first, because restoring
   automatic mode would force a pointless recalculation in a throwaway
   instance. XLSprint then checks that the Excel process exited, and
   terminates it if it didn't.
9. **verify_output.** Checks that the original's hash is unchanged, that both
   traces validate fail-closed, and that every planned sheet and requested
   macro appears in the trace.
10. Writes `report.html`, and deletes the copy unless `--keep-copy` is set.

Any failed check stops the run with a non-zero exit. A run also fails when
any `run.pass`, `calc.full`, or `calc.recalc` step doesn't finish `ok`, or
when clock drift is far larger than the probe uncertainty. Every pass and
every Excel stage has a deadline (`--timeout`, default 900 s). If a dialog or
hang makes Excel miss its deadline, the Excel process is terminated and the
run fails, instead of hanging.

## Output formats

`OUT/` is a local directory you choose. Network (UNC) paths are refused.

| File | Format |
|---|---|
| `trace-on.jsonl`, `trace-off.jsonl` | JSON Lines: a header, balanced `B`/`E` events and `M` markers, then a footer. Events carry `id`, `parent`, `pass`, `depth`, `ns`, `clock`, `kind`, `name`, and allow-listed `attrs`. Schema: [docs/DESIGN.md](docs/DESIGN.md). |
| `formulas.json` | Structural diagnostics: groups by fingerprint, counts, feature flags, areas, defined names and (when supplied) semantic annotations. |
| `report.html` | A standalone page (inline CSS, JS, and SVG, with no network access). It contains the route diagram, timeline, drill-down, hotspots, phase breakdown, overhead, model meaning, formula diagnostics, and measurement limits. |
| `run.json` | Arguments (paths reduced to basenames), versions, Excel version/build/bitness/threads, clock-probe results, and validation results. |

Traces are bounded (`--max-events`, default 100 000 on the VBA side and
200 000 on the host). An overflow marks the trace truncated, and a truncated
trace fails validation instead of producing a partial report.

## Privacy

- **Never written:** cell values, formula text, string or number literals, or
  file contents. Formulas are identified only by a salted hash of their
  normalised form (`fingerprint`) and a group id (`G0001`). Both the trace writer and the
  validator enforce an attribute allow-list and a value pattern that excludes
  `=` and `"`.
- **Identifiers:** sheet names, defined names, UDF or procedure names, and
  identifiers in `run.json` arguments and error text are hashed by default
  (`h:` + sha256(salt+name)[:10], with a random salt per run). Range
  addresses keep their A1 coordinates, and local paths in persisted error text
  become `<path>`. `--names clear` keeps names readable. Even then, any name
  containing characters such as `=` `"` `*` `+` `<` `>` is hashed.
- **Local only:** all output stays in your output directory. Nothing is sent
  over the network, and the report loads no external resources.
- **Disposable copies:** the original workbook is only hashed. All Excel work
  happens on a copy that is discarded.
- **Before sharing a report**, check that the structural identifiers in it are
  fine to share.

## Adding VBA instrumentation

See [docs/INSTRUMENTATION.md](docs/INSTRUMENTATION.md).

## Limitations

- **Per-formula runtime:** unavailable. Excel exposes no per-formula timer.
- **Threads:** thread placement and per-thread timing are unavailable.
  `Application.MultiThreadedCalculation` settings are recorded.
- **Calc chain:** the cost of Excel's internal calc-chain reordering is not
  separable from the calculation that contains it.
- **Screen updating:** off during profiling, so phase 4 of Excel's calculation
  (updating visible windows) is excluded.
- **VBA procedures:** only procedures you instrument yourself are timed. There
  is no automatic hook into arbitrary VBA.
- **Undetectable functions:** `Application.Volatile` inside UDFs is not
  detected (the VBA project is not parsed), and database functions that
  reference PivotTables are not detected. Conditional-format and
  data-validation formulas are not inspected.
- **Data tables:** they recalculate single-threaded, once per substitution
  (Microsoft). XLSprint counts them but does not time each substitution.
- **Clock alignment:** VBA and host spans are compared only to within the
  measured clock-probe uncertainty.
- **Spilled dynamic arrays:** a range step is not extended to cover them.
  Calculating the anchor cell recalculates the spill.
- **Fingerprints:** in hashed mode they are HMAC-SHA256 with a per-run salt,
  so they can't be compared across runs. Use `--names clear` if you need
  cross-run comparison.

## Validation status

See [docs/VALIDATION.md](docs/VALIDATION.md). The tests use synthetic fixtures
only. As of 0.1.1 a synthetic smoke run has passed on Windows 11 with
Excel 16 (64-bit). The full Windows integration checklist is not done, and no
real-workbook timings exist yet.

## License

MIT, see [LICENSE](LICENSE). `MicroTimer` is adapted from Microsoft's MIT-licensed sample code; see
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
