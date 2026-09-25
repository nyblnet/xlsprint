# Validation status

This file separates what has been validated, and how, from what has not. Every
timing that appears in the test suite or in `xlsprint selftest` output is
**synthetic**: it was generated, not measured on any workbook.

## 1. Synthetic validation (done, runs anywhere)

`python -m pytest` runs only on generated fixtures: workbooks built by
`xlsprint/synthbook.py` from generated numbers, traces built by
`xlsprint/synth.py` or inline in the tests, and fake VBA event files.

| Area | What is checked |
|---|---|
| Trace writer and validator | Balanced and nested begin/end, parent/depth/pass consistency, monotonic clocks, duplicate ids, truncation and drops, footer counts, the attribute allow-list and value pattern (rejects `=`, `"`, newlines; accepts Unicode sheet names), expected-instrumentation checks, and rule 8 (each trace-on pass has full and recalc). Each rule has a failing case. |
| Analysis | Interval union with no double counting of nested or overlapping spans, self time, median and IQR summaries, VBA→host offset alignment, drill-down, hotspots, phases, and overhead resolvability. |
| Formula inspection | Relative-R1C1 normalisation and grouping, literal stripping, area compaction, volatile/UDF/single-threaded detection (including the CELL and ADDRESS argument rules), prefixes, names, and a privacy test showing that no sentinel values or formula text reach `formulas.json`. |
| Report | Every section present; synthetic and invalid banners; HTML escaping; no external URLs; every number carries a class; unknown model fields never rendered; refuses an invalid trace unless `--allow-invalid`; integration with the real analysis. |
| Runner (pure parts) | Plan building and encoding, path safety (UNC, mapped drives, same folder), hash check, merging VBA events into validated traces, footer and instrumentation fail-closed checks, and redaction consistency with `formulas.json`. |
| VBA module (static only) | Structure lint (Option Explicit, PtrSafe declares, balanced blocks, labels, no dialogs), allow-list parity with `trace.py`, and a Python port of the JSON escaper round-tripped through `json.loads`. |

`xlsprint selftest --out DIR` runs the offline pipeline end to end on
synthetic data and labels the output SYNTHETIC.

## 2. Windows Excel integration (partial)

### Recorded synthetic Windows smoke

On 2026-09-25, the VBA module and COM runner completed trace-off and trace-on
passes against a generated synthetic workbook in desktop Excel. The VM ran
Windows 11 ARM64 with Excel 16.0 build 20326.0, 64-bit, and eight calculation
threads. Both traces validated, the report was written, there were no run
warnings, and the dedicated Excel process closed successfully. This exercises
the basic COM/VBA instrumentation path on a synthetic fixture; it does not
complete the integration checklist below or establish real-workbook results.

No real workbook, formula inventory, traces, report, or timings are included
in this repository.

### Remaining Windows integration checklist

Use a synthetic or disposable non-confidential workbook for the remaining
cases. `xlsprint selftest` writes `synthetic-workbook.xlsx`.

1. Install with `pip install -e .[windows]`, and turn on *Trust access to the
   VBA project object model*.
2. Run `xlsprint profile synthetic-workbook.xlsx --out C:\xlsprint-out\t1 --repeats 3 --names clear`.
   - Expect exit 0, `trace-on.jsonl` and `trace-off.jsonl` both VALID,
     `report.html` written, and the work copy deleted.
3. **Trust access off.** Rerun with the setting off. It must fail closed with
   the instructions message and leave no Excel process behind (check Task
   Manager).
4. **Import.** The module imports as `XLSprintTimer`, and `XSP_Version`
   returns `xlsprint-vba/1`. This covers both 32-bit and 64-bit Office, if
   available.
5. **Clock.** In `run.json`, the clock-probe uncertainty is small (tens of µs
   at most), and `drift_ns` is small between the first and last probe.
6. **Plausibility.**
   - Full calc ≥ recalc.
   - Sheet times are of the same order as recalc.
   - A `--range` over the period-to-date `SUM($A$1:$A1)` block is slower than
     the running-total block, as in Microsoft's example.
7. **Macro.** Add the shim from `docs/INSTRUMENTATION.md` to a test macro and
   run with `--macro`. The nested spans appear.
   - Without the shim's begin/end, the wrapper span alone should still
     satisfy `--macro`.
   - A mismatched `XSP_End` must make the run fail.
8. **Overhead.** The overhead section shows on/off medians. A small workbook
   will probably show "not resolvable".
9. **Original untouched.** The original file's hash and modification time are
   unchanged.
10. **Timeout.** Use a `--macro` that calls `MsgBox` and run with
    `--timeout 20`. The run must fail with "timed out in profile_trace_on
    pass 1", and no `EXCEL.EXE` may remain.
11. **Clean exit.** After a normal run, no `EXCEL.EXE` remains, and
    `run.json` warnings don't mention the terminate fallback.
12. **Plan file.** Use a sheet with a non-ASCII name (e.g. `Modèle`). It
    must be timed, which shows the UTF-16LE plan file round-trips. Try once
    with FileSystemObject blocked, to exercise the fallback reader.
13. **Arrays.** Put a `--range` over part of a CSE array. The step succeeds
    (no 1004), and its `cells` and `areas` attributes cover the whole array
    exactly once.
14. **Macro errors.** A `--macro` that raises an unhandled error is recorded
    as a step with status `error`, and doesn't hang. A macro that runs `End`
    makes the run fail, not hang.
15. **Things the author could not confirm:**
    - pywin32 named arguments on `Workbooks.Open` and `Close`;
    - imported VBA running in an unsaved `.xlsx` under `AutomationSecurity=1`;
    - `&H` hex literals in `Case` lists;
    - the unqualified `FreeFile`, `LOF`, and `Get #` calls in the fallback
      reader;
    - `Hwnd` → PID lookup (without a PID, timeouts cannot kill Excel;
      `run.json` warns);
    - the FileSystemObject write path, and its fallback on non-ANSI output
      paths.

Record the Excel version, build, bitness, Windows version and outcome of each
step here, and mark real-workbook timings as such.

### Known risks to watch

- A user macro that shows a dialog blocks until `--timeout` expires. Excel
  is then terminated and the run fails. Keep macros non-interactive.
- Array extension relies on openpyxl's view of CSE arrays. Dynamic-array
  spill ranges are not extended.
- The ranges in the drill-down are bounded blocks chosen by the tool. A true
  bottleneck may straddle blocks, so narrow down with `--range`.
