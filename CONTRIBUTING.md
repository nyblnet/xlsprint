# Contributing

## Ground rules

- **The contract:** [docs/DESIGN.md](docs/DESIGN.md) defines the trace
  schema, measurement classes, and module APIs. A change to it goes in the
  same change as the code that needs it.
- **Fixtures:** only synthetic ones. Never commit workbooks, traces, or
  reports made from real data. `.gitignore` blocks `*.xls*` and output
  directories.
- **No values or formula text in outputs.** Any new trace attribute has to be
  added to the allow-list in both `trace.py` and `XLSprintTimer.bas`. The
  parity test fails otherwise.
- **Label every new number** in the report as `measured`, `derived`, `static`,
  or `n/a`. Anything derived shows its formula.
- **Dependencies:** stdlib and openpyxl for the core. New dependencies need an
  MIT-compatible license, recorded in `THIRD_PARTY_NOTICES.md`.

## Verification before a change is merged

1. `python -m pytest`: the whole suite passes.
2. `xlsprint selftest --out <tmp>`: both traces are VALID and the report
   renders.
3. When a change touches `runner.py` or `XLSprintTimer.bas`, run the Windows
   integration checklist in [docs/VALIDATION.md](docs/VALIDATION.md) and
   record the Excel build and result in the PR.
4. Get a review from someone who didn't write the change. Reviewers check
   privacy (no values or formula text), fail-closed behaviour, and measurement
   labelling.
