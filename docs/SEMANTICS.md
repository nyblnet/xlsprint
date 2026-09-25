# Semantic annotations

XLSprint accepts an optional JSON sidecar with `profile --semantics`. It adds
owner-authored labels and purpose text to defined names, formula groups,
explicit ranges, and VBA procedure spans. It never derives financial meaning
from formula text.

```json
{
  "schema": "xlsprint.semantics/1",
  "workbook_sha256": "<64 lowercase hex characters>",
  "annotations": [
    {
      "id": "net-generation",
      "label": "Net generation",
      "intent": "Convert gross production into saleable MWh after availability and losses.",
      "category": "Operations",
      "source": "Model owner",
      "match": {"defined_name": "NetGeneration"}
    },
    {
      "id": "master-route",
      "label": "Generate and converge the asset model",
      "intent": "Run the master route that imports assumptions, calculates the asset, and writes outputs.",
      "category": "Master model",
      "match": {"span": {"kind": "vba.proc", "name": "Main.zzMain"}}
    },
    {
      "id": "forecast-output-block",
      "label": "Forecast output block",
      "intent": "Calculate the monthly operating forecast output region.",
      "category": "Operating forecast",
      "match": {"range": {"sheet": "Operations", "address": "F12:F183"}}
    }
  ]
}
```

`workbook_sha256` is required for a map generated for a specific model. XLSprint
rejects a hash mismatch before opening the disposable copy in Excel. It is
optional only for reusable procedure-level annotations such as macro names;
the report warns when a map is not workbook-bound.

Each annotation must have a unique lowercase `id`, a `label`, an `intent`, and
exactly one selector:

- `defined_name`: a defined name, or an object with `name` and optional `scope`;
- `range`: exact worksheet name and A1 address. It matches an exact timed range,
  name, or formula-group area when that span exists in the run;
- `formula_group`: exact worksheet and a formula-group area;
- `span`: an XLSprint trace kind plus optional `name`, `sheet`, and `address`.

The report shows the semantic description first and retains the measured
operation, range address, and redacted workbook identifier as technical
details. It distinguishes an annotation that did not resolve in the workbook
from one that resolved but was not timed in this run. A description labels a
measured region; it does not prove that each formula in the region contributed
the reported time. Excel does not expose per-formula runtime during normal
recalculation.

When names are hashed (the default), clear workbook selectors are resolved
while XLSprint has the workbook open for static inspection; persisted
`formulas.json` retains only the same salted hashes used by the trace. The
human-authored labels and intent are emitted because the user explicitly
provided the sidecar. Treat the sidecar and generated report as model metadata
that may reveal business context, and keep them local unless approved for
sharing. No cell values, formula text, or literals are included.
