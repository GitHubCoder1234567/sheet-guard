# Forma

Forma is an Excel task pane for inspecting workbook data, reviewing proposed edits and making approved changes through Office.js. Its Render service serves the task pane and brokers optional GroqCloud requests. Local analysis stays in Excel.

## Try the workspace

- Ask **Review my spreadsheet** or **Clean this dataset** for bounded workbook findings. Select a recommendation to see the affected cells and a preview.
- Select a rectangular dataset with headers and ask **Total Revenue by Month**, substituting your column headers.
- Ask **Compare datasets**, capture two selections, then choose their shared identifier. Repeated keys are excluded from row matching.
- Select a clean range and ask **Create a PivotTable for total Revenue by Month**. Review the proposed report before creating a new worksheet.
- Select one numeric column without headers and ask **Prevent negative values**. Forma previews the rule and refuses to replace an existing validation rule.
- Open **Menu** for the audit, recommendations, visualisations, history and settings. Cloud chat requires consent and a configured service.

Local comparison and totals read up to 50,000 rows and 160,000 cells per selection in bounded Office.js calls. Workbook-wide audit and event monitoring use smaller bounded samples. PivotTable previews use the regular, smaller editable-range limit. Cloud chat uses limited context and the backend keeps credentials server-side.

## Develop and test

Start changes on a development branch and open a pull request for review. Verify a staging release before production. The production manifest retains the original application identifier and Render endpoint.

- Run `node --test tests/agent.test.mjs tests/engine.test.mjs tests/excel.test.mjs`.
- Run `python3 -m unittest discover -s tests`. Set `SHEETGUARD_TEST_HTTPS=1` to include loopback TLS.
- Install Playwright and Chromium, then run `node --test tests/ui.test.mjs`. Set `CHROMIUM_EXECUTABLE` to use an installed Chrome and `SHEETGUARD_TEST_ORIGIN` to test staging assets.
- Run `python3 manifest.py --check SheetGuardAI.xml` and Microsoft's Office manifest validator before distributing an updated XML.

## Groq on Render

Configure the existing Forma staging service through its Environment page. The existing start command is `python3 server.py`.

| Variable | Value |
| --- | --- |
| `AI_PROVIDER` | `groq` |
| `AI_MODEL` | `openai/gpt-oss-20b` |
| `GROQ_API_KEY` | GroqCloud API key stored as a Render environment secret |
| `SHEETGUARD_MODE` | `production` |
| `SHEETGUARD_ACCESS_PASSWORD` | A separate password for authorised task-pane users, stored as a Render environment secret |
| `HOST` | `0.0.0.0` |

Redeploy after saving. In Forma Settings, enable AI data sharing and enter the deployment access password. Users enter the access password, while the provider key remains on Render. With no provider key, existing local features remain available. The legacy OpenAI integration is available only if an administrator explicitly sets `AI_PROVIDER=openai` and configures its key.

Groq strict structured output returns an allowlisted plan. Forma dispatches the selected workbook tool locally. Workbook inspection can make one bounded follow-up call. Exact totals, counts, duplicate groups, and column statistics are computed locally and shown with references. Chart, style and dashboard requests create proposals through the existing approval flow. [Groq structured-output documentation](https://console.groq.com/docs/structured-outputs)

Aggregation and duplicate analysis cover the selected contiguous data range (or a single cell's surrounding region when supported) within the documented size limits. Worksheet inspection and formula evidence use bounded samples. A keyed comparison needs two explicit captures. Larger, ambiguous, hidden, or unsupported ranges can require a more specific selection.

## Forma 4 evidence increment

Ask **Which product category has the highest sales?**, then ask **Where did the total come from?** or select a category in the ranked table. Forma keeps source rows in the current pane session, re-reads the source before answering a follow-up, and navigates using selection only. Ask **Group categories case-insensitively** to explicitly combine case and whitespace variants. The original exact analysis keeps variants separate.

On supported Excel hosts, **Fix the inconsistent category labels** can prepare a single literal-cell normalization from the latest verified report. Review the preview and approve it explicitly. The verified edit can be reversed in History. Calculations include rows inside the requested range regardless of filters or hiding.

The fixture is `tests/fixtures/sales-audit.json`. Include `tests/evidence.test.mjs` in the Node test command. See [architecture audit](docs/forma-4-audit.md) and [the release and deployment notes](docs/forma-4-release.md).

### Reliability validation

Use **Settings → Diagnostics** to verify workbook access and export a safe report. The optional live Groq check sends a synthetic greeting through the server. Diagnostics omit workbook names, ranges, cell contents and passwords. The `/health` response includes the hosted build identifier when Render supplies it.

Run `sh tools/verify.sh` with Node, Python 3, OpenSSL and Playwright. Set `CHROMIUM_EXECUTABLE` if using an existing Chromium installation. Run `node tools/benchmark.mjs` for the in-memory synthetic benchmark. See the [architecture audit](docs/reliability-audit.md), [release evidence and known limits](docs/reliability-foundation-release.md), and [Excel host verification and rollback procedure](docs/real-excel-validation.md).

### Conditional rankings

Select a dataset with its header row and ask `Rank "Team" by total "Hours" where "Status" = "Open" and "Priority" >= 2`. Use exact headers. The result shows ranked totals and lets you inspect contributing measure and filter cells. Text comparisons are case-sensitive. Numeric filters compare underlying Excel numbers. The plan supports up to four conditions joined by `and` and runs entirely inside the add-in.

See the [Forma 5.0 roadmap](docs/forma-5-roadmap.md) and [verification report](docs/forma-5-conditional-analysis.md).

Append `in table "Actuals"`, `in named range "CurrentData"`, or `in worksheet "Operations"` to target a source without changing the selection. A worksheet request requires a single unambiguous dataset. Use **Audit → Discover worksheet datasets** for a fresh profile of the active sheet. See the [31-section completion matrix](docs/forma-5-completion-matrix.md).

Use `Calculate sum and average of "Hours" in table "Schedule"` to run up to four supported statistics on the same source. The task pane shows each result and its included and excluded counts. Choose a result’s source button to verify its contributing cells against the workbook.
