# Forma 5.0 core verification

**Release status is Automated testing complete.** The review stack adds workbook target resolution, current discovery profiles and composed local statistics. It has no candidate deployment or real Excel report. The existing manifest and deployments were left in place.

## Workbook access and discovery

Discovery reads its worksheet identity and a bounded sample in the same Office.js run. It reports typed column observations, formula and error counts, explicit currency-format codes and coverage. A sampled or ambiguous layout cannot become a complete analytical dataset. A scoped request can use a native table, a workbook-scoped named range, or a single unambiguous worksheet dataset. Scoped evidence re-resolves the resource before reusing it. Global table resolution uses the documented Table owner and requires ExcelApi 1.4 for the null-object lookup. [Microsoft Table](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table) [TableCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablecollection) [NamedItem](https://learn.microsoft.com/en-us/javascript/api/excel/excel.nameditem)

## Composed read-only analytics

A supported question can calculate up to four statistics from one resolved dataset. Supported operations are sum, average, minimum, maximum, median, numeric count, non-empty count and distinct count. Forma validates the operations and exact column before calculating. Each operation reports included and excluded counts. Its evidence button reopens the source and compares the captured identity, values, types, errors and display formats before recomputing the contributing cells. Each result stays in the task pane and is omitted from hosted conversation history.

The current planner supports specific read-only request patterns. The statistics path uses a secure host with Web Crypto for snapshot verification and preserves exact Excel types. Snapshot comparison does not make an atomic transaction and covers only the selected cells. Changes to captured formula expressions invalidate statistics evidence even when the result is identical. General dependency graphs, arbitrary date/period interpretation, write planning and broad natural-language follow-up remain covered by the [completion matrix](forma-5-completion-matrix.md) and staged roadmap.

## Evidence

| Suite | Results | Evidence boundary |
| --- | --- | --- |
| JavaScript | 69 passed | Unit, seeded cases, source identity, ambiguity and replay tests |
| Python including HTTPS | 19 passed | Local service tests and provider doubles |
| Chromium | 26 passed | Task pane with an Office.js test double |
| Real Excel and hosted candidate | Unverified | Live host report and isolated release candidate required |

The complete automated suite passed without failures or skips. Named-table browser coverage verified that a request can use a different worksheet and that deleting its table prevents a stale source follow-up. Composed statistics coverage verified one source read for several calculations and stale evidence after value changes that preserve the total. Continue with the new entries in the [live Excel checklist](real-excel-validation.md).

## Dependent calculations

Both the local request compiler and the consented Groq tool can produce a comparison of two column totals. The graph validates up to 12 read-only operations before reading Excel, checks operand types and dependency cycles, and evaluates shared nodes once. The task pane displays the numerator, denominator, arithmetic, row counts and replayable contributors and exclusions. Evidence is kept on the client. Explicit local plans accept the same selection, table, named-range and worksheet scopes as other statistics. Hosted comparisons use the current selected dataset. A different target requires a fresh selection or an explicit local plan.

A percentage divides two sums over the same numeric rows and multiplies by 100. Zero denominators, unequal contributing rows, differing recognized currency formats and non-finite results produce errors. Excel dates can be numeric serials and unmarked units cannot be inferred. Users must choose compatible measures. All rows in the rectangle, including hidden or filtered rows, are in scope. Separate tests cover 24 generated layouts, both operand cell sets, stale data, the visible calculation, and a paraphrased prompt with a simulated provider.

The next planner increment should support typed row-set and grouped aggregate results with original row identities. Move conditional ranking onto those operators, then implement a two-period grouped comparison with independently validated period bounds and drill-down for both periods. The existing rank engine and its regression cases provide the comparison oracle. Candidate promotion requires a live configured-model probe and real Excel evidence for this commit.

The integration run of the existing in-memory ranking benchmark processed 50,000 records in 39 ms with a 38 MiB heap delta. The 100,000-record case exceeded the configured budget. These measurements exclude Office.js reads, the task pane and provider latency.
