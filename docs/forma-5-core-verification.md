# Forma 5.0 core verification

**Release status is Automated testing complete.** The review stack adds workbook target resolution, current discovery profiles and composed local statistics. It has no candidate deployment or real Excel report. The existing manifest and deployments were left in place.

## Workbook access and discovery

Discovery reads its worksheet identity and a bounded sample in the same Office.js run. It reports typed column observations, formula and error counts, explicit currency-format codes and coverage. A sampled or ambiguous layout cannot become a complete analytical dataset. A scoped request can use a native table, a workbook-scoped named range, or a single unambiguous worksheet dataset. Scoped evidence re-resolves the resource before reusing it. Global table resolution uses the documented Table owner and requires ExcelApi 1.4 for the null-object lookup. [Microsoft Table](https://learn.microsoft.com/en-us/javascript/api/excel/excel.table) [TableCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablecollection) [NamedItem](https://learn.microsoft.com/en-us/javascript/api/excel/excel.nameditem)

## Composed read-only analytics

A supported question can calculate up to four statistics from one resolved dataset. Supported operations are sum, average, minimum, maximum, median, numeric count, non-empty count and distinct count. Forma validates the operations and exact column before calculating. Each operation reports included and excluded counts. Its evidence button reopens the source and compares the captured identity, values, types, errors and display formats before recomputing the contributing cells. Each result stays in the task pane and is omitted from hosted conversation history.

The current planner supports specific read-only request patterns. The statistics path uses a secure host with Web Crypto for snapshot verification and preserves exact Excel types. Snapshot comparison does not make an atomic transaction and does not distinguish formula changes that preserve the same observed values and formats. General dependency graphs, arbitrary date/period interpretation, write planning and broad natural-language follow-up remain covered by the [completion matrix](forma-5-completion-matrix.md) and staged roadmap.

## Evidence

| Suite | Results | Evidence boundary |
| --- | --- | --- |
| JavaScript | 63 passed | Unit, seeded cases, source identity, ambiguity and replay tests |
| Python including HTTPS | 18 passed | Local service tests and provider doubles |
| Chromium | 24 passed | Task pane with an Office.js test double |
| Real Excel and hosted candidate | Unverified | Live host report and isolated release candidate required |

The complete automated suite passed without failures or skips. Named-table browser coverage verified that a request can use a different worksheet and that deleting its table prevents a stale source follow-up. Composed statistics coverage verified one source read for several calculations and stale evidence after value changes that preserve the total. Continue with the new entries in the [live Excel checklist](real-excel-validation.md).
