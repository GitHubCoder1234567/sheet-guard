# Forma reliability foundation release evidence

**Status: Automated testing complete.** The candidate passed local verification at levels 1 and 2. Level 3 requires Excel with Office.js. Level 4 requires evidence across Excel for the web, Windows and Mac. The existing staging service remains on fa1f913. The candidate has no staged deployment or live provider verification.

| Capability | Automated | Live Excel | Limits |
| --- | --- | --- | --- |
| Workbook access and extraction | Passed | Pending | Bounded contiguous reads and a conservative discovery heuristic. Native tables and named ranges exercised through doubles. Direct chat lookup by named-range label is unavailable. |
| Calculations and provenance | Passed | Pending | Single-column statistics and grouped totals use typed values and source coordinates. Recognized currency codes can flag mixed formats. Hidden and filtered rows are included. Source rows represent a captured snapshot. |
| AI orchestration | Passed | Pending | Local ranking and simulated provider failure preserve verified results. Live Groq behavior for this candidate is pending. |
| Formula analysis | Passed | Pending | Selected-cell expression and reference inspection. Cross-sheet dependency tracing remains limited. |
| Data quality | Passed | Pending | Workbook scan is bounded and its diagnoses depend on available inputs. |
| Editing and undo | Passed | Pending | Existing preview and scoped verification exercised by doubles. Multi-object partial writes and full rollback remain unverified. |
| Security and privacy | Passed | Pending | Diagnostics omit workbook values, paths and credentials. The synthetic probe is authenticated using existing service access controls. Broader authorization and live workbook integration require separate review. |
| Performance | Passed for measured cases | Pending | In-memory aggregation of 50,000 synthetic records took 34.44 ms in one local run with a 34.58 MiB heap difference. The 100,000-row case exceeded current limits. These figures exclude Office and network time and do not measure peak heap. |
| Compatibility | Chromium and manifest checks passed | Pending | Chromium ran through a simulated Office bridge. See the host matrix in the live validation guide. |

The verification suite passed 51 JavaScript unit and adapter tests, 18 Python tests including HTTPS, and 19 Chromium tests. No automated failures or skips were recorded. The 50 logical corpus and holdout cases produced 40 matches against an independent integer oracle and 10 expected safe declines. The corpus varies layout, names and values. Its protection, filter and cross-sheet metadata do not prove native Excel behavior. The seeded property test checks 90 reordered data sets. Exercised adapter faults include missing resources, stale object codes, invalid named-range ownership and truncated rows.

The Sales Audit query returned a ranked table and exact contributing cells with an Excel double. The same browser fixture covered explicit approval, verified undo and stale-evidence rejection. The supported release gates and setup instructions are in [live Excel verification](real-excel-validation.md). The implementation plan and remaining work are in the [architecture audit](reliability-audit.md).
