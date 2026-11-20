# Forma 5.0 conditional analysis

**Release status is Automated testing complete.** The candidate builds on the reliability foundation in draft PR #5. Its scope is a read-only conditional ranking plan and eight native worksheet fixture scenarios. The Render builds and manifest configuration were left intact.

## What changed

A request that matches the documented conditional ranking syntax resolves exact headers in the current selection, checks typed predicates, groups the eligible rows and ranks their totals. The plan is validated before data execution. Blank, error, and incompatible predicate values are counted separately from compatible values that do not match. Numeric text is kept as text. The existing ranking evidence now retains the predicate cell coordinates and values for contributing records.

Source follow-ups reopen the original range or native table and verify the observations before displaying records. A relevant predicate value changing makes the prior evidence stale even if the totals remain identical. Filtered results support source follow-ups. They do not produce edit proposals or change category normalization. Recognized filtered questions and their results do not enter later hosted chat history. The add-in's established consent rules still govern ordinary hosted chat.

## Verification and remaining work

| Check | Evidence | Status |
| --- | --- | --- |
| JavaScript unit and adapter tests | 55, including 75 seeded conditional ranking cases across five header sets | Passed |
| Python and local HTTPS | 18, including simulated Groq failures and privacy-safe diagnostics | Passed |
| Chromium task pane | 20 with an Office.js test double, including conditional ranking, privacy, exact filter cells, stale evidence, and existing approval and undo | Passed |
| Native fixture structure and round-trip | Eight worksheet cases across two XLSX packs. The saved files were reopened and each ranking, amount cell and filter cell was checked against the independent record oracle. | Passed |
| Live Excel for the web, Windows and Mac | Requires an authorized test URL and an Excel session | Unverified |
| Live configured Groq and staging candidate | No new hosted deployment | Unverified |

No automated failures or skips were recorded for the final candidate. The established in-memory benchmark and limits remain documented in the reliability foundation report. Real Office.js performance and native edit coverage remain pending. The current plan requires a contiguous selected dataset, uses the stored numeric values, includes hidden and Excel-filtered rows in that selection, and supports up to four conditions joined by AND. Joins, period parsing, OR clauses, broad natural-language plans and the rest of the native corpus are tracked in the [roadmap](forma-5-roadmap.md).

## Reproduce native fixtures

Run `node tools/build-live-fixtures.mjs OUTPUT_DIR [PREVIEW_DIR]` in a workspace where the bundled `@oai/artifact-tool` is available. The script exports starter and holdout XLSX packs and a separate JSON answer key. Every worksheet has its own range and question. Use a disposable copy of each workbook for live tests. Keep the expected answers outside the live workbook. Use the [real Excel procedure](real-excel-validation.md) to capture build, host and observed outcome.
