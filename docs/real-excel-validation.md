# Forma live Excel verification

The candidate has completed automated testing. Live Excel and staging verification require separate evidence tied to the deployed commit.

| Host | Required for release | Evidence on this branch |
| --- | --- | --- |
| Excel for the web | Yes | Pending |
| Excel for Windows | Yes | Pending |
| Excel for Mac | Yes | Pending |
| Excel for iPad | Exploratory | Pending |

## Set up a candidate

1. Run `sh tools/verify.sh` with Node, Python 3, OpenSSL and Playwright available. For an existing compatible Chromium executable, set `CHROMIUM_EXECUTABLE`. Run `node tools/benchmark.mjs` and retain its output as an in-memory baseline.
2. Review the PR. With approval for an isolated candidate URL, deploy its exact commit with the existing Dockerfile, /health check, Groq model and access controls. Check the service name and commit in Render. Keep the working staging deployment available while testing. Do not change production configuration.
3. Verify /health returns the candidate’s 12-character Git SHA. Generate and validate a candidate XML with `python3 manifest.py --origin https://YOUR-CANDIDATE-HOST --output forma-candidate.xml` and `python3 manifest.py --check forma-candidate.xml`. In Excel for the web use Add-ins → More Add-ins → My Add-ins → Upload My Add-in, when enabled. On desktop use your organization’s approved sideloading or add-in catalog workflow. Use a test profile or remove any conflicting sideloaded XML with the same add-in ID before installing the candidate.
4. Use a disposable workbook. Use a worksheet called **Sales Audit** and enter the fixture below in **B4:D8**. Use a fresh copy for destructive or conflicting cases. Inspect Settings → Diagnostics and enable **Include live Groq check**. Run the checks and download the safe report. Confirm its build matches /health.

| Row | B | C | D |
| --- | --- | --- | --- |
| 4 | Product Category | Total Sales | Record ID |
| 5 | East | 30 | R1 |
| 6 | West | 40 | R2 |
| 7 | East | 17 | R3 |
| 8 | West | -4 | R4 |

## Run each scenario

| ID | Procedure and expected observation |
| --- | --- |
| connect | Open the pane and confirm it reports a readable workbook only after Excel returns a cell from the current selection. |
| worksheet-resolution | Ask for the active sheet name, then switch worksheets and ask again. Each reply uses the active sheet. |
| renamed-sheet | Rename Sales Audit. Repeat a queued or prior source request and a new request with the updated name. The app resolves the new name or reports that the earlier target needs refreshing. |
| deleted-sheet | Delete a referenced sheet in a copy and repeat the prior request. The pane explains that the target is missing or stale. |
| table-and-totals | Turn B4:D8 into an Excel table with a native totals row. Select C5 and ask for grouped sales. Lookup uses table data and excludes the displayed totals row. |
| named-range | Define SalesInput as B4:D8 with workbook scope. Select it using Excel’s Name Box and ask for grouped sales in the selection. Confirm the coordinates. Direct chat lookup by named-range label is outside this increment. |
| ambiguous-data | Add another independent table with a blank column between it and the first. A sheet-wide question requests a precise selection instead of combining regions. |
| typed-values | In a copy, replace C8 with text entered as apostrophe-4. Ask for the sum of Total Sales in the selection. Expect 87 with one excluded numeric text entry. |
| sales-ranking | Ask “Analyse the Sales Audit worksheet. Rank every product category by total sales. Show the exact cells used for each calculation.” Expect East 47 from C5 and C7, West 36 from C6 and C8, in that order. |
| column-calculation | Select B4:D8 and ask “Calculate the median of Total Sales in the selection”. Expect 23.5 with C5 through C8 as contributors. |
| source-selection | Open the East evidence, check R1 and R3, select a cited cell, then confirm no data was edited. |
| formula-errors | Add =1/0 to a disposable copy, select it and explicitly ask for formula errors. Expect a report of the error and its coordinate. |
| provider-probe | Enable the optional synthetic check. Confirm Groq’s validated reply passes and that no customer data was used. |
| provider-fallback | In a candidate test session with AI consent enabled, block /api/agent in browser DevTools or use an isolated candidate with the provider unavailable. Ask for the ranking and an explanation. Confirm the ranked totals and exact cells remain visible. |
| report-privacy | Search the exported diagnostics for sheet names, source addresses, values, passwords and prompt text. None should appear. |
| preview-cancel | Request a supported formatting change, inspect Preview and cancel. Confirm the workbook remains unchanged. |
| stale-preview | Start a supported change, edit targeted cells in Excel and then approve. Confirm Forma refuses the stale change. |
| apply-and-undo | Approve a supported single-range format change. Confirm the workbook reflects it and Forma verifies it. Undo using the pane, then confirm the original formatting. |
| protected-sheet | Protect the worksheet before a proposed change. Confirm prohibited changes are blocked and reads report their actual status. |
| large-range | Paste synthetic data at increasing sizes within the supported 50,000-row and cell limits. Record Office read and total response time. Confirm over-limit requests explain the limit. |

| discovery | Open Audit → Discover worksheet datasets. With only the sample table present, confirm it reports B4:D8 as the candidate and Total Sales as numeric. |
| resource-planning | On another sheet, create the named Excel table TestOrders from the Work orders fixture, including its header. Switch to an unrelated worksheet and ask the same conditional ranking with the suffix `in table "TestOrders"`. Check the table owner's cells. Repeat for a workbook-scoped range and an unambiguous worksheet. Remove the original source and confirm the follow-up refuses the old result. |
| formula-inspection | In a copy, enter =C5+C6 and select the formula cell. Ask “Explain this formula.” Verify the expression, current result of 70 and its references. |
| cross-sheet | In Compare, capture the sample on two sheets. Change one amount on the second sheet and compare by Record ID. Expect one changed record. |
| key-relationships | In a disposable workbook, put headers `Order key` and `Hours` in A1:B1 on the first sheet and enter X, X, Y and Left in A2:A5. On a second sheet, put `Dispatch key` and `Status` in A1:B1 and enter X, X, Y and Right in A2:A5. Use Compare datasets to capture both A1:B5 ranges. Under Inspect key relationships select the two key columns. Expect five matching row pairs and a many-to-many warning for X, with A2 and A3 named on both sheets. |
| ai-conversation | With consent enabled and a configured staging password, ask an ordinary question. Confirm the response addresses it. |
| conditional-ranking | Open the Work orders fixture. Select B4:E12 and ask `Rank "Workshop" by total "Hours" where "State" = "Open" and "Priority" >= 2`. Expect Forge 17, Assembly 11, Packing 5. Open Forge and confirm C5, C7 and C11 contributed. Change E5 from 2 to 4 and reopen the old Forge evidence. Expect an explanation that the source changed even though the total would be the same. |
| evidence-followup | Ask for the records behind the East total after the ranking. Verify the two rows and exact source references. |
| quality-audit | In a copy, add a duplicated record and a formula error. Run Audit and verify the observed findings and source area shown. |
| recovery | After a missing-sheet or stale-source error, choose the valid input again. Repeat the read and check that it succeeds. |
| unsupported-api | Exercise an action whose ExcelApi requirement is missing on a supported legacy host or use a host-limited operation such as a noncontiguous selection. Confirm the UI explains the limitation. |

Use sanitized observations. Keep the downloaded diagnostics as the `diagnostics` field of the report. Add a `tester` identifier and a `results` array with an entry shaped like `{"id":"sales-ranking","status":"passed","expected":"47 and 36 with four cells","observed":"47 and 36 with four matching cells"}` for each scenario. Use `failed` or `not_tested` where appropriate. Run `node tools/verify-live-report.mjs BUILD web.json windows.json mac.json`. The script checks the presence and consistency of recorded evidence. It relies on the tester’s observations.

## Promotion and rollback

Review the automated results, diagnostics, workbook outcomes and provider logs together before asking for staging promotion. If authorized to update existing staging, use Render to deploy the exact reviewed SHA and repeat /health, asset and Excel smoke checks. The existing staging XML loads the app at its configured URL. Reopen its task pane after an approved deploy. When using an isolated candidate URL, use its matching manifest.

If a promoted staging deployment regresses, select its prior live deployment in Render and use Rollback, or redeploy that recorded SHA. Confirm /health and test the task pane again. The foundation baseline is `fa1f913da12b6ddf82a7d070878b145d4a7ea5d7`. Keep promotion to production behind explicit approval.
