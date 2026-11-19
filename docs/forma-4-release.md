# Forma 4 evidence and approval increment

## Changes

- Ranked calculations retain a versioned, session-local record ledger with exact cells, order IDs when a single supported ID header exists, underlying numeric values, displayed values and number formats when available, grouping policy, excluded records and reasons, and a snapshot signature. Rows reconcile to deterministic sums. Per-report selection and paged source rows keep the pane compact.
- Follow-ups resolve category names and aggregate amounts against the ledger. Exact labels remain separate. Case-insensitive, whitespace-normalized grouping requires explicit opt-in. Source reads verify prior evidence before using it.
- A category fix can propose and preview one literal case or whitespace change. Apply revalidates the target, checks the original value and formatting, and verifies the result. Undo rechecks the changed target before restoring it. The scoped operation excludes protected, merged, spilling, validated, hidden, or table targets. Worksheet navigation changes selection only.

## Privacy and limits

Ledger records and signatures are kept in memory. They clear when the pane closes and are excluded from model conversation history. Users who consent to hosted AI can still send prompts, bounded selected-cell samples and workbook summaries to Render and Groq through the existing chat path. Render holds the provider key and task-pane access password. The shared password is not per-user identity; the broker's concurrency control and size checks do not replace tenant-specific authorization, quotas or redaction. Hosting logs record HTTP diagnostics, not workbook payloads. Local evidence tools require no hosted AI.

Ranks include hidden and filtered rows within the requested range. Exact grouping counts repeated records separately. Format observations describe presentation, and do not imply currency conversions. The read is bounded by the existing row, cell and text budgets. Excel tables can be read by selecting their cells; this increment does not identify the native table name. Rechecking a source verifies the fixed original address and the evidence used in the total, and does not track records appended outside that address.

## Deployment and compatibility

The draft PR is an isolated increment based on the existing Forma development branch. Production and the working staging service keep their current releases pending approval. Use the same Python server, Dockerfile, Groq configuration and staging manifest ID after an approved branch promotion. The existing XML references the staging service; it does not package the hosted app. Use the prior staging XML for the current release. Verify the hosted health endpoints, asset versions, official manifest validation, and test on a real workbook before promoting the branch.

Reading and evidence require Excel's existing Office.js capabilities. Category writes require ExcelApi 1.13 and follow the add-in's existing write safeguards. The in-pane reversal is session-only. Native Excel undo for Office.js-driven changes is not promised.

## Backlog

1. Verify on an actual Sales Audit workbook in Excel for the web and supported Windows and Mac builds. Verify that appended rows are detected for named-table evidence.
2. Extend versioned evidence to all numerical tools and distinguish checked numerical results from cloud-generated prose for every analytical intent. Add read/write race tests and stronger access controls.
3. Extend bounded normalization to reviewed multi-row edits and explicit survivor rules. Add validated report export, joins, period analysis and chart previews from evidence.
4. Expand locale-aware cleaning, formula diagnostics, edit-preview coverage and cross-platform accessibility tests. More advanced reporting, workflow execution, broad performance diagnostics and mobile testing remain out of scope for this increment.

## Verification

41 JavaScript unit and adapter tests, 13 Python tests including HTTPS, and 17 Chromium task-pane tests passed. Chromium used an Excel API test double and synthetic workbook data. The full browser sequence verified 4,780 from E2 and E3, the alternate 6,905 normalized total, selecting and restoring the source location, preview before writes, verified apply and undo, and rejection after source changes. The existing staging XML passed local schema validation. Live Excel, live Groq for this branch, and deployment smoke tests remain pending.
