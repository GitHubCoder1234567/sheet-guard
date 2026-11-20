# Forma 5.0 architecture and implementation plan

Baseline is the existing reliability foundation, commit 223e914cd830 and draft PR #5. The architecture inventory and prior failure evidence are in [the reliability audit](reliability-audit.md). This increment extends that source branch so the deployed staging and production branches remain available for comparison.

| Area | Current implementation and risk | Next change and exit evidence |
| --- | --- | --- |
| Workbook and resource access | Office.js uses bounded sync batches and ID-based resolution. A sampled workbook audit cannot establish whole-workbook coverage. Access failures and host API differences can prevent a complete result. | Expand object and dataset inventory and handle ambiguity before choosing a range. Test native tables, stale objects, protection, and resource errors on real hosts. |
| Data model and discovery | Typed cells preserve raw values, reported type, display format and source address. Discovery identifies some contiguous regions. Format-derived date and unit semantics and multi-level headers are incomplete. | Distinguish value from presentation before adding date arithmetic, unit policies or cross-sheet inference. |
| Planning and analytics | Simple questions route locally. Hosted orchestration accepts one plan and at most one follow-up calculation. Basic grouped aggregation cannot compose user filters without losing record identity if the input grid is compacted. | Validate an explicit read → filter → group → rank plan. Preserve original row offsets and revalidate filter cells before showing later source evidence. Extend to safe joins and other statistics incrementally. |
| Sources and conversation | Ranking ledger is session-local and source rechecks use the original address or table name. Text history is bounded. References to a filtered result require the predicate as well as the observed records. | Keep the predicate and its observations together with the aggregate. Broader context, sorting and multi-result references need explicit state and replay rules. |
| Groq and Python API | Server-side key, structured response validation, categorized failures and bounded retries already exist. Earlier live synthetic evidence identified invalid tool sequences. It does not establish the cause of every historical 502. | Keep local verified quantities authoritative. Validate provider contracts and increasingly rich plans separately against the configured model. |
| Formulas, quality, and reports | Existing formula references, audit observations, charts, pivots and comparison remain available. Findings and previews have limited coverage. | Add native dependency evidence and validated joins, then expand previews without moving writes outside the approval pipeline. |
| Editing and undo | Existing scoped previews, stale checks and supported undo are tested through doubles. Multi-object writes can fail partway. | Add fault injection around synchronization boundaries before expanding chart, pivot or dashboard writes. Verify approval, post-write state and undo on real hosts. |
| Privacy, UI, and releases | The shared access password and consent gate secure current hosted paths within their documented scope. Local plans can avoid workbook disclosure. Current manifest and Render services are release-controlled. | Keep formulas and raw records local for deterministic plans. Preserve existing task-pane design and keyboard-usable controls. Expand native corpus, run live tests, and report release evidence per candidate. |

## Implementation checklist

- [x] Reuse the initial repository audit and assess the dependencies and risks for this increment.
- [x] Execute typed conditional ranking end to end with exact source cells and stale-source detection.
- [ ] Extend workbook discovery and metadata for tables, names, filters, hidden rows and relationships on supported hosts.
- [ ] Extend composable analytics to joins, period comparisons, statistics and additional evidence-backed follow-ups.
- [ ] Expand Groq plan validation and run the configured model against synthetic inputs.
- [ ] Expand formula, data-quality, preview and undo coverage without weakening approval safeguards.
- [x] Generate starter and holdout packs with independently checked native worksheet scenarios.
- [ ] Complete native corpus coverage including protected, hidden, filtered and large workbooks.
- [ ] Pass real-host Excel acceptance and review deployment evidence. See [the real Excel validation guide](real-excel-validation.md).
