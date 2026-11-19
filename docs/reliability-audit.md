# Forma reliability engineering audit

## Reviewed baseline

Staging source commit `fa1f913`. Reviewed the README, Dockerfile, add-in XML and manifest generator, Python server and provider tests, all task-pane modules, CSS and support page, and the existing JavaScript, browser and fixture tests. The adjacent parser-hardening candidate belongs to a separate development branch.

| Layer | Current behavior and risk |
| --- | --- |
| Excel access | The adapter batches Office.js calls and validates bounded A1 rectangles. Sheet lookups are repeated at call sites, and host exceptions can escape without an operation-specific explanation. Cached sheet metadata can disagree with live ID/name bindings after a rename or deletion. A reported “resource does not exist” is consistent with such failures, but there is insufficient evidence to attribute that historical request to a specific object. |
| Discovery and shape | The automatic audit samples visible sheets and profiles the first row as a possible header. Table names are recorded, but tables and named ranges are not first-class read targets. A used rectangle can contain multiple unrelated datasets, report titles or total rows. Empty formatted regions, hidden data and noncontiguous selections require different handling. |
| Values and provenance | Selection, positional audit samples, and evidence-backed totals use different snapshots. Rankings have a local row ledger and stale recheck. Other calculations may return only a summary or a source rectangle. Display formats are not units; date serials and date-like text must not be interchanged. |
| AI and 502s | The broker uses strict Groq JSON and an allowlisted local dispatcher. Earlier live synthetic ranking probes reproduced `invalid_tool_sequence` when a tool was accompanied by another intent or finding. The current staged named-sheet ranking path calculates locally and treats the provider as an optional explanation. Other wording can still route through the broker and encounter planning errors. Several classes previously collapsed into an indistinguishable 502. |
| Safety of writes | There is explicit preview, stale snapshot comparison and post-write verification. Formula, number and eligible literal-label operations have scoped reversal. Chart, PivotTable and dashboard operations may partially succeed and have no broad automated rollback. Staleness checks are not a workbook-wide transaction or lock. Some preview paths use approximated visuals. |
| UI and health | A successful Office readiness callback displays a connection indicator, but it does not prove the workbook is readable. There is no systematic workbook self-check or safe diagnostic export on the baseline. |
| Backend and security | Shared access protection, request caps, concurrency limits and HTTPS are present. The shared password is not user-level identity. Model data is consent-gated, but text summaries and bounded samples may contain personal data. Preserve credentials on Render and keep workbook values out of diagnostics. |
| Release and compatibility | Existing Render services reference different source commits while the XML points at hosted URLs. Local tests use Excel doubles and do not prove host-specific behavior. Production promotion needs actual Excel evidence and explicit approval. |

## Engineering plan and checklist

- [x] Review architecture and distinguish confirmed failure paths from hypotheses.
- [x] Harden target resolution, typed host errors and workbook discovery without replacing the write pipeline.
- [x] Add read-only lightweight and comprehensive diagnostics, authenticated provider probe and safe export.
- [x] Unify range and table representations, publish typed result contracts and extend provenance to additional numerical operations.
- [x] Expand calculation and quality coverage with holdout fixtures. Treat unsupported formats and ambiguous data conservatively.
- [ ] Extend tool orchestration and provider fault coverage without giving model-generated plans workbook write authority.
- [ ] Harden approval state, conflict handling and reversible operations. Verify partial failures on actual hosts.
- [x] Establish synthetic corpus, reproducible runtime benchmarks and release gates.
- [ ] Verify Excel Web, Windows and macOS with a registered live-host test harness. Leave test level pending until evidence exists.
- [ ] Validate staging and obtain explicit approval for production.

Priority order is workbook access and diagnostics, consistency of numerical evidence, write safety, then breadth and polish. Keep release status at Development or Automated testing complete until a separate Excel host and staging gate passes.

## Follow-on work

1. Exercise the core Office adapter on real Excel before expanding the write surface. Add fixtures for named tables, filtered ranges, unsupported APIs and edits interrupted after sync.
2. Extend the typed calculation and provenance contract to filters, time-series and cross-sheet joins. Add locale-aware parsing and explicit unit policies, and preserve exclusion reasons in the task pane.
3. Continue provider and orchestration hardening as its own reviewed change. Cover invalid envelopes, recovery limits, tool-plan progress, stage timings and verified answers during upstream failure.
4. Extend staged previews and conflict recovery for charts, pivots and dashboards. Add live evidence of supported write verification and undo.
5. Validate larger native workbooks, platform-specific behavior and comparative response time before promoting a new release.
