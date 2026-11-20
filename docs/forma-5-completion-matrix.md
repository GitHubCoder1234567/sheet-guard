# Forma 5.0 completion matrix

This tracks the 31 sections of the Forma 5.0 specification against the staged review stack. **I** means implemented for the stated scope. **P** means partial. **N** means not started or no implementation evidence. **U** marks unit, synthetic, or local service tests. **M** marks integration with an Excel or provider double. **X** requires real Excel evidence for the exact candidate build. **D** means documentation inspected. Status applies to individual requirements in each row. An implemented component does not imply the section is complete. All real Excel entries are **unverified** for this stack. Exporting and reopening an XLSX fixture does not qualify as X.

| § | Requirement | Implementation and evidence | Remaining |
| --- | --- | --- | --- |
| 1 | Workbook-agnostic intelligence | **P U M** Local assessments, rankings and relationship profiles bind fields at runtime across unrelated layouts. | Open-ended reasoning and general workbook transformation are incomplete. |
| 2 | Engineering principles | **I U M** Allowlisted actions, explicit consent and approval, no generated code execution, bounded reads, separate provider and workbook errors, and replayed ranking evidence. **P U M** Provenance coverage, precision, unsupported-case detection, privacy, and approval enforcement across features. | Whole-app invariants, partial-edit incident tests, and live integrity checks. |
| 3 | Existing-code audit and risk | **I D** Architecture, Office integration, provider, audit, tools, UI, security and release reviewed. The earlier invalid Groq tool sequence is documented. **P D** Historical incident attribution. | Correlated evidence for other historical 502 and missing-resource incidents. |
| 4 | Workbook and dataset discovery | **I U M** Sheet enumeration, active identity, visibility, protection status where reported, bounded current samples, native tables, workbook-range name listing, and candidate headers and regions. **P U M** Used range, chart and pivot metadata, role inference, duplicate headers, totals, sparse data, ambiguity and multiple visible regions. | **N** Workbook-wide relationship inference, multi-level headers, merged/validation/filtered layout inventory, and noncontiguous record discovery. |
| 5 | Excel access | **I U M** Central host error translation, worksheet ID/name checks, selection, explicit table and workbook name targets, bounded data and formula reads, navigation, supported edits and charts. **P U M** Host capability tests, stale/protection handling and typed failure outcomes. | **N** Uniform operation-ID result envelopes across all adapter methods, same-operation concurrency guarantees, and live host coverage. |
| 6 | Shared data model | **I U M** Raw/displayed value, Excel type, formula when requested, format, coordinates, and typed blank/error semantics. **P U M** Shared snapshots for all analytical paths and recognizing explicit currency codes. | **N** Reliable locale-aware conversion, units, time/date semantics, workbook identity beyond the session, and complete dataset identity. |
| 7 | General tool engine | **I U M** Enumerate/inspect, read, formula and duplicate checks, filtering/grouping/ranking, statistics, compare, chart/edit proposals, approved applies and scoped undo. **P U M** Quality validation, reconciliation and structured error/evidence contracts. | **N** General typed sort and join tools, materialized results and composable provenance across every tool. |
| 8 | Query planning | **I U M** Conditional rankings and multiple column statistics compose target resolution, bounded reads, validated operations and evidence. **P U M** Existing hosted allowlisted two-step planning and completion checks. | **P U M** Bounded calculation graph with shared intermediate sums, two-input ratios and percentages. General relational operators, flexible composition and verified period-change explanations remain open. |
| 9 | Multi-turn conversation | **I U M** Session-local ranking follow-up, evidence replay, refinement and scoped edit history. **P U M** Hosted short history and ambiguous-reference handling. | **N** Reusable multi-result filters, periods and sorting across turns. |
| 10 | Deterministic analytics | **I U M** Sum, average, numeric and non-empty counts, min/max/median, distinct count, grouped rankings and AND-filtered totals. **P U M** Cross-sheet comparison, key reconciliation and pivot-style summaries. | **I U M** Ratio and percentage of two column totals with aligned numeric rows. **I U M** Follow-up shares of the grand total reuse a freshly checked filtered or unfiltered grouping ledger. Source controls expose the exact rows and predicates for each numerator and the denominator. **N** General grouped comparisons across measures, period and row-wise comparisons, variance/deviation, running and moving statistics, period comparison, robust outlier and general sort operations. |
| 11 | Calculation provenance | **I U M** Exact group, measure, predicate cells, repeated identifiers when available, exclusion counts and stale replay for rankings. **I U M** Composed statistics replay with exact contributing cells. **P U M** Legacy single-statistic and cross-sheet evidence coverage. | **N** Complete contracts for edits, joins, scenarios, units, arbitrary follow-up and every analytical result. |
| 12 | Groq integration | **I U M** Existing configured model path, strict structured outputs, bounded retries, provider and network error classification, privacy-safe logs, local verified results after 502. **P U M** Model compatibility. | Live configured-model tests for the exact candidate, richer plan validation and incident review. |
| 13 | Formula intelligence | **I U M** Selected formula display, references, current result and Excel errors. **P U M** Formula pattern checks and safe eligible repairs. | **N** Comprehensive dependency graphs, intended-reference inference, and efficiency analysis. |
| 14 | Data quality | **I U M** Existing duplicates, blanks, certain inconsistent labels, format and Excel formula errors. **P U M** Confidence, typed impact, missing and invalid keys, scoped corrections. | **N** Comprehensive outliers, unit/date semantics, near-duplicates, validation violations and broken join checks. |
| 15 | Cross-sheet intelligence | **I U M** Explicit comparisons, exact typed key profiling, missing and duplicated keys, cardinality and expansion warnings with sample cells. | **N** Materialized joins/anti-joins, composite and candidate key inference, lookups and combined summaries. |
| 16 | Professional analytics | **P U M** Reusable fields, profiles, rankings, grouping and comparisons cover several domains. | **N** General trends, segmentation beyond exact groups, scenario analysis, and complete KPI and distribution tooling. |
| 17 | Visualization | **P U M** Chart, pivot and dashboard proposals, native chart write support, selected explanation and preview. | **N** General suitability across requested types, exact native preview fidelity, and full evidence integration. |
| 18 | Safe edits | **I U M** Existing explicit approval, stale validation, post-write verification and supported restricted undo. **P U M** Multiple edit types and previews. | **N** Transactional recovery and comprehensive preview/undo for multi-object operations. |
| 19 | Large workbooks | **I U M** Batching, cancellation, progress, payload caps, and evidence pagination. **P U** In-memory benchmark from 10 through 50,000 records. | Native host timings, 100,000+ row full coverage, peak memory and resumable reads. |
| 20 | Self-diagnostics | **I U M** Automatic cell read, Settings report, range types/formulas, synthetic calculations/provenance, hosted health, optional synthetic Groq probe. **P U M** Object checks and capability reporting. | Verified access to names/tables in real host, and preview/write tests on disposable workbooks. |
| 21 | Automated testing | **I U M** Unit, browser, Python HTTPS, property, fault and regression coverage. **P U M** Multi-tool and security integration. | **N** Complete native fault matrix, partial-write interruption tests, broad security and prompt-injection coverage. |
| 22 | Unfamiliar corpus | **I U** Fifty logical corpus/holdout cases and eight native worksheet scenarios with independent expected results. **P U** Domain and shape diversity. | Native feature fixtures and all 30 requested families at release coverage, real Excel results and corpus-wide safety, performance and provenance metrics. |
| 23 | Real Excel | **I D** Repeatable host checklist, disposable workbooks and build/diagnostic evidence schema. **P U M** Levels 1 and 2. | **N** Levels 3 and 4 for this branch. No connected Excel session was available. |
| 24 | Release gate | **I U D** Automated suite, evidence gate rejecting missing host results, documented matrix and release stages. **P D** Release evidence. | Real-host report and review for security, corruption and approval risks before promotion. |
| 25 | User experience | **P M** Existing brand, narrow pane, native controls, source detail, diagnostics and progressive disclosure. | Accessibility audit across hosts, richer contextual suggestions and preview clarity. |
| 26 | Security/privacy | **I U M** Server-side keys, existing shared access, request limits, consent, sanitization and data minimization. **P U M** Plan allowlists and prompt-data handling. | Tenant-grade authorization/rate limits, full redaction policy and security coverage, independent audit. |
| 27 | Platform compatibility | **I U M D** Hosted XML checks, Office.js runtime capability checks, and host-specific checklist. **P U M** Compatibility fallbacks. | Live results for web, Windows and macOS and documented mobile capability evidence. |
| 28 | Implementation roadmap | **I D** Staged engineering plan and reviewable commits. **P U M D** Core software phases are underway. | Complete the ordered backlog and obtain live-host and release evidence. |
| 29 | Required deliverables | See itemized deliverable list below. | Every section marked partial or not started remains an acceptance dependency. |
| 30 | Final acceptance challenge | **P U M** Unrelated holdout data, independent totals and cells, and safe-decline and mock editing scenarios. | **N** Full blind challenge through real Excel with end-to-end safety and failure metrics. |
| 31 | Final execution | **I D** Audited baseline, active branches, test runs, draft PRs and release boundaries. **P** Systematic implementation continues. | Essential criteria in sections 4–24 and 30 remain open. |

## Section 29 deliverables

| # | Deliverable | Status and evidence |
| --- | --- | --- |
| 1 | Architecture audit | **I D** Reliability audit and roadmap |
| 2 | Risk assessment | **I D** Architecture risks and operational release controls |
| 3 | Centralized Excel adapter | **P U M** Shared resolution and reads. Uniform envelopes remain pending |
| 4 | Universal discovery | **P U M** Fresh bounded discovery and explicit objects. Complex layout inference remains pending |
| 5 | Shared data model | **P U M** Typed cells and sources. Locale and units remain pending |
| 6 | General tools | **P U M** Statistics plus two-input comparisons. General composition and joins remain pending |
| 7 | Intelligent planner | **P U M** Conditional analytical flow, validated scalar graph and limited hosted planner |
| 8 | Deterministic engine | **P U M** Basic statistics, ratios of column totals and ranked, filtered sums |
| 9 | Provenance | **P U M** Ranking ledger, independently replayed ratio operands and sample key evidence |
| 10 | Conversational memory | **P U M** Session-local evidence and bounded history |
| 11 | Groq integration | **P U M** Provider validation and retries. Live candidate probe remains pending |
| 12 | Formula tools | **P U M** Basic references, patterns and errors |
| 13 | Quality tools | **P U M** Existing audit and explicit quality checks |
| 14 | Cross-sheet analysis | **P U M** Compare and key profiling |
| 15 | Visualization recommendations | **P U M** Existing chart/report proposals |
| 16 | Safe editing | **P U M** Scoped approval, verification and undo pipeline |
| 17 | Advanced previews | **P M** Selected operations and layouts |
| 18 | Supported undo | **P U M** Scope-limited reversal. Native checks remain pending |
| 19 | Self-diagnostic module | **P U M** Read-only checks and safe export |
| 20 | Synthetic generator | **P U** Two native packs and logical corpus. Full native coverage remains pending |
| 21 | Automated framework | **I U M** JavaScript, Python and Chromium suite |
| 22 | Property tests | **I U** Reordering, renaming and derived independent totals |
| 23 | Fault injection | **P U M** Missing/partial read, invalid provider output and stale evidence. Partial writes remain pending |
| 24 | Permanent regressions | **I U M** Sales Audit, stale source, typed filters and historical parser defects |
| 25 | Real Excel workflow | **I D** Manual procedures and evidence gate |
| 26 | Compatibility matrix | **P D** Hosts listed. X evidence pending for all |
| 27 | Performance benchmarks | **P U** In-memory baselines. Host time and peak memory pending |
| 28 | Security review | **P U M D** Existing audit and targeted tests. Broader review pending |
| 29 | Documentation | **I D** Audit, roadmaps, release and host guides |
| 30 | Deployment and rollback instructions | **I D** Candidate, promotion, rollback checks |
| 31 | Reviewable PRs | **I D** Stacked draft branches with independent diffs |
| 32 | Release-readiness report | **P D** Automated status only. Live release evidence pending |
| 33 | Prioritized limitations | **I D** Current section matrix and ordered roadmap |

## Integrating the review stack

PR #6 targets the reliability foundation branch from PR #5. PR #7 targets the PR #6 branch. GitHub comparison confirmed that PR #7 contains one additional commit above the head of PR #6 and is zero commits behind it as reviewed. The relationship module is independently testable, but its app integration, test changes and host gate share files and ancestry with PR #6.

The integration branch descends from #5 `223e914`, #6 `ff48b5f`, #7 `d3a0e54`, #8 `31b050e` and #9 `b5502d8` in that order. Review each diff against its parent. Assemble on a development integration branch. If using squash or rebasing, replay each child’s own delta onto the new parent and check that the resulting PR diff shows only its intended files. After assembling the stack, run the full suite and verify the published candidate commit in isolated Excel and Groq tests. A staging or production promotion remains a separate approval.

## Core requirement detail

| § | Requirement or closely related requirements | Status |
| --- | --- | --- |
| 4 | Enumerate sheets, active identity, hidden state, used range and native tables | **I U M** for the bounded Office.js paths |
| 4 | Protection, named ranges, charts and supported PivotTables | **P U M** Active-sheet counts and protection, bounded workbook names, explicit range reads |
| 4 | Merged regions, filters, data validation and workbook-wide limits | **N** as discovery inventories. Some edit safeguards already check these properties |
| 4 | Header discovery, multiple tables, blank separators and repeated headers | **I U M** for conservative contiguous cases |
| 4 | Multi-level headers, summary/total/footers, sparse or noncontiguous data and hidden/filtered rows | **P U M** Some explicit totals and ambiguous layouts are recognized. Complex regions and per-row visibility remain pending |
| 4 | Identifiers, dimensions, measures, formulas and column confidence | **P U M** Header, type, distinctness, formula, error and format observations only |
| 4 | Dates, times, units, status, lookup keys, adjacent relationships and suggested operations | **P U M** Explicit currency formats, role hints and chosen-key profiling. General inference remains pending |
| 5 | Worksheet, range, named object, table and selected-cell access | **I U M** Shared adapter and bounded source reads |
| 5 | Stale or removed resources, protection, hidden state, unsupported and partial data | **I U M** for tested isolated failure cases. Concurrent edits and platform coverage remain partial |
| 5 | Operation ID, returned and resolved resources, execution status and safe diagnostic envelopes for all operations | **P U M** Common classified errors. Uniform envelopes and operation IDs remain pending |
| 8 | Recognize intent, bind dataset and fields, validate allowed plan | **P U M** Local typed patterns and existing provider schema |
| 8 | Execute dependent operations, verify results, combine outputs and render | **I U M** for targeted ranking and same-column statistics. Dependent totals use a bounded graph with validated references and operand types. General relational composition remains pending |
| 8 | Reuse verified results across turns and explain period changes | **P U M** Ranking follow-ups, percentage shares of the prior grouped total and statistic evidence replay. General multi-result memory, period logic and causal explanations remain pending |
| 11 | Exact source, contributors, exclusions, operation, filters and stale checks | **I U M** for rankings and composed statistics within their documented scope |
| 11 | Units, original formula dependencies, efficient provenance across every tool | **P U M** Explicit format codes and local storage. Full unit and formula provenance remains pending |
| 12 | Current model configuration, request schema, retries, rate limits and sanitized error distinctions | **I U M** within the existing Groq contract. Candidate live model and broader fault evidence remain pending |
| 18 | Validation, before/after preview, approval, verification and reversal for supported changes | **P U M** Scoped formatting/formula/label workflows and selected visual actions. Native behavior, partial-apply recovery and broader previews remain pending |
| 20 | Auto Office initialization and actual cell read | **I U M** |
| 20 | Detailed workbook, range, calculation, provenance, backend and optional Groq status | **I U M** for bounded read-only exercises |
| 20 | Name resolution, preview and edit capability | **P U M** Capability metadata and explicit untested write status. Real host checks remain pending |
| 22 | Logical fixtures and holdout for sales, transactions, finance, education, inventory, budgets, workforce, project/fulfillment, formulas, messy data and layout variants | **P U** Synthetic matrix and independent oracles exist. Dataset names alone do not establish domain coverage |
| 22 | Native fixtures for work orders, electives, warehouse, survey, dispatch, archive, unusual names and lab | **I U** Eight authored cases with XLSX round-trip checks |
| 22 | Full native corpus for complex financial/healthcare, multi-sheet joins, hidden/filtered/protected layouts, invalid dates and 100,000+ rows | **N** for complete native fixture and acceptance coverage. Some logical variants are already tested |
| 23–24 | Genuine Excel host reports and authorization to promote | **N** Candidate release has no X evidence or promotion approval |
