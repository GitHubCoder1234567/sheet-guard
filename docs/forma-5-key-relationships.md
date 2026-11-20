# Forma 5.0 cross-sheet key relationships

**Release status is Automated testing complete.** This increment adds a read-only relationship check under Compare datasets. The existing shared-header field comparison remains available. Select the key column for each captured range. The headings can differ.

Exact key matching preserves Excel numbers, text and Booleans as distinct types. Text comparisons are case-sensitive and retain whitespace. Empty and whitespace-only keys count as missing. Excel errors and unsupported types count as invalid. The report reconciles matched, unmatched, missing and invalid records with the two captures.

Forma counts the observed cardinality for each shared key and estimates the number of row pairs in an exact inner join. It shows representative source key cells on both sheets and warns when a many-to-many match would multiply rows. The engine counts the pairs without building the Cartesian product. The existing per-range read limits still apply.

## Verification

- 58 JavaScript tests, 18 Python tests including HTTPS and 21 Chromium tests passed. No automated failures or skips were recorded.
- With an Excel API double, differently named key fields produced five matching pairs, and exact key cells appeared from both sheets. The operation made no Office writes or AI requests.
- A synthetic fault and capacity test rejected incomplete data and checked a 20,000 by 12,000 key collision without creating 240 million result rows.
- Real Excel and a candidate deployment remain unverified. See [the live Excel procedure](real-excel-validation.md) and its key-relationship scenario.

This is a profile over the captured ranges. Recapture after workbook changes. Sample evidence is bounded to three examples per relationship type and four cell references per side of each example. Other join types, composite keys, inferred business relationships and materialized join output remain deferred in the [Forma 5.0 roadmap](forma-5-roadmap.md).
