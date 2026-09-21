CALL ANALYTICS DASHBOARD - V8 CLIENT MATCH

Purpose
- Reconcile campaign attribution to the supplied Book2.xlsx client source for the 2026-09-18 report.
- Preserve Answered classification: answered rows are counted without customer de-duplication.
- Drop classification: Abandoned/drop only, Connected customer status when available, global customer de-duplication, queue routing, Missed excluded.
- Test calls excluded by exact normalized Customer Name: Vishal Test Call, Nirgun Test Call, Test Call.

Client source findings
- Book2.xlsx Sheet1 contains 309 source rows: 285 Answered + 24 DROP CALL.
- Its CAMP NAME column is authoritative when present.
- For Sr Number 917996165333, all 57 source rows are CAMP NAME = Press 4 NBEMS NEET-PG 2026, including DTMF/Queue values that may show Press 2 or Press 3.
- Therefore this version attributes that shared Sr to Press 4 for the supplied client reconciliation.
- The raw CSV has 389 rows; it contains 80 rows not present in Book2. Five of those are Answered test calls and are excluded by the test-call rule.

Expected reconciliation for the supplied raw CSV
- Answered: 285
- Drop: 24
- Total identified: 309
- Press 4 NBEMS NEET-PG 2026: 48 Answered, 9 Drop, Total 57

Important
- A new localStorage key is used so stale mappings from older dashboard versions do not silently change the result.
- Campaign Setup edits remain supported.
- UI files are kept unchanged from the previous dashboard version.
