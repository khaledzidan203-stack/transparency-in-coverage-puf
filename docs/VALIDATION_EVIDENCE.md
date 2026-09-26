# Validation evidence summary

Release inspection date: 2026-09-26. Historical workbook summaries below were read from existing local evidence, not rerun. Temporary Excel files remain outside Git. SHA-256 identifies the inspected evidence without exposing local paths.

## Historical runtime DAX evidence

Workbook: `09_DAX_RUNTIME_RECONCILIATION.xlsx`

SHA-256: `27b79cd6ac55ed352182f5ba6e5239e0e48f4bbdc759ab1e2d13ed88e9a2ff98`

| Summary item | Recorded result |
|---|---|
| ValidationStatus | PASS |
| CoreMeasureChecks | 32 |
| ReasonChecks | 10 |
| AvailabilityChecks | 18 |
| DQContextChecks | 3 |
| TotalFailures | 0 |

These are historical runtime results. No fresh runtime DAX execution was performed during finalization. Later visual filter behavior is not certified by this workbook. See [known limitations](KNOWN_LIMITATIONS.md).

## Fresh release checks

See [release checklist](GITHUB_RELEASE_CHECKLIST.md) for the newly executed KPI, semantic-model, PBIR, compilation, link, security and preservation checks. The pre-DAX validator initially failed two stale expectations (no measures table and zero measures); only the validator expectations were updated to the saved release inventory.
