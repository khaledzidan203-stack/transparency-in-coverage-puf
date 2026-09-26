# Validation evidence summary

## Fresh runtime DAX reconciliation

The fresh post-fix workbook `.local-review/09_DAX_RUNTIME_RECONCILIATION.xlsx` was inspected read-only for the final public-release update on 2026-09-26. All detail rows in the four reconciliation sheets have PASS status. The workbook remains ignored and is not committed. The runtime run was already completed before this documentation update; it was not executed again here.

SHA-256: `5beaef1fe85d47fbbc0d133f736a983d3eac06fd8f92f9a7830b0e3d3b1b6ee7`

| Summary item | Result |
|---|---:|
| Validation status | PASS |
| Core measure checks | 32 |
| Reason checks | 10 |
| Availability checks | 18 |
| DQ context checks | 3 |
| Total failures | 0 |

Power BI and project source were read-only during runtime reconciliation, as reported for the completed run.

## Page 10

PASS: the supplied corrected screenshot displays ISSUER=5, PLAN=33 and Total=38. Static comparison with the previous commit confirms that the only DAX changes are the two `KEEPFILTERS` wrappers inside `Current Context DQ Exception Count`. No other measure changed. The previous issue is resolved.

## Final repository checks

See [release checklist](GITHUB_RELEASE_CHECKLIST.md) for compilation, structural validation, link, security and preservation results. Script 44 is retained as a historical targeted semantic-model patch and was not executed during this update.
