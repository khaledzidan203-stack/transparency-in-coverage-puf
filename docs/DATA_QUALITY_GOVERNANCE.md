# Data quality governance

Policy: preserve, calculate, flag and disclose. The [data contract](DATA_CONTRACT.md) and [KPI contract](KPI_CONTRACT.md) remain authoritative for analytical meaning.

The exception register contains 38 rows: 36 `OPEN_REVIEW` (3 issuer and 33 plan denied-greater-than-received conditions) and 2 known-source rows for issuer 97725. Delta Dental's 1,203 filed and 16,518 overturned internal appeals reconcile to the source's 1,373.07% but remain a disclosed source anomaly.

Availability statuses distinguish numeric availability, unavailable data, small-cell suppression, plan-type exclusions, new-entity exclusions and missing URLs. A reported zero stays numeric zero; a suppression marker never becomes zero. Source issues are not repaired by silently clipping or imputing values.

The 161 resubmissions-greater-than-denied observations are diagnostics. The source records events and does not establish one resubmission per denied claim. Report intensity as events per 100 denied claims.

The DQ register intentionally has no physical semantic-model relationships. Context is applied by measures. The previous page-10 scope-chart issue is resolved: `KEEPFILTERS` intersects the existing Scope context instead of replacing it. The corrected screenshot shows ISSUER=5, PLAN=33 and Total=38. The raw register remains inspectable and unchanged.
