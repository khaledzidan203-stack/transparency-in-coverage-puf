# DAX ↔ KPI Contract Audit

Status: **REVIEWED — NO MODEL CHANGE APPLIED**  
Scope: saved `_Measures.tmdl` versus `KPI_CONTRACT.md` for the current PUF 2026 / Experience 2024 release.

## Purpose

This audit separates two questions:

1. Does the current saved model reproduce the already validated current-release results?
2. Does every DAX eligibility rule exactly encode the broader written KPI contract for future data states?

The first question is already supported by the retained runtime evidence: 63 DAX reconciliation checks passed with zero failures after the approved DQ context fix. The second question requires a stricter semantic comparison and exposes several latent contract differences documented below.

No DAX, TMDL, canonical data, report visual or screenshot is changed by this document.

## Confirmed aligned areas

The saved model follows the contract in these major areas:

- issuer and plan calculations remain separate;
- availability is evaluated before numeric aggregation;
- suppressed/unavailable values are not treated as zero;
- denial rates are ratios of totals, not averages of row-level percentages;
- comparable overall denial rates require all four received/denied network metrics to be available;
- denial-reason composition uses the fully comparable ten-reason plan population;
- resubmissions are expressed as events per 100 denied claims rather than a bounded probability;
- availability metrics use the governed applicable-status set;
- the DQ register remains disconnected and entity context is applied in DAX;
- the approved `KEEPFILTERS` correction preserves DQ scope context.

## Latent semantic differences

### 1. Entity-level positive-denominator rule

The KPI contract requires ratio eligibility to include a positive denominator at entity level, for example:

```text
numerator metric AVAILABLE
AND denominator metric AVAILABLE
AND denominator > 0
```

The Python KPI validator explicitly applies the positive-denominator condition before aggregating the eligible population.

Several saved DAX ratio measures build eligibility from availability only, then call `DIVIDE` on the aggregate denominator. This pattern appears in:

- Issuer In-Network Denial Rate
- Issuer Out-of-Network Denial Rate
- Issuer Comparable Overall Denial Rate
- Issuer Resubmission Events per 100 Denied — In Network
- Issuer Resubmission Events per 100 Denied — Out of Network
- Internal Appeal Overturn Rate
- External Appeal Overturn Rate
- Plan In-Network Denial Rate
- Plan Out-of-Network Denial Rate
- Plan Comparable Overall Denial Rate
- Plan Resubmission Events per 100 Denied — In Network
- Plan Resubmission Events per 100 Denied — Out of Network
- Comparable Disenrollment-to-Enrollment Ratio
- Internal Appeal Overturn Rate — Excluding Known Source Exception

For the validated current dataset, retained runtime reconciliation reports zero failures. Therefore this is treated as a **latent contract-hardening issue**, not evidence that the published current-release figures are wrong.

### 2. Comparable claims totals use a narrower DAX population than KPI-005 / KPI-006

The written contract defines:

- KPI-005 Comparable Total Claims Received: both received metrics available.
- KPI-006 Comparable Total Claims Denied: both denied metrics available.

The saved DAX measures `Issuer Comparable Claims Received - Total` and `Issuer Comparable Claims Denied - Total` both intersect all four received/denied availability sets. That produces a stricter common four-metric population than the individual KPI-005 and KPI-006 definitions.

This may be intentional for dashboard comparability, but the semantic choice should be explicit rather than implicit.

Recommended resolution options:

- keep the stricter four-metric DAX population and amend the contract/display name to state that these are four-metric comparable totals; or
- change the DAX measures to match KPI-005 and KPI-006 exactly, followed by full runtime reconciliation and screenshot review.

No option is applied by this audit.

## Current decision

Do **not** patch DAX solely because a written contract and implementation differ. The current release has validated runtime evidence and preserved report screenshots. Any semantic-model change would require:

1. explicit approval of the intended business definition;
2. targeted DAX change only;
3. fresh Python-to-DAX reconciliation;
4. semantic-model validation;
5. affected visual review;
6. refreshed validation evidence and protected-artifact baseline.

Until then, the saved model remains the release implementation and this audit makes the difference visible instead of silently rewriting history.

## Recommended next decision

The safest next step is to keep the current validated model unchanged and treat the two items above as **documented hardening candidates**. They can be resolved in a separately versioned analytical change if the project is extended to a new PUF year or if exact contract equivalence is required.