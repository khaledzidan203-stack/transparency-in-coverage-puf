---
document_id: TIC-KPI-CONTRACT-001
title: Transparency in Coverage PUF — KPI Contract
version: "1.0"
status: APPROVED_FOR_VALIDATION
puf_year: 2026
experience_year: 2024
depends_on:
  - docs/DATA_CONTRACT.md
  - data/canonical/*
---

# Transparency in Coverage PUF — KPI Contract

## 1. Purpose

This contract defines the analytical measures that may be published from the governed canonical Transparency in Coverage PUF model.

The contract prevents:

- averaging rates that must be calculated as ratios of totals;
- mixing plan-grain and issuer-grain metrics;
- treating suppressed/unavailable values as zero;
- combining partially available in-network and out-of-network data into misleading totals;
- interpreting claim resubmission counts as unique-claim probabilities;
- treating appeals as a direct subset of claims denied without evidence;
- presenting self-reported disclosure metrics as plan or issuer quality scores.

All dashboard, SQL, Python, Excel, and application calculations must conform to this contract.

---

# 2. Analytical Roles by Grain

## 2.1 Issuer grain — primary market-level claims and appeals view

Use `FactIssuerTransparency` for:

- issuer-level claims received;
- issuer-level claims denied;
- issuer-level claim resubmissions;
- internal appeals filed and overturned;
- external appeals filed and overturned;
- aggregate claims/denial/resubmission analysis by State, Issuer, and Exchange Type.

Business grain:

```text
Experience Year × State × Issuer
```

## 2.2 Plan grain — plan drill-down, denial reasons, and enrollment view

Use `FactPlanTransparency` for:

- plan-level claims received;
- plan-level claims denied;
- plan-level claim resubmissions;
- plan-level denial reasons;
- average monthly enrollment;
- average monthly disenrollment;
- Plan Type / Offering Type / Metal Level analysis.

Business grain:

```text
Experience Year × Plan
```

Plan-level facts may not be summed together with issuer-level facts in the same numerator or denominator.

---

# 3. Availability Eligibility Rule

Every metric calculation must evaluate source availability.

A numeric metric is eligible only when:

```text
AvailabilityStatus = AVAILABLE
```

The following are not zero:

```text
NOT_AVAILABLE
SUPPRESSED_SMALL_CELL
NOT_REQUIRED_PLAN_TYPE
NOT_APPLICABLE_NEW_ENTITY
SOURCE_MISSING_UNEXPECTED
INVALID_NUMERIC_SOURCE
```

For a ratio KPI:

```text
eligible entity
= numerator metric AVAILABLE
AND denominator metric AVAILABLE
AND denominator > 0
```

The numerator and denominator must be aggregated over the exact same eligible entity set.

---

# 4. Core Volume KPIs

## KPI-001 — Issuer Claims Received — In Network

**Entity level:** Issuer  
**Source:** `FactIssuerTransparency.ClaimsReceivedInNetwork`  
**Aggregation:** SUM over rows where the metric is AVAILABLE  
**Format:** whole number  
**Meaning:** reported in-network claim volume for the selected eligible issuer population.

Do not imply that the total represents all Exchange claims when availability is incomplete.

---

## KPI-002 — Issuer Claims Received — Out of Network

**Entity level:** Issuer  
**Source:** `FactIssuerTransparency.ClaimsReceivedOutOfNetwork`  
**Aggregation:** SUM over rows where the metric is AVAILABLE  
**Format:** whole number

---

## KPI-003 — Issuer Claims Denied — In Network

**Entity level:** Issuer  
**Source:** `FactIssuerTransparency.ClaimsDeniedInNetwork`  
**Aggregation:** SUM over rows where the metric is AVAILABLE  
**Format:** whole number

---

## KPI-004 — Issuer Claims Denied — Out of Network

**Entity level:** Issuer  
**Source:** `FactIssuerTransparency.ClaimsDeniedOutOfNetwork`  
**Aggregation:** SUM over rows where the metric is AVAILABLE  
**Format:** whole number

---

## KPI-005 — Comparable Total Claims Received

This is an all-network KPI and requires both network components to be available for each issuer.

Eligible issuer condition:

```text
ClaimsReceivedInNetwork AVAILABLE
AND ClaimsReceivedOutOfNetwork AVAILABLE
```

Formula:

```text
SUM(
    ClaimsReceivedInNetwork
  + ClaimsReceivedOutOfNetwork
)
over eligible issuers only
```

**Entity level:** Issuer  
**Format:** whole number

Do not construct this KPI by independently summing the two network measures over different issuer populations.

---

## KPI-006 — Comparable Total Claims Denied

Eligible issuer condition:

```text
ClaimsDeniedInNetwork AVAILABLE
AND ClaimsDeniedOutOfNetwork AVAILABLE
```

Formula:

```text
SUM(
    ClaimsDeniedInNetwork
  + ClaimsDeniedOutOfNetwork
)
over eligible issuers only
```

---

# 5. Denial KPIs

## KPI-007 — In-Network Denial Rate

Eligible issuer condition:

```text
ClaimsDeniedInNetwork AVAILABLE
AND ClaimsReceivedInNetwork AVAILABLE
AND ClaimsReceivedInNetwork > 0
```

Formula:

```text
SUM(ClaimsDeniedInNetwork)
/
SUM(ClaimsReceivedInNetwork)
```

where both SUMs use the exact same eligible issuer rows.

**Format:** percentage  
**Aggregation rule:** ratio of totals  
**Forbidden calculation:** average of issuer-level denial percentages

---

## KPI-008 — Out-of-Network Denial Rate

Eligible issuer condition:

```text
ClaimsDeniedOutOfNetwork AVAILABLE
AND ClaimsReceivedOutOfNetwork AVAILABLE
AND ClaimsReceivedOutOfNetwork > 0
```

Formula:

```text
SUM(ClaimsDeniedOutOfNetwork)
/
SUM(ClaimsReceivedOutOfNetwork)
```

**Format:** percentage  
**Aggregation rule:** ratio of totals

---

## KPI-009 — Comparable Overall Denial Rate

Eligible issuer condition:

```text
all four metrics AVAILABLE:
- ClaimsReceivedInNetwork
- ClaimsReceivedOutOfNetwork
- ClaimsDeniedInNetwork
- ClaimsDeniedOutOfNetwork

AND
ClaimsReceivedInNetwork + ClaimsReceivedOutOfNetwork > 0
```

Formula:

```text
SUM(ClaimsDeniedInNetwork + ClaimsDeniedOutOfNetwork)
/
SUM(ClaimsReceivedInNetwork + ClaimsReceivedOutOfNetwork)
```

using the same comparable issuer population.

Do not calculate:

```text
(KPI-007 + KPI-008) / 2
```

---

# 6. Resubmission KPIs

CMS reports the number of claim resubmissions for claims previously denied. The source definition does not establish a one-resubmission-per-denied-claim constraint.

Therefore the project must not label the following ratio as a probability or bounded "rate."

## KPI-010 — In-Network Resubmissions per 100 Denied Claims

Eligible issuer condition:

```text
ClaimsResubmittedInNetwork AVAILABLE
AND ClaimsDeniedInNetwork AVAILABLE
AND ClaimsDeniedInNetwork > 0
```

Formula:

```text
SUM(ClaimsResubmittedInNetwork)
/
SUM(ClaimsDeniedInNetwork)
* 100
```

**Display unit:** resubmission events per 100 denied claims

Values above 100 are permitted and are not automatically a DQ failure.

---

## KPI-011 — Out-of-Network Resubmissions per 100 Denied Claims

Formula:

```text
SUM(ClaimsResubmittedOutOfNetwork)
/
SUM(ClaimsDeniedOutOfNetwork)
* 100
```

with the same availability and denominator rules.

---

## KPI-012 — Comparable Total Resubmission Events

Eligible issuer condition:

```text
ClaimsResubmittedInNetwork AVAILABLE
AND ClaimsResubmittedOutOfNetwork AVAILABLE
```

Formula:

```text
SUM(
    ClaimsResubmittedInNetwork
  + ClaimsResubmittedOutOfNetwork
)
```

over the common eligible issuer population.

---

# 7. Appeals KPIs

Appeal counts are issuer-level measures.

Do not divide appeals by claims denied as a default KPI. CMS defines appeals as requests for review of adverse determinations, and the source does not establish a one-to-one denominator relationship to the claim-denial metrics.

## KPI-013 — Internal Appeals Filed

```text
SUM(InternalAppealsFiled)
```

where AVAILABLE.

---

## KPI-014 — Internal Appeals Overturned

```text
SUM(InternalAppealsOverturned)
```

where AVAILABLE.

---

## KPI-015 — Internal Appeal Overturn Rate

Eligible issuer condition:

```text
InternalAppealsFiled AVAILABLE
AND InternalAppealsOverturned AVAILABLE
AND InternalAppealsFiled > 0
```

Formula:

```text
SUM(InternalAppealsOverturned)
/
SUM(InternalAppealsFiled)
```

using the same eligible issuer set.

**Aggregation rule:** ratio of totals  
**Forbidden:** average of `InternalAppealsOverturnPctPublished`

The published percentage field remains available for source reconciliation and issuer-detail display.

Known source anomaly `Issuer_ID = 97725` is preserved. When included, the calculated rate may exceed 100%; dashboards must display a DQ/source-exception indicator rather than silently clipping the value.

---

## KPI-016 — External Appeals Filed

```text
SUM(ExternalAppealsFiled)
```

where AVAILABLE.

---

## KPI-017 — External Appeals Overturned

```text
SUM(ExternalAppealsOverturned)
```

where AVAILABLE.

---

## KPI-018 — External Appeal Overturn Rate

Eligible issuer condition:

```text
ExternalAppealsFiled AVAILABLE
AND ExternalAppealsOverturned AVAILABLE
AND ExternalAppealsFiled > 0
```

Formula:

```text
SUM(ExternalAppealsOverturned)
/
SUM(ExternalAppealsFiled)
```

using the same eligible issuer set.

---

# 8. Plan-Level Claims KPIs

The plan-level equivalents of claims received, denied, denial rates, and resubmission intensity follow the same ratio-of-totals and common-population rules as issuer-level KPIs.

They must be clearly named with `Plan` in technical definitions and must never be added to issuer-level figures.

Recommended plan-level analytical measures:

```text
Plan Claims Received — In Network
Plan Claims Received — Out of Network
Plan Claims Denied — In Network
Plan Claims Denied — Out of Network
Plan In-Network Denial Rate
Plan Out-of-Network Denial Rate
Plan Overall Denial Rate — Comparable Population
Plan Resubmissions per 100 Denied Claims — In Network
Plan Resubmissions per 100 Denied Claims — Out of Network
```

---

# 9. Denial Reason KPIs

The plan fact contains ten reported denial-reason metrics.

They are:

1. Referral / prior authorization required
2. Out-of-network provider
3. Services excluded
4. Not medically necessary — excluding behavioral health
5. Not medically necessary — behavioral health only
6. Benefit limit reached
7. Member not covered
8. Investigational / experimental / cosmetic
9. Administrative reason
10. Other

## KPI-019 — Denial Reason Count

For a selected denial reason:

```text
SUM(reason count)
```

over plans where that reason is AVAILABLE.

Use this for reason-specific volume analysis.

---

## KPI-020 — Comparable Denial Reason Composition %

This KPI is permitted only on a plan population where **all ten denial-reason metrics are AVAILABLE**.

For reason `r`:

```text
SUM(Reason_r)
/
SUM(All 10 reported denial-reason counts)
```

over the exact same comparable plans.

**Interpretation:**

> share of the reported denial-reason counts within the fully comparable population.

Do not label this automatically as:

```text
share of all denied claims
```

until validation proves that the reason fields form the required mutually exclusive and exhaustive decomposition.

---

# 10. Enrollment KPIs

Enrollment metrics are plan-level average monthly measures.

## KPI-021 — Reported Average Monthly Enrollment

```text
SUM(AverageMonthlyEnrollment)
```

over plans where the metric is AVAILABLE.

This is a **reported available-population total**, not an estimate for suppressed or unavailable plans.

Always pair with availability coverage.

---

## KPI-022 — Reported Average Monthly Disenrollment

```text
SUM(AverageMonthlyDisenrollment)
```

over plans where the metric is AVAILABLE.

Always pair with availability coverage.

---

## KPI-023 — Comparable Disenrollment-to-Enrollment Ratio

Eligible plan condition:

```text
AverageMonthlyEnrollment AVAILABLE
AND AverageMonthlyDisenrollment AVAILABLE
AND AverageMonthlyEnrollment > 0
```

Formula:

```text
SUM(AverageMonthlyDisenrollment)
/
SUM(AverageMonthlyEnrollment)
```

using the same eligible plan set.

**Interpretation:** ratio of reported average monthly disenrollment to reported average monthly enrollment for the comparable plan population.

Do not label this as annual churn, retention rate, or member attrition probability.

---

# 11. Data Availability KPIs

Data availability is a first-class analytical subject in this project.

For a selected metric:

## KPI-024 — Available Entity Count

```text
COUNTROWS(metric availability rows where StatusCode = AVAILABLE)
```

---

## KPI-025 — Applicable Entity Count

Applicable denominator:

```text
AVAILABLE
+ NOT_AVAILABLE
+ SUPPRESSED_SMALL_CELL
+ SOURCE_MISSING_UNEXPECTED
+ INVALID_NUMERIC_SOURCE
```

Exclude structurally non-applicable statuses:

```text
NOT_REQUIRED_PLAN_TYPE
NOT_APPLICABLE_NEW_ENTITY
```

---

## KPI-026 — Data Availability Rate

```text
Available Entity Count
/
Applicable Entity Count
```

If Applicable Entity Count = 0, return blank.

---

## KPI-027 — Suppression Rate

```text
SUPPRESSED_SMALL_CELL count
/
Applicable Entity Count
```

---

## KPI-028 — Unavailable Rate

```text
NOT_AVAILABLE count
/
Applicable Entity Count
```

---

## KPI-029 — Structural Non-Applicability Count

```text
NOT_REQUIRED_PLAN_TYPE
+ NOT_APPLICABLE_NEW_ENTITY
```

This is disclosed separately and is not included in the applicability denominator.

---

# 12. Data Quality KPIs

## KPI-030 — Open DQ Exception Count

```text
COUNTROWS(
    DQExceptionRegister
    where ResolutionState = OPEN_REVIEW
)
```

Current validated build expectation:

```text
36 open-review rows
```

---

## KPI-031 — Known Source Exception Row Count

```text
COUNTROWS(
    DQExceptionRegister
    where IsKnownSourceException = TRUE
)
```

Current validated build expectation:

```text
2 rows
```

---

## KPI-032 — Known Source Exception Entity Count

```text
DISTINCTCOUNT(EntityKey)
where IsKnownSourceException = TRUE
```

Current validated build expectation:

```text
1 issuer
```

---

## KPI-033 — Diagnostic Observation Count

Count of analytical observations intentionally classified as informational rather than DQ failures.

Current validated build expectation:

```text
161
```

The initial observation type is:

```text
RESUBMISSION_EVENTS_GT_DENIED_CLAIMS
```

---

# 13. Population / Coverage Companion Measures

Every headline ratio should be capable of displaying its eligible population.

Minimum companion measures:

```text
Eligible Issuer Count
Eligible Plan Count
Excluded Due to Suppression Count
Excluded Due to Not Available Count
Excluded Due to Structural Non-Applicability Count
```

For volume totals built only from available values, display a data-availability measure beside the KPI when the user could otherwise interpret the value as a complete market total.

---

# 14. DQ-Aware KPI Behavior

A DQ exception does not automatically remove an entity from analysis.

Default treatment:

```text
PRESERVE
CALCULATE ACCORDING TO SOURCE VALUES
FLAG
DISCLOSE
```

Exclusion is allowed only in a separately labeled sensitivity view, for example:

```text
Internal Appeal Overturn Rate — Reported
Internal Appeal Overturn Rate — Excluding Known Source Exceptions
```

The default reported KPI must not silently exclude known source values.

---

# 15. Aggregation Rules

## Additive counts

Counts may be summed only:

- at their approved grain;
- across mutually exclusive entities;
- without mixing issuer and plan facts.

## Ratios

All rates and ratios use:

```text
ratio of aggregated numerator and denominator
```

over one common eligible population.

Never use:

```text
AVERAGE(row-level rate)
```

unless a future KPI contract explicitly defines an unweighted average as a separate analytical statistic.

## Percentages supplied by source

Published appeal percentages are:

- non-additive;
- not averaged for aggregate KPIs;
- retained for source reconciliation and record-level presentation.

---

# 16. Comparison Rules

Valid descriptive comparisons include:

```text
State vs State
Issuer vs Issuer
Plan vs Plan
QHP vs SADP vs SHOP
Plan Type
Metal Level
Exchange Type
In Network vs Out of Network
```

Only compare entities for a KPI when the KPI eligibility rule is met.

For comparative tables, show or make accessible:

```text
numerator
denominator
eligible entity count
availability rate
DQ flag
```

---

# 17. Interpretation Guardrails

The project may describe:

```text
higher/lower observed volume
higher/lower observed ratio
concentration
distribution
outlier
source exception
data availability
suppression
reported denial-reason mix
```

The project must not infer from these metrics alone:

```text
best insurer
worst insurer
best plan
worst plan
quality score
clinical quality
patient outcomes
profitability
fraud
wrongful denial
causal effect
```

Claims, denials, appeals, and enrollment disclosures are analytical observations, not direct quality ratings.

---

# 18. Validation Requirements Before BI Publication

The KPI implementation checkpoint must validate:

1. issuer and plan ratios use common eligible populations;
2. no special-status value becomes zero;
3. aggregate rates are ratios of totals;
4. published appeal percentages reconcile to source counts within approved tolerance;
5. plan-level and issuer-level claims are not double counted;
6. total-network KPIs require both network components;
7. resubmission intensity may exceed 100 without automatic DQ failure;
8. enrollment totals are labeled as reported/available-population values;
9. denial-reason composition uses a fully comparable plan population;
10. known source anomalies remain traceable;
11. eligible entity counts reconcile to availability facts;
12. cross-grain plan-to-issuer reconciliation is tested only where all child plan values required for the metric are AVAILABLE.

---

# 19. KPI Validation Gate

The next implementation must produce one review artifact:

```text
Desktop/
└── Transparency_PUF_Review/
    └── 03_KPI_VALIDATION_REPORT.xlsx
```

No temporary KPI-validation outputs are authorized inside the project.

Final implementation code:

```text
src/validation/03_validate_kpis.py
```

The script must read only final canonical project files and must not modify them.

---

# 20. Evidence State

```text
CANONICAL DATASET        = BUILT
CANONICAL BUILD          = PASS WITH PRESERVED SOURCE/DQ EXCEPTIONS
KPI CONTRACT             = APPROVED FOR VALIDATION
KPI VALIDATION           = NOT YET RUN
EDA                      = NOT YET STARTED
SQL                      = NOT YET AUTHORIZED
POWER BI                 = NOT YET AUTHORIZED
```
