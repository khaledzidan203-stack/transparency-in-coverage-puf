---
document_id: TIC-DATA-CONTRACT-001
title: Transparency in Coverage PUF — Canonical Data Contract
version: "1.0"
status: APPROVED_FOR_IMPLEMENTATION
source_file: Transparency_in_Coverage_PUF.xlsx
puf_year: 2026
experience_year: 2024
---

# Transparency in Coverage PUF — Canonical Data Contract

## 1. Purpose

This contract defines the canonical analytical representation of the CMS Transparency in Coverage PUF 2026 source.

The contract exists to:

- prevent double counting caused by mixed issuer-level and plan-level measures in one flat source;
- preserve source disclosure-status semantics such as `*`, `**`, `***`, `N/A`, and `Missing URL`;
- separate identifiers, dimensions, measures, availability state, and data-quality evidence;
- define stable grain, keys, types, relationships, and validation rules before transformation;
- provide one governed source for Python, SQL, Power BI, Excel, and future analytical applications.

No downstream implementation may override these definitions without an explicit project-contract change.

---

# 2. Evidence Boundary

## 2.1 Source identity

- PUF publication/label year: `2026`
- analytical/experience year represented by the source: `2024`
- analytical source sheets:
  - `Transparency 2026 - Ind QHP`
  - `Transparency 2026 - Ind SADP`
  - `Transparency 2026 - SHOP`
- analytical source rows: `4,956`
- distinct `Plan_ID`: `4,956`
- distinct `Issuer_ID`: `348`
- distinct states: `30`

## 2.2 Source preservation

The workbook in `data/raw/` is immutable.

Transformations must never overwrite the source workbook.

Every canonical output must be reproducible from the immutable RAW workbook and the approved transformation code.

---

# 3. Canonical Architecture

```text
RAW CMS Workbook
    |
    v
Canonical transformation
    |
    +--> DimReportingPeriod
    +--> DimState
    +--> DimIssuer
    +--> DimPlan
    +--> DimMetric
    +--> DimAvailabilityStatus
    |
    +--> FactIssuerTransparency
    +--> FactPlanTransparency
    |
    +--> FactIssuerMetricAvailability
    +--> FactPlanMetricAvailability
    |
    +--> DQExceptionRegister
```

The two analytical facts remain separate because issuer-level and plan-level measures have different grains.

---

# 4. Grain Contracts

## 4.1 DimReportingPeriod

```text
One row = one governed analytical reporting/experience period
```

Current row:

```text
PUF Year        = 2026
Experience Year = 2024
Period Start    = 2024-01-01
Period End      = 2024-12-31
```

Primary key:

```text
ReportingPeriodKey
```

---

## 4.2 DimState

```text
One row = one state represented in the current canonical source
```

Business key:

```text
StateCode
```

Current source validation requires one `Exchange_Type` per state.

If multiple PUF years are later appended and a state's exchange type changes by year, `Exchange_Type` must be moved to a governed State × Period structure rather than silently overwriting history.

---

## 4.3 DimIssuer

```text
One row = one Issuer_ID in the current source
```

Business key:

```text
IssuerID
```

Current-source invariant:

```text
IssuerID -> one StateCode
IssuerID -> one IssuerName
IssuerID -> one ExchangeType
IssuerID -> one New-to-Exchange flag
IssuerID -> one SADP-only flag
IssuerID -> one claims-policy URL
IssuerID -> one rate-review URL
IssuerID -> one financial-information value/status
```

---

## 4.4 DimPlan

```text
One row = one Plan_ID
```

Business key:

```text
PlanID
```

Current-source invariant:

```text
Plan_ID is unique across all three analytical sheets.
Plan_ID embeds the same StateCode as the source State field.
```

---

## 4.5 FactIssuerTransparency

Authoritative business grain:

```text
Experience Year × State × Issuer
```

Current source row count expected:

```text
348
```

Although `Issuer_ID` currently maps to exactly one state, State remains part of the business grain because CMS defines issuer-level claims/appeals data at state level.

Primary uniqueness rule:

```text
ReportingPeriodKey + StateKey + IssuerKey
```

No issuer-level measure may be summed directly from the flat source before deduplication to this grain.

---

## 4.6 FactPlanTransparency

Authoritative business grain:

```text
Experience Year × Plan
```

Current source row count expected:

```text
4,956
```

Primary uniqueness rule:

```text
ReportingPeriodKey + PlanKey
```

The fact must also carry conformed foreign keys:

```text
StateKey
IssuerKey
```

to allow State and Issuer filters to work directly without dimension-to-dimension relationships.

---

## 4.7 Metric Availability Facts

### FactIssuerMetricAvailability

```text
One row = Experience Year × State × Issuer × Issuer Metric
```

Expected current row count:

```text
348 × 12 = 4,176
```

### FactPlanMetricAvailability

```text
One row = Experience Year × Plan × Plan Metric
```

Expected current row count:

```text
4,956 × 18 = 89,208
```

These tables preserve disclosure/availability semantics independently from numeric analytical values.

---

# 5. Canonical Availability Status Contract

| Source value/state | Canonical StatusCode | Meaning |
|---|---|---|
| numeric/text business value present | `AVAILABLE` | Usable source value exists |
| `*` | `NOT_AVAILABLE` | Data not available for issuer/plan |
| `**` | `SUPPRESSED_SMALL_CELL` | Suppressed due to small cell size |
| `***` | `NOT_REQUIRED_PLAN_TYPE` | Data not required due to plan type |
| `N/A` | `NOT_APPLICABLE_NEW_ENTITY` | Not applicable because issuer/plan offering is new to Exchange |
| `Missing URL` | `MISSING_URL` | Financial-information URL unavailable |
| unexpected blank/NULL in a governed required source field | `SOURCE_MISSING_UNEXPECTED` | Missing state not defined by CMS legend |
| non-special non-numeric value in a governed numeric metric | `INVALID_NUMERIC_SOURCE` | Requires DQ review |

Rules:

1. A source status code is never converted to zero.
2. `0` is a valid numeric value and remains `AVAILABLE`.
3. A suppressed/unavailable metric has `MetricValue = NULL`.
4. Source status meaning must remain queryable after transformation.
5. `SOURCE_MISSING_UNEXPECTED` and `INVALID_NUMERIC_SOURCE` create DQ exceptions.

---

# 6. Canonical Data Types

| Semantic class | Canonical logical type | SQL-oriented type |
|---|---|---|
| Issuer ID | text | `varchar(5)` |
| Plan ID | text | `varchar(14)` |
| State | text | `char(2)` |
| categorical label | text | `varchar(...)` |
| Yes/No flag | boolean | `bit` |
| URL | nullable text | `nvarchar(2048)` |
| claim/appeal count | nullable whole number | `bigint` |
| appeal percentage | nullable decimal | `decimal(18,4)` |
| average monthly enrollment/disenrollment | nullable decimal | `decimal(18,4)` |
| year key | integer | `smallint` |
| date | date | `date` |
| availability status | governed enum/text | `varchar(32)` |

Important:

- `Issuer_ID` is an identifier, not a measure, even if Excel stores it numerically.
- `Plan_ID` must remain text.
- IDs must never be summed or averaged.
- percentage values are stored as percentage points exactly as published, not divided by 100 in the canonical layer.
- the source anomaly above 100% is preserved and flagged rather than clipped.

---

# 7. Source-to-Canonical Mapping — All 44 Source Columns

| # | Source column | Canonical owner | Canonical field / treatment | Type |
|---:|---|---|---|---|
| 1 | `Individual/SHOP` | `DimPlan` | `MarketSegment` | text enum |
| 2 | `Exchange_Type` | `DimState` | `ExchangeType` | text enum |
| 3 | `State` | `DimState` + fact FKs | `StateCode` | char(2) |
| 4 | `Issuer_Name` | `DimIssuer` | `IssuerName` | text |
| 5 | `Issuer_ID` | `DimIssuer` + fact FKs | `IssuerID` | varchar(5) |
| 6 | `Is_Issuer_New_to_Exchange?(Yes_or_No)` | `DimIssuer` | `IsIssuerNewToExchange` | boolean |
| 7 | `SADP_Only` | `DimIssuer` | `IsSADPOnlyIssuer` | boolean |
| 8 | `Plan_ID` | `DimPlan` + plan fact FK | `PlanID` | varchar(14) |
| 9 | `Plan_Type` | `DimPlan` | `PlanType` | text enum |
| 10 | `QHP or SADP?` | `DimPlan` | `PlanOfferingType` | text enum |
| 11 | `Metal_Level` | `DimPlan` | `MetalLevel` | text enum |
| 12 | `URL_Claims_Payment_Policies` | `DimIssuer` | `ClaimsPaymentPoliciesURL` | nullable URL |
| 13 | `Issuer_Claims_Received_Out_of_Network` | `FactIssuerTransparency` | `ClaimsReceivedOutOfNetwork` + availability row | bigint nullable |
| 14 | `Issuer_Claims_Received_In_Network` | `FactIssuerTransparency` | `ClaimsReceivedInNetwork` + availability row | bigint nullable |
| 15 | `Issuer_Claims_Denied_Out_of_Network` | `FactIssuerTransparency` | `ClaimsDeniedOutOfNetwork` + availability row | bigint nullable |
| 16 | `Issuer_Claims_Denied_In_Network` | `FactIssuerTransparency` | `ClaimsDeniedInNetwork` + availability row | bigint nullable |
| 17 | `Issuer_Claims_Resubmitted_Out_of_Network` | `FactIssuerTransparency` | `ClaimsResubmittedOutOfNetwork` + availability row | bigint nullable |
| 18 | `Issuer_Claims_Resubmitted_In_Network` | `FactIssuerTransparency` | `ClaimsResubmittedInNetwork` + availability row | bigint nullable |
| 19 | `Issuer_Internal_Appeals_Filed` | `FactIssuerTransparency` | `InternalAppealsFiled` + availability row | bigint nullable |
| 20 | `Issuer_Number_Internal_Appeals_Overturned` | `FactIssuerTransparency` | `InternalAppealsOverturned` + availability row | bigint nullable |
| 21 | `Issuer_Percent_Internal_Appeals_Overturned` | `FactIssuerTransparency` | `InternalAppealsOverturnPctPublished` + availability row | decimal nullable |
| 22 | `Issuer_External_Appeals_Filed` | `FactIssuerTransparency` | `ExternalAppealsFiled` + availability row | bigint nullable |
| 23 | `Issuer_Number_External_Appeals_Overturned` | `FactIssuerTransparency` | `ExternalAppealsOverturned` + availability row | bigint nullable |
| 24 | `Issuer_Percent_External_Appeals_Overturned` | `FactIssuerTransparency` | `ExternalAppealsOverturnPctPublished` + availability row | decimal nullable |
| 25 | `Plan_Number_Claims_Received_Out_of_Network` | `FactPlanTransparency` | `ClaimsReceivedOutOfNetwork` + availability row | bigint nullable |
| 26 | `Plan_Number_Claims_Received_In_Network` | `FactPlanTransparency` | `ClaimsReceivedInNetwork` + availability row | bigint nullable |
| 27 | `Plan_Number_Claims_Denied_Out_of_Network` | `FactPlanTransparency` | `ClaimsDeniedOutOfNetwork` + availability row | bigint nullable |
| 28 | `Plan_Number_Claims_Denied_In_Network` | `FactPlanTransparency` | `ClaimsDeniedInNetwork` + availability row | bigint nullable |
| 29 | `Plan_Number_Claims_Resubmitted_Out_of_Network` | `FactPlanTransparency` | `ClaimsResubmittedOutOfNetwork` + availability row | bigint nullable |
| 30 | `Plan_Number_Claims_Resubmitted_In_Network` | `FactPlanTransparency` | `ClaimsResubmittedInNetwork` + availability row | bigint nullable |
| 31 | `Plan_Number_Claims_Denied_Referral_Required` | `FactPlanTransparency` | `DeniedReferralOrPriorAuthorization` + availability row | bigint nullable |
| 32 | `Plan_Number_Claims_Denied_Due_To_Out_Of_Network` | `FactPlanTransparency` | `DeniedDueToOutOfNetworkProvider` + availability row | bigint nullable |
| 33 | `Plan_Number_Claims_Denied_Services_Excluded` | `FactPlanTransparency` | `DeniedServicesExcluded` + availability row | bigint nullable |
| 34 | `Plan_Number_Claims_Denied_Not_Medically_Necessary_Excluding_Behavioral_Health` | `FactPlanTransparency` | `DeniedNotMedicallyNecessaryExclBH` + availability row | bigint nullable |
| 35 | `Plan_Number_Claims_Denied_Not_Medically_Necessary_Behavioral_Health_Only` | `FactPlanTransparency` | `DeniedNotMedicallyNecessaryBHOnly` + availability row | bigint nullable |
| 36 | `Plan_Number_Claims_Denied_Due_To_Enrolle_Benefit_Limit_Reached` | `FactPlanTransparency` | `DeniedBenefitLimitReached` + availability row | bigint nullable |
| 37 | `Plan_Number_Claims_Denied_Due_To_Member_Not_Covered` | `FactPlanTransparency` | `DeniedMemberNotCovered` + availability row | bigint nullable |
| 38 | `Plan_Number_Claims_Denied_Due_To_Investigational_Experimental_Cosmetic_Proceduce` | `FactPlanTransparency` | `DeniedInvestigationalExperimentalCosmetic` + availability row | bigint nullable |
| 39 | `Plan_Number_Claims_Denied_Due_To_Administrative_Reason` | `FactPlanTransparency` | `DeniedAdministrativeReason` + availability row | bigint nullable |
| 40 | `Plan_Number_Claims_Denied_Other` | `FactPlanTransparency` | `DeniedOther` + availability row | bigint nullable |
| 41 | `Rate_Review` | `DimIssuer` | `RateReviewURL` | nullable URL |
| 42 | `Financial_Information` | `DimIssuer` | `FinancialInformationURL` + status | nullable URL |
| 43 | `Average Monthly Enrollment` | `FactPlanTransparency` | `AverageMonthlyEnrollment` + availability row | decimal nullable |
| 44 | `Average Monthly Disenrollment` | `FactPlanTransparency` | `AverageMonthlyDisenrollment` + availability row | decimal nullable |

---

# 8. Derived Canonical Metadata

The following are derived and do not come from one of the 44 source columns.

## 8.1 DimReportingPeriod

```text
ReportingPeriodKey = 2024
ExperienceYear     = 2024
PUFYear            = 2026
PeriodStartDate    = 2024-01-01
PeriodEndDate      = 2024-12-31
```

## 8.2 Source lineage

For every plan:

```text
SourceSheet
SourceExcelRow
SourceFileName
SourceFileSHA256
```

For issuer records consolidated from repeated plan rows:

```text
SourcePlanRowCount
SourceSheetCount
SourceSheets
SourceFileName
SourceFileSHA256
```

Lineage fields are technical/audit fields and should be hidden from normal BI consumers.

---

# 9. DimMetric Contract

`DimMetric` contains the 30 numeric analytical metrics:

```text
12 issuer-level metrics
18 plan-level metrics
```

Minimum fields:

```text
MetricKey
MetricCode
MetricDisplayName
EntityLevel            -- ISSUER | PLAN
MetricFamily           -- CLAIMS | DENIALS | RESUBMISSIONS | APPEALS | ENROLLMENT
NetworkScope           -- IN_NETWORK | OUT_OF_NETWORK | ALL | NOT_APPLICABLE
IsPercentage
IsAdditive
SourceColumnName
ExperienceYear
```

Rules:

- count metrics are additive only at their governed entity grain and across mutually exclusive entities;
- published appeal percentages are non-additive;
- percentages must never be summed;
- aggregate appeal percentages are recalculated as ratio-of-totals from filed/overturned counts when an aggregate rate is required;
- enrollment/disenrollment are plan-level averages and require explicit aggregation semantics before any higher-level total is published.

---

# 10. DimAvailabilityStatus Contract

Minimum rows:

| StatusKey | StatusCode | IsNumericAvailable | IsDisclosureLimited | IsDQIssue |
|---:|---|---:|---:|---:|
| 1 | `AVAILABLE` | 1 | 0 | 0 |
| 2 | `NOT_AVAILABLE` | 0 | 1 | 0 |
| 3 | `SUPPRESSED_SMALL_CELL` | 0 | 1 | 0 |
| 4 | `NOT_REQUIRED_PLAN_TYPE` | 0 | 0 | 0 |
| 5 | `NOT_APPLICABLE_NEW_ENTITY` | 0 | 0 | 0 |
| 6 | `MISSING_URL` | 0 | 0 | 0 |
| 7 | `SOURCE_MISSING_UNEXPECTED` | 0 | 0 | 1 |
| 8 | `INVALID_NUMERIC_SOURCE` | 0 | 0 | 1 |

---

# 11. Relationship Contract

Target semantic relationships:

```text
DimReportingPeriod 1 ---- * FactIssuerTransparency
DimReportingPeriod 1 ---- * FactPlanTransparency

DimState           1 ---- * FactIssuerTransparency
DimState           1 ---- * FactPlanTransparency

DimIssuer          1 ---- * FactIssuerTransparency
DimIssuer          1 ---- * FactPlanTransparency

DimPlan            1 ---- * FactPlanTransparency

DimMetric          1 ---- * FactIssuerMetricAvailability
DimMetric          1 ---- * FactPlanMetricAvailability

DimAvailabilityStatus 1 -- * FactIssuerMetricAvailability
DimAvailabilityStatus 1 -- * FactPlanMetricAvailability
```

Default relationship direction:

```text
Dimension -> Fact
single direction
```

No direct fact-to-fact relationships.

No many-to-many relationship is authorized.

---

# 12. Transformation Rules

## 12.1 Text

- trim leading/trailing whitespace;
- preserve official business labels unless a separate standardized attribute is intentionally created;
- never rewrite source labels silently.

## 12.2 Issuer ID

```text
convert to text
validate exactly 5 digits
do not treat as number
```

## 12.3 Plan ID

```text
convert to text
validate exactly 14 characters
validate pattern: 5-digit Issuer ID + 2-letter State + 7 digits
validate embedded State == source State
validate embedded Issuer ID == source Issuer_ID
```

## 12.4 Boolean flags

Only:

```text
Yes -> True
No  -> False
```

Any other nonblank value is a DQ exception.

## 12.5 Numeric metrics

Processing sequence:

```text
raw source value
→ detect CMS special status
→ if special: numeric value = NULL + mapped StatusCode
→ else guarded numeric conversion
→ if conversion succeeds: value + AVAILABLE
→ if conversion fails: NULL + INVALID_NUMERIC_SOURCE + DQ exception
```

No imputation.

No replacement of unavailable/suppressed values with zero.

---

# 13. KPI-Safety Rules Established by the Data Contract

The data contract does not yet define the full KPI dictionary, but it establishes these mandatory calculation rules.

## 13.1 Denial rate

At a governed aggregation level:

```text
Denied Claims / Received Claims
```

must be implemented as ratio-of-totals from comparable available values.

Do not average plan-level denial rates.

## 13.2 Resubmission rate

```text
Resubmitted Claims / Denied Claims
```

only when numerator and denominator are semantically comparable and available.

## 13.3 Appeal overturn rate

Published source percentages are preserved for source reconciliation.

For analytical aggregation:

```text
SUM(Overturned) / SUM(Filed) × 100
```

Do not average published percentages.

## 13.4 Denial reason composition

Plan denial-reason fields describe plan-level denial categories.

They must not automatically be assumed to form a mutually exclusive exhaustive decomposition of every denial count unless the KPI validation stage proves the required relationship.

## 13.5 Enrollment / disenrollment

Average Monthly Disenrollment is a subset of Average Monthly Enrollment at the source-defined level.

Any higher-level aggregation must be explicitly validated before publication because these are average monthly measures rather than simple event counts.

---

# 14. Data Quality Contract

The canonical transformation must detect and report at least:

```text
duplicate Plan_ID
duplicate issuer fact grain
missing Plan_ID
missing Issuer_ID
missing State
invalid Issuer_ID pattern
invalid Plan_ID pattern
Plan_ID state mismatch
Plan_ID issuer mismatch
issuer attribute inconsistency
issuer measure inconsistency across repeated plan rows
unexpected blank required fields
invalid numeric source
negative claim/appeal counts
denied > received
resubmitted > denied (review, not automatic failure)
disenrollment > enrollment
appeals overturned > appeals filed
published appeal percentage mismatch
percentage outside expected 0–100 range
```

A DQ finding never authorizes silent source correction.

---

# 15. Known Source Exception Register — Initial

Current confirmed source exception:

```text
Issuer_ID: 97725
Internal Appeals Filed: 1,203
Internal Appeals Overturned: 16,518
Published Internal Appeal Overturn %: 1,373.07
```

Treatment:

```text
PRESERVE SOURCE
FLAG DQ
DO NOT CLIP
DO NOT IMPUTE
DO NOT DELETE
```

The published percentage reconciles mathematically to the published numerator and denominator, so this is preserved as a source-reported logical anomaly rather than a transformation error.

---

# 16. Canonical Build Acceptance Criteria

The transformation checkpoint passes only if all required gates pass.

## 16.1 Dimensions

```text
DimReportingPeriod rows = 1
DimState rows           = 30
DimIssuer rows          = 348
DimPlan rows            = 4,956
```

## 16.2 Facts

```text
FactIssuerTransparency rows = 348
FactPlanTransparency rows   = 4,956
```

## 16.3 Availability

```text
FactIssuerMetricAvailability rows = 4,176
FactPlanMetricAvailability rows   = 89,208
```

## 16.4 Integrity

```text
duplicate Plan business keys                    = 0
duplicate issuer fact grain                     = 0
broken Plan -> Issuer keys                      = 0
broken Plan -> State keys                       = 0
broken Issuer -> State keys                     = 0
unexplained source-row loss                     = 0
issuer repeated-measure inconsistencies         = 0
unexpected invalid numerics                     = 0 unless registered DQ
source special-state preservation mismatches    = 0
```

## 16.5 Reconciliation

```text
Internal appeal published % formula comparison = PASS for comparable population
External appeal published % formula comparison = PASS for comparable population
```

Known registered anomalies remain visible and do not cause silent alteration.

---

# 17. Output Contract for the Canonical Build

Final project outputs only:

```text
data/canonical/
├── dim_reporting_period.csv
├── dim_state.csv
├── dim_issuer.csv
├── dim_plan.csv
├── dim_metric.csv
├── dim_availability_status.csv
├── fact_issuer_transparency.csv
├── fact_plan_transparency.csv
├── fact_issuer_metric_availability.csv
├── fact_plan_metric_availability.csv
└── dq_exception_register.csv
```

Review evidence is written outside the project to the real Windows Desktop review folder:

```text
Transparency_PUF_Review/
└── 02_CANONICAL_BUILD_REPORT.xlsx
```

No temporary discovery/transformation files are authorized inside the project tree.

---

# 18. Evidence State

```text
CANONICAL DATA CONTRACT = APPROVED FOR IMPLEMENTATION
CANONICAL DATASET        = NOT YET BUILT
KPI CONTRACT             = NOT YET APPROVED
EDA                      = NOT YET STARTED
SQL                      = NOT YET AUTHORIZED
POWER BI                 = NOT YET AUTHORIZED
```
