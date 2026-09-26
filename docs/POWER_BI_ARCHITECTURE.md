---
document_id: TIC-PBI-ARCH-001
title: Transparency in Coverage PUF — Power BI Architecture & Page Design
version: "1.0"
status: APPROVED_FOR_IMPLEMENTATION
depends_on:
  - docs/DATA_CONTRACT.md
  - docs/KPI_CONTRACT.md
  - docs/ANALYTICAL_STORY.md
---

# Transparency in Coverage PUF — Power BI Architecture & Page Design

## 1. Architecture Decision

### Primary serving layer

Use the final governed canonical CSV files in:

```text
data/canonical/
```

as the Power BI Import-mode source.

### SQL decision

SQL Server is **not required as the primary serving layer for the current release**.

Reason:

- the governed canonical model is already materialized;
- the dataset is modest in size;
- the largest canonical fact is under 100k rows;
- Python has already performed the complex source parsing, status preservation, grain separation, and validation;
- adding SQL now would introduce another transformation/serving layer without a current performance or operational need.

SQL may be introduced later if:

- multiple PUF years are appended;
- refresh orchestration becomes automated;
- row volumes grow materially;
- centralized serving is required;
- the project explicitly needs a separate SQL implementation for portfolio demonstration.

No SQL layer should be added merely to increase tool count.

---

# 2. Power BI Project Format

Use:

```text
Power BI Project (PBIP)
```

Preferred workflow:

```text
Canonical data
→ Power Query import
→ semantic model
→ governed DAX
→ reconciliation
→ page-by-page PBIR development
→ visual QA
→ final release
```

Do not create report visuals before:

1. data source paths are parameterized;
2. relationships are validated;
3. hidden technical columns are defined;
4. core measures are implemented;
5. DAX-to-Python KPI reconciliation passes.

---

# 3. Tables to Load

Load only final canonical tables:

```text
DimReportingPeriod
DimState
DimIssuer
DimPlan
DimMetric
DimAvailabilityStatus

FactIssuerTransparency
FactPlanTransparency
FactIssuerMetricAvailability
FactPlanMetricAvailability

DQExceptionRegister
```

Do not load:

```text
RAW Excel workbook
discovery review workbooks
canonical build review workbook
KPI validation workbook
EDA workbook
temporary files
```

The review workbooks are evidence, not BI sources.

---

# 4. Semantic Model

## 4.1 Core relationships

```text
DimReportingPeriod[ReportingPeriodKey]
    1 → *
FactIssuerTransparency[ReportingPeriodKey]

DimReportingPeriod[ReportingPeriodKey]
    1 → *
FactPlanTransparency[ReportingPeriodKey]

DimState[StateKey]
    1 → *
FactIssuerTransparency[StateKey]

DimState[StateKey]
    1 → *
FactPlanTransparency[StateKey]

DimIssuer[IssuerKey]
    1 → *
FactIssuerTransparency[IssuerKey]

DimIssuer[IssuerKey]
    1 → *
FactPlanTransparency[IssuerKey]

DimPlan[PlanKey]
    1 → *
FactPlanTransparency[PlanKey]

DimMetric[MetricKey]
    1 → *
FactIssuerMetricAvailability[MetricKey]

DimMetric[MetricKey]
    1 → *
FactPlanMetricAvailability[MetricKey]

DimAvailabilityStatus[StatusKey]
    1 → *
FactIssuerMetricAvailability[AvailabilityStatusKey]

DimAvailabilityStatus[StatusKey]
    1 → *
FactPlanMetricAvailability[AvailabilityStatusKey]
```

All relationships:

```text
One-to-many
Single-direction
Dimension → Fact
```

No many-to-many relationships.

No direct fact-to-fact relationship.

## 4.2 DQExceptionRegister

Keep `DQExceptionRegister` disconnected in the first release.

Reason:

It contains both issuer and plan entity keys in one field and should not create ambiguous polymorphic relationships.

Use governed DAX with `TREATAS` when a page needs to test whether the current IssuerID or PlanID has an applicable DQ exception.

---

# 5. Power Query Rules

Power Query is a loading and type-assignment layer only.

Allowed:

- parameterized folder path;
- CSV import;
- column type enforcement;
- friendly query names;
- optional technical column removal after validation.

Not allowed:

- business KPI calculation;
- source-value correction;
- suppression-to-zero conversion;
- issuer/plan grain merging;
- hidden reconciliation logic;
- silently dropping DQ rows.

All material business transformations remain upstream in the canonical Python layer.

---

# 6. Model Visibility

## Visible business fields

### DimState

```text
StateCode
ExchangeType
```

### DimIssuer

```text
IssuerID
IssuerName
IsIssuerNewToExchange
IsSADPOnlyIssuer
ClaimsPaymentPoliciesURL
RateReviewURL
FinancialInformationURL
FinancialInformationStatusCode
```

### DimPlan

```text
PlanID
MarketSegment
PlanType
PlanOfferingType
MetalLevel
```

### DimMetric

```text
MetricDisplayName
EntityLevel
MetricFamily
NetworkScope
```

### DimAvailabilityStatus

```text
StatusCode
```

## Hidden technical fields

Hide surrogate keys and lineage fields from report consumers unless required for diagnostics.

Examples:

```text
StateKey
IssuerKey
PlanKey
MetricKey
ReportingPeriodKey
SourceExcelRow
SourceFileSHA256
SourcePlanRowCount
SourceSheetCount
SourceSheets
```

---

# 7. Measure Display Folders

Recommended measure folders:

```text
01 Population
02 Claims
03 Denials
04 Network Comparison
05 Resubmissions
06 Appeals
07 Denial Reasons
08 Enrollment
09 Availability
10 Data Quality
11 Sensitivity
99 Utility
```

No implicit numeric-column aggregation should be exposed to end users.

Hide raw numeric fact columns after validated DAX measures are created.

---

# 8. Core Measure Families

The Power BI model must implement the KPI contract, including:

## Population

```text
Issuer Count
Plan Count
Eligible Issuer Count
Eligible Plan Count
```

## Claims

```text
Claims Received — In Network
Claims Received — Out of Network
Comparable Claims Received — Total
```

## Denials

```text
Claims Denied — In Network
Claims Denied — Out of Network
Comparable Claims Denied — Total
In-Network Denial Rate
Out-of-Network Denial Rate
Comparable Overall Denial Rate
```

## Resubmissions

```text
Resubmission Events — In Network
Resubmission Events — Out of Network
Resubmission Events per 100 Denied — In Network
Resubmission Events per 100 Denied — Out of Network
```

## Appeals

```text
Internal Appeals Filed
Internal Appeals Overturned
Internal Appeal Overturn Rate

External Appeals Filed
External Appeals Overturned
External Appeal Overturn Rate
```

## Denial Reasons

```text
Denial Reason Count
Comparable Denial Reason Composition %
```

## Enrollment

```text
Reported Average Monthly Enrollment
Reported Average Monthly Disenrollment
Comparable Disenrollment-to-Enrollment Ratio
```

## Availability

```text
Available Entity Count
Applicable Entity Count
Data Availability Rate
Suppression Rate
Unavailable Rate
Structural Non-Applicability Count
```

## DQ

```text
Open DQ Exception Count
Known Source Exception Count
Current Issuer Has DQ Flag
Current Plan Has DQ Flag
```

---

# 9. Sensitivity Measures

Do not silently remove DQ exceptions.

For selected KPIs, implement separate sensitivity measures.

Examples:

```text
Comparable Overall Denial Rate — Reported
Comparable Overall Denial Rate — Excluding Denied>Received Exceptions

Internal Appeal Overturn Rate — Reported
Internal Appeal Overturn Rate — Excluding Known Source Exception
```

The default headline measure is the reported-source measure.

Sensitivity measures must always be explicitly labeled.

---

# 10. Report Navigation

Recommended report structure:

```text
00 INDEX
01 Executive Overview
02 Claims & Denials
03 Network Comparison
04 Denial Reasons
05 Appeals & Resubmissions
06 Enrollment & Availability
07 State Explorer
08 Issuer Explorer
09 Plan Explorer
10 Data Quality & Methodology
```

Every analytical page includes:

```text
Home / INDEX button
Page title
Scope subtitle
Filter context
Data availability / DQ context where relevant
```

---

# 11. Page Specifications

## 00 — INDEX

Purpose:

Provide report orientation and controlled navigation.

Tiles:

```text
Executive Overview
Claims & Denials
Network Comparison
Denial Reasons
Appeals & Resubmissions
Enrollment & Availability
State Explorer
Issuer Explorer
Plan Explorer
Data Quality & Methodology
```

Include a compact scope card:

```text
CMS Transparency in Coverage PUF 2026
Experience Year 2024
30 states
348 issuers
4,956 plans
```

---

## 01 — Executive Overview

Purpose:

Answer: “What does the dataset show at a high level?”

Headline cards:

```text
Reported Claims Received — Comparable
Reported Claims Denied — Comparable
Comparable Overall Denial Rate
Eligible Issuers
Eligible Plans
Data Availability indicator
Open DQ Exceptions
```

Primary visuals:

1. In-network vs out-of-network denial rate
2. Reported claims volume by network
3. Overall denial rate by State
4. Denial reason composition
5. Availability summary

Required annotation:

```text
Reported self-disclosed metrics.
Not a quality ranking.
Rates use comparable eligible populations.
```

---

## 02 — Claims & Denials

Purpose:

Detailed claims-volume and denial analysis.

Visuals:

- Claims Received: In vs Out
- Claims Denied: In vs Out
- Denial rate by State
- denial rate by Market Segment
- denial rate by Plan Type
- denial rate by Metal Level
- claims-volume distribution

Required companion fields:

```text
Eligible entity count
Availability rate
DQ indicator
```

---

## 03 — Network Comparison

Purpose:

Understand differences between in-network and out-of-network reported patterns.

Visuals:

- In-network denial rate
- Out-of-network denial rate
- network gap in percentage points
- State comparison scatter:
  - X = In-network denial rate
  - Y = Out-of-network denial rate
  - Size = comparable claims volume
- issuer comparison scatter with DQ flag

Important:

Out-of-network rates above 100 must remain visible and be marked as DQ/source exceptions rather than clipped.

---

## 04 — Denial Reasons

Purpose:

Explain reported denial-reason composition.

Default population:

```text
Plans where all 10 reason fields are AVAILABLE
```

Headline:

```text
Comparable Plans = 1,089
```

Visuals:

- horizontal bar chart of 10 denial reasons
- reason composition by Market Segment
- reason composition by Plan Type
- reason detail matrix

Mandatory footnote:

```text
Composition of reported reason-count fields in the fully comparable population.
Not automatically the share of all denied claims.
```

---

## 05 — Appeals & Resubmissions

Purpose:

Describe post-denial processes.

Sections:

### Resubmissions

- events per 100 denied claims
- In vs Out Network
- Issuer vs Plan grain
- count of entities above 100

### Appeals

- Internal Appeals Filed
- Internal Appeals Overturned
- Internal Overturn Rate
- External Appeals Filed
- External Appeals Overturned
- External Overturn Rate

Include:

```text
Reported vs sensitivity rate
DQ badge for Issuer 97725
```

---

## 06 — Enrollment & Availability

Purpose:

Prevent overinterpretation of incomplete enrollment data.

Headline cards:

```text
Enrollment Available Plans       2,344
Disenrollment Available Plans    1,340
Enrollment Availability          51.13%
Disenrollment Availability       29.23%
```

Visuals:

- status distribution by enrollment metric
- reported enrollment by segment
- reported disenrollment by segment
- comparable disenrollment-to-enrollment ratio

This page must emphasize source completeness before totals.

---

## 07 — State Explorer

Purpose:

State-level descriptive exploration.

Filters:

```text
State
Exchange Type
Market Segment
Plan Offering Type
Plan Type
Metal Level
```

Visuals:

- State headline KPIs
- In vs Out Network
- eligible issuers/plans
- issuer list
- denial reason composition
- DQ indicator

State rates must display eligible population and DQ status.

---

## 08 — Issuer Explorer

Purpose:

Issuer drill-down without ranking issuers as good/bad.

Selector:

```text
Issuer Name / Issuer ID
```

Visuals:

- claims received
- claims denied
- denial rates by network
- resubmissions
- appeals
- appeal overturn rates
- issuer's plans
- availability matrix
- DQ/source-exception banner

Include source URLs where present:

```text
Claims Payment Policies
Rate Review
Financial Information
```

---

## 09 — Plan Explorer

Purpose:

Plan-level drill-down.

Selector:

```text
Plan ID
```

Display:

```text
Issuer
State
Market Segment
Plan Offering Type
Plan Type
Metal Level
```

Visuals:

- claims received/denied by network
- plan denial rates
- resubmissions
- ten denial reasons
- enrollment/disenrollment availability and values
- plan DQ banner

---

## 10 — Data Quality & Methodology

Purpose:

Make the analytical limitations auditable.

Headline cards:

```text
Open DQ Exceptions         36
Known Source Exception      2 rows
Known Exception Entities    1
```

Visuals:

- DQ by rule
- DQ by scope
- availability by metric
- suppression rate
- methodology table
- status definitions

Display the project rules:

```text
*     NOT_AVAILABLE
**    SUPPRESSED_SMALL_CELL
***   NOT_REQUIRED_PLAN_TYPE
N/A   NOT_APPLICABLE_NEW_ENTITY
```

Include methodology notes:

```text
ratio of totals
common eligible population
no suppression-to-zero
no silent correction
issuer and plan grain remain separate
```

---

# 12. Global Slicers

Use only slicers that are semantically valid on the page.

Common slicers:

```text
State
Exchange Type
Issuer
Market Segment
Plan Offering Type
Plan Type
Metal Level
```

Do not place Plan-level slicers on issuer-only pages unless the filter propagation is explicitly designed and validated.

---

# 13. Visual Design Direction

Professional healthcare/regulatory analytics style.

Principles:

- clean light canvas;
- strong information hierarchy;
- restrained accent colors;
- consistent KPI-card system;
- minimal decorative shapes;
- high contrast and accessible labels;
- visible units;
- no chart junk;
- no 3D visuals;
- no pie charts for large category sets;
- horizontal bars for denial reasons;
- scatterplots for network comparisons;
- matrices only where detail is genuinely useful.

Use conditional formatting for:

```text
DQ flag
availability weakness
source suppression
network gap
```

Do not use red/green to imply moral or quality judgment.

---

# 14. Tooltip Design

Create governed tooltip pages for:

```text
Issuer
Plan
State
Metric Availability
```

Tooltip content should include:

```text
eligible population
numerator
denominator
availability rate
DQ status
```

This allows concise visuals without hiding methodological context.

---

# 15. Performance Rules

Import mode.

Expected model scale is small.

Requirements:

- load only required columns;
- hide surrogate keys;
- avoid calculated columns for KPI logic;
- prefer measures;
- avoid bidirectional relationships;
- avoid many-to-many;
- avoid unnecessary calculated tables;
- keep DQ table disconnected unless a specific relationship design is justified.

---

# 16. Validation Gate Before Visual Build

Before creating analytical pages:

```text
Model relationships                       PASS
Row counts                                PASS
Broken relationship keys                  0
Core DAX measure values vs Python          PASS
Availability measures vs canonical facts  PASS
DQ measures                               PASS
Reported/sensitivity measures             PASS
```

Recommended validation artifact:

```text
Desktop/
└── Transparency_PUF_Review/
    └── 05_POWER_BI_MODEL_VALIDATION.xlsx
```

No visual build should begin until the semantic model gate passes.

---

# 17. Release Sequence

```text
1. Create PBIP shell manually in Power BI Desktop
2. Parameterize canonical data folder
3. Load final canonical tables
4. Define relationships
5. Set visibility / formatting / categories
6. Create governed DAX measures
7. Validate DAX against Python
8. Create INDEX page
9. Build pages one at a time
10. Run visual/filter QA
11. Document screenshots and final reconciliation
12. Release
```

---

# 18. Current Project State

```text
RAW SOURCE                   APPROVED
CANONICAL MODEL              APPROVED
KPI CONTRACT                 APPROVED
KPI VALIDATION               APPROVED
EDA                          APPROVED
ANALYTICAL STORY             APPROVED
SQL SERVING LAYER            NOT REQUIRED CURRENTLY
POWER BI ARCHITECTURE        APPROVED FOR IMPLEMENTATION
PBIP                         NOT YET CREATED
DAX                          NOT YET IMPLEMENTED
REPORT PAGES                 NOT YET BUILT
```
