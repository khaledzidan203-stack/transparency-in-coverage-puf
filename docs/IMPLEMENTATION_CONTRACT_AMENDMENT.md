# Implementation Contract Amendment

Status: **CURRENT RELEASE GOVERNANCE AMENDMENT**  
Applies to: **PUF 2026 / Experience 2024**  
Release baseline: `9e95b28b8aaaf449145ce12065b1801ed4a4a772`

## Purpose

The original `DATA_CONTRACT.md` and `POWER_BI_ARCHITECTURE.md` remain the approved design history for grain, KPI meaning, availability semantics, identifiers and source preservation. The saved PBIP/TMDL implementation is authoritative for the relationship topology of the current release.

This amendment resolves the documented difference between the earlier conceptual relationship sketches and the implemented semantic model. It does **not** change canonical data, DAX, Power Query, relationships, PBIR, screenshots or validated analytical results.

## Current implemented relationship topology

The current TMDL contains **14 active, single-direction relationships**:

```text
DimState
  └─> DimIssuer
       ├─> FactIssuerTransparency
       ├─> FactIssuerMetricAvailability
       └─> DimPlan
            ├─> FactPlanTransparency
            └─> FactPlanMetricAvailability

DimReportingPeriod
  ├─> FactIssuerTransparency
  ├─> FactPlanTransparency
  ├─> FactIssuerMetricAvailability
  └─> FactPlanMetricAvailability

DimMetric
  ├─> FactIssuerMetricAvailability
  └─> FactPlanMetricAvailability

DimAvailabilityStatus
  ├─> FactIssuerMetricAvailability
  └─> FactPlanMetricAvailability
```

`DQExceptionRegister` remains intentionally disconnected; governed DAX applies entity context explicitly.

## Superseded relationship sketches

For the current release only, the implemented topology above supersedes direct conceptual paths previously shown in:

- `DATA_CONTRACT.md` section **Relationship Contract** where direct State/Issuer-to-plan-fact relationships were sketched.
- `POWER_BI_ARCHITECTURE.md` section **Core relationships** where those direct relationships were also part of the intended design.

Those earlier sections remain valuable design history but are not the current physical semantic-model topology.

## What remains unchanged and authoritative

The following original contract rules remain unchanged:

- issuer business grain = Experience Year × State × Issuer;
- plan business grain = Experience Year × Plan;
- issuer and plan facts must never be summed together;
- no fact-to-fact relationship is authorized;
- no many-to-many relationship is authorized;
- availability status remains a first-class governed concept;
- suppression/unavailability is never converted to numeric zero;
- analytical rates use governed comparable populations and ratio-of-totals logic;
- source anomalies are preserved, flagged and disclosed rather than silently repaired.

## Precedence

For the current release:

1. Canonical grain and business meaning: `DATA_CONTRACT.md` and `KPI_CONTRACT.md`.
2. Physical implemented relationship topology: saved TMDL plus this amendment.
3. Report implementation: saved PBIR/PBIP.
4. Validation status: `VALIDATION_EVIDENCE.md` and `GITHUB_RELEASE_CHECKLIST.md`.

Any future relationship change requires a new reviewed amendment or a versioned contract update plus revalidation.