# Project Case Study

## CMS Transparency in Coverage PUF Analytics

This project turns the CMS Transparency in Coverage Public Use File into a governed analytical model and source-controlled Power BI solution. The implementation is deliberately designed around data meaning, grain, availability and validation rather than treating the source as a flat reporting table.

## 1. Problem

The source contains insurer- and plan-level reporting with different analytical grains, partial disclosure availability, suppression, structural non-applicability and documented source anomalies. A naive model can easily double count issuer values, convert unavailable metrics into zero, average percentages incorrectly or imply complete-market results where the source does not support them.

The project therefore starts with four controls:

1. preserve the source and its meaning;
2. define analytical grain before aggregation;
3. separate numeric values from availability semantics;
4. validate every published analytical layer independently.

## 2. Source and scope

- CMS Transparency in Coverage PUF — publication year 2026 / experience year 2024.
- Public insurer/plan-level data only; no patient-level information.
- 4,956 distinct plans.
- 348 issuers.
- 30 states.
- Source workbook is hash-pinned before transformation.

The project does not infer insurer quality, wrongful denial, fraud or causality from descriptive transparency metrics.

## 3. Analytical architecture

```text
CMS source workbook
        ↓
Python discovery and canonical transformation
        ↓
11 governed canonical CSV tables
        ↓
Issuer grain + Plan grain + Availability facts
        ↓
Power BI import semantic model (TMDL)
        ↓
43 governed DAX measures
        ↓
11-page PBIR analytical report
        ↓
Python validation + runtime DAX reconciliation + GitHub Actions
```

The issuer business grain is `Experience Year × State × Issuer`. The plan business grain is `Experience Year × Plan`. These facts remain separate and are never added together.

## 4. Governance decisions

### Availability is modeled as data

Metric disclosure state is not inferred from numeric values. Availability facts preserve explicit states such as `AVAILABLE`, `NOT_AVAILABLE`, `SUPPRESSED_SMALL_CELL`, `NOT_REQUIRED_PLAN_TYPE`, `NOT_APPLICABLE_NEW_ENTITY` and `MISSING_URL`.

### Zero remains zero

A reported numeric zero is not reclassified as missing, suppressed or unavailable.

### Ratios use governed populations

Rates are calculated as ratio-of-totals over a common eligible population rather than averaging row-level percentages. Comparable metrics require the governed component availability defined by the KPI contract and saved semantic implementation.

### Source anomalies remain visible

The governing policy is:

**PRESERVE → CALCULATE → FLAG → DISCLOSE**

The project does not silently clip, repair or impute unusual source values merely because they look implausible.

## 5. Semantic model

The saved Power BI model contains:

- six dimensions for reporting period, state, issuer, plan, metric and availability status;
- separate issuer and plan analytical facts;
- separate issuer and plan availability facts;
- an intentionally disconnected `DQExceptionRegister`;
- 43 saved DAX measures;
- 14 active, single-direction relationships.

The saved TMDL is authoritative for physical relationship topology. Earlier conceptual relationship sketches are retained as design history and reconciled through the implementation contract amendment.

## 6. Data quality and exception handling

The governed exception register contains 36 `OPEN_REVIEW` rows and 2 preserved known-source rows. These are not automatically treated as invalid records; each remains visible for interpretation.

One preserved example is issuer `97725`, for which the source reports 1,203 internal appeals filed and 16,518 overturned. The published percentage reconciles mathematically to those source values, so the project preserves and flags the anomaly instead of rewriting it.

## 7. Power BI analytical delivery

The source-controlled PBIP/PBIR release contains 11 report pages and 371 saved visuals covering:

- executive overview;
- network comparison;
- denials and denial reasons;
- appeals and resubmissions;
- enrollment;
- issuer/state exploration;
- data availability;
- data quality and methodology.

Manual report refinements remain in source control rather than being regenerated from historical builder scripts.

## 8. Validation strategy

Validation is layered rather than represented by one generic pass/fail flag.

- source integrity and canonical grain checks;
- KPI reconciliation in Python;
- semantic-model structure validation;
- PBIR structure and navigation checks;
- runtime DAX reconciliation against governed baselines;
- release auditing for links, publication hygiene and sensitive paths;
- GitHub Actions static validation and protected-path checks.

Fresh recorded runtime DAX reconciliation contains 63 checks with zero failures: 32 core, 10 denial-reason, 18 availability and 3 DQ-context checks.

## 9. Engineering outcome

The result is not only a dashboard. It is a reproducible analytical system with explicit source provenance, grain contracts, KPI contracts, availability semantics, exception governance, source-controlled semantic/report definitions and an evidence-backed validation path.

The implementation is intentionally conservative: when documentation and saved implementation differ, the difference is documented instead of silently rewriting the validated release. Any future semantic change requires a separately reviewed change and fresh validation.

## Evidence

- [Project README](../README.md)
- [Implemented architecture](ARCHITECTURE.md)
- [Data contract](DATA_CONTRACT.md)
- [KPI contract](KPI_CONTRACT.md)
- [Implementation contract amendment](IMPLEMENTATION_CONTRACT_AMENDMENT.md)
- [DAX/KPI contract audit](DAX_KPI_CONTRACT_AUDIT.md)
- [Data quality governance](DATA_QUALITY_GOVERNANCE.md)
- [Validation evidence](VALIDATION_EVIDENCE.md)
- [Report page guide](REPORT_PAGE_GUIDE.md)
- [Known limitations](KNOWN_LIMITATIONS.md)
