# Technical Walkthrough

This walkthrough gives a concise path through the implemented Transparency in Coverage PUF analytics system without changing or rebuilding the validated release.

## 60–90 second walkthrough

1. **Start with the source boundary**
   - CMS Transparency in Coverage PUF, publication year 2026 / experience year 2024.
   - Public insurer/plan-level data only; no patient-level data.

2. **Show the grain-first design**
   - Issuer facts remain at `Experience Year × State × Issuer`.
   - Plan facts remain at `Experience Year × Plan`.
   - The two grains are never added together.

3. **Show governed availability**
   - Metric availability is modeled separately from numeric facts.
   - Suppressed, unavailable and structurally non-applicable values remain explicit rather than becoming zero.

4. **Show the analytical contract**
   - Rates use ratio-of-totals over common eligible populations.
   - Denial, appeal, resubmission and enrollment measures follow explicit KPI rules.
   - Resubmission intensity may exceed 100 because it represents events per 100 denied claims, not a probability.

5. **Show the semantic model**
   - Canonical CSVs feed the Power BI import model.
   - The saved model contains 43 DAX measures and 14 active single-direction relationships.
   - `DQExceptionRegister` remains intentionally disconnected and is applied through explicit DAX context.

6. **Show the Power BI analytical path**
   - Executive Overview → Network Comparison → Denials → Denial Reasons → Appeals → Resubmissions → Enrollment → Issuer & State Explorer → Data Availability → Data Quality & Methodology.

7. **Close with validation**
   - 11 report pages and screenshots are preserved.
   - Fresh runtime DAX reconciliation passed 32 core, 10 reason, 18 availability and 3 DQ-context checks with zero failures.
   - GitHub Actions runs the static publication quality gate and protected-path checks.

## Key interpretation boundary

The project analyzes reported CMS transparency data. It does not rank insurer quality, establish wrongful denial, infer causality, or treat partially available metrics as complete-market estimates.

## Useful entry points

- [README](../README.md)
- [Implemented architecture](ARCHITECTURE.md)
- [Implementation contract amendment](IMPLEMENTATION_CONTRACT_AMENDMENT.md)
- [KPI contract](KPI_CONTRACT.md)
- [DAX/KPI contract audit](DAX_KPI_CONTRACT_AUDIT.md)
- [Validation framework](VALIDATION_FRAMEWORK.md)
- [Validation evidence](VALIDATION_EVIDENCE.md)
- [Report page guide](REPORT_PAGE_GUIDE.md)
