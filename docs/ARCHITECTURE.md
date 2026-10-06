# Implemented project architecture

> **Governance note:** the saved PBIP/TMDL is authoritative for the current physical relationship topology. The [implementation contract amendment](IMPLEMENTATION_CONTRACT_AMENDMENT.md) reconciles the current model with earlier conceptual relationship sketches in the original contracts.

The saved PBIP, not historical builders, defines the release. [POWER_BI_ARCHITECTURE](POWER_BI_ARCHITECTURE.md) records the original design specification; this document describes the implemented model.

The raw workbook feeds discovery and transformation, producing 11 canonical CSVs. Power Query imports them through `fxLoadCanonicalCsv` using the existing `DataFolderPath` parameter. The report is stored as PBIR and the semantic model as TMDL. There is no implemented SQL serving layer.

| Layer | Objects and role |
|---|---|
| Entity dimensions | DimReportingPeriod, DimState, DimIssuer, DimPlan |
| Governance dimensions | DimMetric, DimAvailabilityStatus |
| Analytical facts | FactIssuerTransparency; FactPlanTransparency |
| Availability facts | FactIssuerMetricAvailability; FactPlanMetricAvailability |
| Exceptions | DQExceptionRegister, intentionally disconnected |
| Measures | _Measures, 43 saved DAX measures |

Issuer business grain is experience year × state × issuer. Plan grain is experience year × plan. The 14 saved relationships are active and single-direction. State filters propagate through DimIssuer, then DimPlan where appropriate; issuer facts connect to DimIssuer and plan facts to DimPlan. The period dimension connects to the four facts. Metric and status dimensions connect to the availability facts. These saved paths supersede conceptual direct-conformed-key sketches in the original contract; the contract's grain and KPI meaning remain unchanged.

DQ entity context is explicitly applied in DAX because the exception register is disconnected. The approved `KEEPFILTERS` correction preserves the Scope filter in `Current Context DQ Exception Count`. Page 10 now passes: ISSUER=5, PLAN=33, Total=38.

The public-release update contains only the approved DQ filter-context DAX correction, the refreshed page-10 screenshot and the saved active-page selection. Analytical data, other DAX, Power Query, relationships and report layout remain unchanged. Users configure the existing data-folder parameter for their own clone. Local `.pbi` state is not versioned.
