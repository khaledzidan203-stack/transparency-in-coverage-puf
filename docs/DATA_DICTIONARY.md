# Canonical dataset dictionary

Business definitions, nullability and field mappings remain in the [data contract](DATA_CONTRACT.md). This inventory describes the committed CSV headers.

| File | Rows | Columns |
|---|---:|---|
| dim_availability_status.csv | 8 | `StatusKey`, `StatusCode`, `IsNumericAvailable`, `IsDisclosureLimited`, `IsDQIssue` |
| dim_issuer.csv | 348 | `IssuerKey`, `IssuerID`, `IssuerName`, `StateKey`, `IsIssuerNewToExchange`, `IsSADPOnlyIssuer`, `ClaimsPaymentPoliciesURL`, `RateReviewURL`, `FinancialInformationURL`, `FinancialInformationStatusCode` |
| dim_metric.csv | 30 | `MetricKey`, `MetricCode`, `MetricDisplayName`, `EntityLevel`, `MetricFamily`, `NetworkScope`, `IsPercentage`, `IsAdditive`, `SourceColumnName`, `ExperienceYear` |
| dim_plan.csv | 4956 | `PlanKey`, `PlanID`, `IssuerKey`, `StateKey`, `MarketSegment`, `PlanType`, `PlanOfferingType`, `MetalLevel`, `SourceSheet`, `SourceExcelRow`, `SourceFileName`, `SourceFileSHA256` |
| dim_reporting_period.csv | 1 | `ReportingPeriodKey`, `PUFYear`, `ExperienceYear`, `PeriodStartDate`, `PeriodEndDate` |
| dim_state.csv | 30 | `StateKey`, `StateCode`, `ExchangeType` |
| dq_exception_register.csv | 38 | `DQExceptionID`, `Severity`, `Scope`, `EntityKey`, `SourceField`, `RuleCode`, `ObservedValue`, `ExpectedCondition`, `SourceSheet`, `SourceExcelRow`, `IsKnownSourceException`, `ResolutionState` |
| fact_issuer_metric_availability.csv | 4176 | `ReportingPeriodKey`, `StateKey`, `IssuerKey`, `MetricKey`, `AvailabilityStatusKey`, `SourceStatusToken` |
| fact_issuer_transparency.csv | 348 | `ReportingPeriodKey`, `StateKey`, `IssuerKey`, `ClaimsReceivedOutOfNetwork`, `ClaimsReceivedInNetwork`, `ClaimsDeniedOutOfNetwork`, `ClaimsDeniedInNetwork`, `ClaimsResubmittedOutOfNetwork`, `ClaimsResubmittedInNetwork`, `InternalAppealsFiled`, `InternalAppealsOverturned`, `InternalAppealsOverturnPctPublished`, `ExternalAppealsFiled`, `ExternalAppealsOverturned`, `ExternalAppealsOverturnPctPublished`, `SourcePlanRowCount`, `SourceSheetCount`, `SourceSheets`, `SourceFileName`, `SourceFileSHA256` |
| fact_plan_metric_availability.csv | 89208 | `ReportingPeriodKey`, `PlanKey`, `MetricKey`, `AvailabilityStatusKey`, `SourceStatusToken` |
| fact_plan_transparency.csv | 4956 | `ReportingPeriodKey`, `StateKey`, `IssuerKey`, `PlanKey`, `ClaimsReceivedOutOfNetwork`, `ClaimsReceivedInNetwork`, `ClaimsDeniedOutOfNetwork`, `ClaimsDeniedInNetwork`, `ClaimsResubmittedOutOfNetwork`, `ClaimsResubmittedInNetwork`, `DeniedReferralOrPriorAuthorization`, `DeniedDueToOutOfNetworkProvider`, `DeniedServicesExcluded`, `DeniedNotMedicallyNecessaryExclBH`, `DeniedNotMedicallyNecessaryBHOnly`, `DeniedBenefitLimitReached`, `DeniedMemberNotCovered`, `DeniedInvestigationalExperimentalCosmetic`, `DeniedAdministrativeReason`, `DeniedOther`, `AverageMonthlyEnrollment`, `AverageMonthlyDisenrollment` |
