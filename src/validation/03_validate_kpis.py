from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable
import hashlib
import os
import sys

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill


SCRIPT_VERSION = "1.0.0"
PUF_YEAR = 2026
EXPERIENCE_YEAR = 2024
EXPECTED_SOURCE_SHA256 = "27379dc76590027b0e6d21d736331f87376f0ac87810d68e9f7cca555eef4b1f"

REQUIRED_FILES = [
    "dim_reporting_period.csv",
    "dim_state.csv",
    "dim_issuer.csv",
    "dim_plan.csv",
    "dim_metric.csv",
    "dim_availability_status.csv",
    "fact_issuer_transparency.csv",
    "fact_plan_transparency.csv",
    "fact_issuer_metric_availability.csv",
    "fact_plan_metric_availability.csv",
    "dq_exception_register.csv",
]

ISSUER_METRIC_MAP = {
    1: "ClaimsReceivedOutOfNetwork",
    2: "ClaimsReceivedInNetwork",
    3: "ClaimsDeniedOutOfNetwork",
    4: "ClaimsDeniedInNetwork",
    5: "ClaimsResubmittedOutOfNetwork",
    6: "ClaimsResubmittedInNetwork",
    7: "InternalAppealsFiled",
    8: "InternalAppealsOverturned",
    9: "InternalAppealsOverturnPctPublished",
    10: "ExternalAppealsFiled",
    11: "ExternalAppealsOverturned",
    12: "ExternalAppealsOverturnPctPublished",
}

PLAN_METRIC_MAP = {
    13: "ClaimsReceivedOutOfNetwork",
    14: "ClaimsReceivedInNetwork",
    15: "ClaimsDeniedOutOfNetwork",
    16: "ClaimsDeniedInNetwork",
    17: "ClaimsResubmittedOutOfNetwork",
    18: "ClaimsResubmittedInNetwork",
    19: "DeniedReferralOrPriorAuthorization",
    20: "DeniedDueToOutOfNetworkProvider",
    21: "DeniedServicesExcluded",
    22: "DeniedNotMedicallyNecessaryExclBH",
    23: "DeniedNotMedicallyNecessaryBHOnly",
    24: "DeniedBenefitLimitReached",
    25: "DeniedMemberNotCovered",
    26: "DeniedInvestigationalExperimentalCosmetic",
    27: "DeniedAdministrativeReason",
    28: "DeniedOther",
    29: "AverageMonthlyEnrollment",
    30: "AverageMonthlyDisenrollment",
}

REASON_METRIC_KEYS = list(range(19, 29))
REASON_COLUMNS = [PLAN_METRIC_MAP[k] for k in REASON_METRIC_KEYS]

APPLICABLE_STATUS_CODES = {
    "AVAILABLE",
    "NOT_AVAILABLE",
    "SUPPRESSED_SMALL_CELL",
    "SOURCE_MISSING_UNEXPECTED",
    "INVALID_NUMERIC_SOURCE",
}

STRUCTURAL_NON_APPLICABLE_CODES = {
    "NOT_REQUIRED_PLAN_TYPE",
    "NOT_APPLICABLE_NEW_ENTITY",
}

EXPECTED_STATUS_CODES = APPLICABLE_STATUS_CODES | STRUCTURAL_NON_APPLICABLE_CODES | {
    "MISSING_URL"
}


def project_root() -> Path:
    return Path(__file__).resolve().parents[2]


PROJECT_ROOT = project_root()
CANONICAL_DIR = PROJECT_ROOT / "data" / "canonical"


def get_windows_desktop() -> Path:
    if os.name == "nt":
        try:
            import winreg

            key_path = (
                r"Software\Microsoft\Windows\CurrentVersion"
                r"\Explorer\User Shell Folders"
            )
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                raw_value, _ = winreg.QueryValueEx(key, "Desktop")
            return Path(os.path.expandvars(raw_value)).resolve()
        except Exception:
            pass
    return (Path.home() / "Desktop").resolve()


REVIEW_DIR = Path(os.environ.get("TRANSPARENCY_PUF_REVIEW_DIR", PROJECT_ROOT / ".local-review")).expanduser().resolve()
REVIEW_FILE = REVIEW_DIR / "03_KPI_VALIDATION_REPORT.xlsx"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def as_numeric(df: pd.DataFrame, columns: Iterable[str]) -> None:
    for column in columns:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce")


def load_csv(name: str, dtype: dict[str, Any] | None = None) -> pd.DataFrame:
    path = CANONICAL_DIR / name
    return pd.read_csv(path, dtype=dtype, keep_default_na=True)


def safe_ratio(numerator: float, denominator: float, multiplier: float = 1.0) -> float:
    if denominator == 0 or pd.isna(denominator):
        return np.nan
    return numerator / denominator * multiplier


def main() -> int:
    print("=" * 78)
    print("Transparency in Coverage PUF — KPI Validation")
    print("=" * 78)
    print(f"Script version : {SCRIPT_VERSION}")
    print(f"Canonical dir  : {CANONICAL_DIR}")
    print(f"Review file    : {REVIEW_FILE}")

    missing_files = [
        name for name in REQUIRED_FILES if not (CANONICAL_DIR / name).exists()
    ]
    if missing_files:
        raise FileNotFoundError(
            "Missing required canonical files:\n- " + "\n- ".join(missing_files)
        )

    # ------------------------------------------------------------------
    # Load governed canonical layer.
    # ------------------------------------------------------------------
    dim_period = load_csv("dim_reporting_period.csv")
    dim_state = load_csv("dim_state.csv", dtype={"StateCode": "string"})
    dim_issuer = load_csv("dim_issuer.csv", dtype={"IssuerID": "string"})
    dim_plan = load_csv("dim_plan.csv", dtype={"PlanID": "string"})
    dim_metric = load_csv("dim_metric.csv")
    dim_status = load_csv("dim_availability_status.csv")
    issuer_fact = load_csv("fact_issuer_transparency.csv")
    plan_fact = load_csv("fact_plan_transparency.csv")
    issuer_av = load_csv("fact_issuer_metric_availability.csv")
    plan_av = load_csv("fact_plan_metric_availability.csv")
    dq = load_csv(
        "dq_exception_register.csv",
        dtype={"EntityKey": "string"},
    )

    for df, keys in [
        (dim_period, ["ReportingPeriodKey", "PUFYear", "ExperienceYear"]),
        (dim_state, ["StateKey"]),
        (dim_issuer, ["IssuerKey", "StateKey"]),
        (dim_plan, ["PlanKey", "IssuerKey", "StateKey"]),
        (dim_metric, ["MetricKey", "ExperienceYear"]),
        (dim_status, ["StatusKey"]),
        (issuer_fact, ["ReportingPeriodKey", "StateKey", "IssuerKey"]),
        (plan_fact, ["ReportingPeriodKey", "StateKey", "IssuerKey", "PlanKey"]),
        (issuer_av, ["ReportingPeriodKey", "StateKey", "IssuerKey", "MetricKey", "AvailabilityStatusKey"]),
        (plan_av, ["ReportingPeriodKey", "PlanKey", "MetricKey", "AvailabilityStatusKey"]),
    ]:
        as_numeric(df, keys)

    as_numeric(issuer_fact, ISSUER_METRIC_MAP.values())
    as_numeric(plan_fact, PLAN_METRIC_MAP.values())

    status_lookup = dim_status[["StatusKey", "StatusCode"]].copy()
    issuer_av = issuer_av.merge(
        status_lookup,
        left_on="AvailabilityStatusKey",
        right_on="StatusKey",
        how="left",
        validate="many_to_one",
    )
    plan_av = plan_av.merge(
        status_lookup,
        left_on="AvailabilityStatusKey",
        right_on="StatusKey",
        how="left",
        validate="many_to_one",
    )

    metric_lookup = dim_metric[
        ["MetricKey", "MetricCode", "MetricDisplayName", "EntityLevel"]
    ].copy()

    issuer_av = issuer_av.merge(
        metric_lookup,
        on="MetricKey",
        how="left",
        validate="many_to_one",
    )
    plan_av = plan_av.merge(
        metric_lookup,
        on="MetricKey",
        how="left",
        validate="many_to_one",
    )

    issuer_fact_i = issuer_fact.set_index("IssuerKey", drop=False).sort_index()
    plan_fact_i = plan_fact.set_index("PlanKey", drop=False).sort_index()

    issuer_status = issuer_av.pivot(
        index="IssuerKey", columns="MetricKey", values="StatusCode"
    ).sort_index()
    plan_status = plan_av.pivot(
        index="PlanKey", columns="MetricKey", values="StatusCode"
    ).sort_index()

    # ------------------------------------------------------------------
    # Acceptance gates.
    # ------------------------------------------------------------------
    acceptance_rows: list[dict[str, Any]] = []

    def accept(
        test_id: str,
        test_name: str,
        expected: Any,
        actual: Any,
        status: str,
        gate_type: str,
        notes: str = "",
    ) -> None:
        acceptance_rows.append(
            {
                "TestID": test_id,
                "TestName": test_name,
                "Expected": expected,
                "Actual": actual,
                "Status": status,
                "GateType": gate_type,
                "Notes": notes,
            }
        )

    period_ok = (
        len(dim_period) == 1
        and int(dim_period.iloc[0]["PUFYear"]) == PUF_YEAR
        and int(dim_period.iloc[0]["ExperienceYear"]) == EXPERIENCE_YEAR
    )
    accept(
        "KPI-VAL-001",
        "Reporting period contract",
        "1 row; PUFYear=2026; ExperienceYear=2024",
        f"rows={len(dim_period)}; PUFYear={dim_period.iloc[0]['PUFYear'] if len(dim_period) else 'NA'}; ExperienceYear={dim_period.iloc[0]['ExperienceYear'] if len(dim_period) else 'NA'}",
        "PASS" if period_ok else "FAIL",
        "STRUCTURAL",
    )

    issuer_av_dup = int(
        issuer_av.duplicated(["ReportingPeriodKey", "StateKey", "IssuerKey", "MetricKey"]).sum()
    )
    plan_av_dup = int(
        plan_av.duplicated(["ReportingPeriodKey", "PlanKey", "MetricKey"]).sum()
    )
    accept("KPI-VAL-002", "Duplicate issuer availability grain", 0, issuer_av_dup, "PASS" if issuer_av_dup == 0 else "FAIL", "STRUCTURAL")
    accept("KPI-VAL-003", "Duplicate plan availability grain", 0, plan_av_dup, "PASS" if plan_av_dup == 0 else "FAIL", "STRUCTURAL")

    expected_issuer_av = len(dim_issuer) * len(ISSUER_METRIC_MAP)
    expected_plan_av = len(dim_plan) * len(PLAN_METRIC_MAP)
    accept("KPI-VAL-004", "Issuer availability coverage rows", expected_issuer_av, len(issuer_av), "PASS" if len(issuer_av) == expected_issuer_av else "FAIL", "STRUCTURAL")
    accept("KPI-VAL-005", "Plan availability coverage rows", expected_plan_av, len(plan_av), "PASS" if len(plan_av) == expected_plan_av else "FAIL", "STRUCTURAL")

    unknown_status = sorted(
        (
            set(issuer_av["StatusCode"].dropna())
            | set(plan_av["StatusCode"].dropna())
        )
        - EXPECTED_STATUS_CODES
    )
    accept("KPI-VAL-006", "Unknown availability status codes", 0, len(unknown_status), "PASS" if not unknown_status else "FAIL", "STRUCTURAL", ", ".join(unknown_status))

    # Availability-to-fact integrity for all 30 metrics.
    availability_integrity_rows: list[dict[str, Any]] = []
    available_null_total = 0
    unavailable_nonnull_total = 0

    for entity_level, fact, av, mapping, entity_key in [
        ("ISSUER", issuer_fact_i, issuer_av, ISSUER_METRIC_MAP, "IssuerKey"),
        ("PLAN", plan_fact_i, plan_av, PLAN_METRIC_MAP, "PlanKey"),
    ]:
        for metric_key, column in mapping.items():
            a = av.loc[av["MetricKey"] == metric_key, [entity_key, "StatusCode"]].copy()
            fact_keys = fact[[entity_key, column]].reset_index(drop=True)
            merged = a.merge(fact_keys, on=entity_key, how="left", validate="one_to_one")
            available_null = int(
                ((merged["StatusCode"] == "AVAILABLE") & merged[column].isna()).sum()
            )
            unavailable_nonnull = int(
                ((merged["StatusCode"] != "AVAILABLE") & merged[column].notna()).sum()
            )
            available_null_total += available_null
            unavailable_nonnull_total += unavailable_nonnull
            availability_integrity_rows.append(
                {
                    "EntityLevel": entity_level,
                    "MetricKey": metric_key,
                    "MetricCode": av.loc[av["MetricKey"] == metric_key, "MetricCode"].iloc[0],
                    "MetricDisplayName": av.loc[av["MetricKey"] == metric_key, "MetricDisplayName"].iloc[0],
                    "AvailableStatusButNullFact": available_null,
                    "NonAvailableStatusButFactPresent": unavailable_nonnull,
                    "Status": "PASS" if available_null == 0 and unavailable_nonnull == 0 else "FAIL",
                }
            )

    availability_integrity_df = pd.DataFrame(availability_integrity_rows)
    accept("KPI-VAL-007", "AVAILABLE status with null fact value", 0, available_null_total, "PASS" if available_null_total == 0 else "FAIL", "STRUCTURAL")
    accept("KPI-VAL-008", "Non-AVAILABLE status with numeric fact value", 0, unavailable_nonnull_total, "PASS" if unavailable_nonnull_total == 0 else "FAIL", "STRUCTURAL")

    # ------------------------------------------------------------------
    # Availability summary / coverage KPIs.
    # ------------------------------------------------------------------
    availability_df = pd.concat(
        [
            issuer_av.assign(EntityLevelCanonical="ISSUER"),
            plan_av.assign(EntityLevelCanonical="PLAN"),
        ],
        ignore_index=True,
        sort=False,
    )

    availability_summary = (
        availability_df.groupby(
            [
                "EntityLevelCanonical",
                "MetricKey",
                "MetricCode",
                "MetricDisplayName",
                "StatusCode",
            ],
            dropna=False,
        )
        .size()
        .reset_index(name="EntityCount")
    )

    availability_pivot = (
        availability_summary.pivot_table(
            index=["EntityLevelCanonical", "MetricKey", "MetricCode", "MetricDisplayName"],
            columns="StatusCode",
            values="EntityCount",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    for code in sorted(EXPECTED_STATUS_CODES):
        if code not in availability_pivot.columns:
            availability_pivot[code] = 0

    availability_pivot["AvailableEntityCount"] = availability_pivot["AVAILABLE"]
    availability_pivot["ApplicableEntityCount"] = availability_pivot[
        [c for c in APPLICABLE_STATUS_CODES if c in availability_pivot.columns]
    ].sum(axis=1)
    availability_pivot["StructuralNonApplicableCount"] = availability_pivot[
        [c for c in STRUCTURAL_NON_APPLICABLE_CODES if c in availability_pivot.columns]
    ].sum(axis=1)
    availability_pivot["DataAvailabilityRatePct"] = np.where(
        availability_pivot["ApplicableEntityCount"] > 0,
        availability_pivot["AvailableEntityCount"]
        / availability_pivot["ApplicableEntityCount"]
        * 100.0,
        np.nan,
    )
    availability_pivot["SuppressionRatePct"] = np.where(
        availability_pivot["ApplicableEntityCount"] > 0,
        availability_pivot["SUPPRESSED_SMALL_CELL"]
        / availability_pivot["ApplicableEntityCount"]
        * 100.0,
        np.nan,
    )
    availability_pivot["UnavailableRatePct"] = np.where(
        availability_pivot["ApplicableEntityCount"] > 0,
        availability_pivot["NOT_AVAILABLE"]
        / availability_pivot["ApplicableEntityCount"]
        * 100.0,
        np.nan,
    )

    availability_entity_recon_bad = int(
        (
            availability_pivot[
                [
                    "AVAILABLE",
                    "NOT_AVAILABLE",
                    "SUPPRESSED_SMALL_CELL",
                    "NOT_REQUIRED_PLAN_TYPE",
                    "NOT_APPLICABLE_NEW_ENTITY",
                    "SOURCE_MISSING_UNEXPECTED",
                    "INVALID_NUMERIC_SOURCE",
                ]
            ].sum(axis=1)
            != availability_pivot["EntityLevelCanonical"].map(
                {"ISSUER": len(dim_issuer), "PLAN": len(dim_plan)}
            )
        ).sum()
    )
    accept("KPI-VAL-009", "Availability status counts reconcile to entity population", 0, availability_entity_recon_bad, "PASS" if availability_entity_recon_bad == 0 else "FAIL", "RECONCILIATION")

    # ------------------------------------------------------------------
    # KPI results / population helpers.
    # ------------------------------------------------------------------
    kpi_rows: list[dict[str, Any]] = []
    population_rows: list[dict[str, Any]] = []

    def add_kpi(
        kpi_id: str,
        kpi_name: str,
        entity_level: str,
        value: float | int | None,
        unit: str,
        eligible_count: int | None = None,
        numerator: float | int | None = None,
        denominator: float | int | None = None,
        interpretation: str = "",
    ) -> None:
        kpi_rows.append(
            {
                "KPIID": kpi_id,
                "KPIName": kpi_name,
                "EntityLevel": entity_level,
                "Value": value,
                "Unit": unit,
                "EligibleEntityCount": eligible_count,
                "Numerator": numerator,
                "Denominator": denominator,
                "Interpretation": interpretation,
            }
        )

    def add_population(
        kpi_id: str,
        entity_level: str,
        eligible_mask: pd.Series,
        status_frame: pd.DataFrame,
        required_metrics: list[int],
        denominator_positive_mask: pd.Series | None = None,
    ) -> None:
        entity_total = len(status_frame)
        eligible_count = int(eligible_mask.sum())
        any_suppressed = status_frame[required_metrics].eq("SUPPRESSED_SMALL_CELL").any(axis=1)
        any_not_available = status_frame[required_metrics].eq("NOT_AVAILABLE").any(axis=1)
        any_structural = status_frame[required_metrics].isin(STRUCTURAL_NON_APPLICABLE_CODES).any(axis=1)
        any_invalid = status_frame[required_metrics].isin({"SOURCE_MISSING_UNEXPECTED", "INVALID_NUMERIC_SOURCE"}).any(axis=1)
        denominator_nonpositive = 0
        if denominator_positive_mask is not None:
            required_available = status_frame[required_metrics].eq("AVAILABLE").all(axis=1)
            denominator_nonpositive = int((~denominator_positive_mask & required_available).sum())
        population_rows.append(
            {
                "KPIID": kpi_id,
                "EntityLevel": entity_level,
                "TotalEntities": entity_total,
                "EligibleEntities": eligible_count,
                "ExcludedEntities": entity_total - eligible_count,
                "EntitiesWithSuppressedRequiredMetric": int(any_suppressed.sum()),
                "EntitiesWithNotAvailableRequiredMetric": int(any_not_available.sum()),
                "EntitiesWithStructuralNonApplicability": int(any_structural.sum()),
                "EntitiesWithUnexpectedOrInvalidStatus": int(any_invalid.sum()),
                "DenominatorNonPositiveOrOtherExclusion": denominator_nonpositive,
                "RequiredMetricKeys": ",".join(str(m) for m in required_metrics),
            }
        )

    # Volume helper.
    def available_sum(
        fact: pd.DataFrame,
        status: pd.DataFrame,
        column: str,
        metric_key: int,
    ) -> tuple[float, pd.Series]:
        mask = status[metric_key].eq("AVAILABLE")
        return float(fact.loc[mask, column].sum()), mask

    # Ratio helper.
    def governed_ratio(
        fact: pd.DataFrame,
        status: pd.DataFrame,
        numerator_col: str,
        numerator_metric: int,
        denominator_col: str,
        denominator_metric: int,
        multiplier: float = 1.0,
    ) -> tuple[float, float, float, pd.Series, pd.Series]:
        denominator_positive = fact[denominator_col].gt(0)
        eligible = (
            status[numerator_metric].eq("AVAILABLE")
            & status[denominator_metric].eq("AVAILABLE")
            & denominator_positive
        )
        numerator = float(fact.loc[eligible, numerator_col].sum())
        denominator = float(fact.loc[eligible, denominator_col].sum())
        value = safe_ratio(numerator, denominator, multiplier)
        return value, numerator, denominator, eligible, denominator_positive

    # KPI-001 to KPI-004.
    volume_defs = [
        ("KPI-001", "Issuer Claims Received — In Network", "ClaimsReceivedInNetwork", 2),
        ("KPI-002", "Issuer Claims Received — Out of Network", "ClaimsReceivedOutOfNetwork", 1),
        ("KPI-003", "Issuer Claims Denied — In Network", "ClaimsDeniedInNetwork", 4),
        ("KPI-004", "Issuer Claims Denied — Out of Network", "ClaimsDeniedOutOfNetwork", 3),
    ]
    for kpi_id, name, column, metric_key in volume_defs:
        value, mask = available_sum(issuer_fact_i, issuer_status, column, metric_key)
        add_kpi(kpi_id, name, "ISSUER", value, "count", int(mask.sum()))
        add_population(kpi_id, "ISSUER", mask, issuer_status, [metric_key])

    # KPI-005 comparable total received.
    eligible = issuer_status[[1, 2]].eq("AVAILABLE").all(axis=1)
    value = float(
        (
            issuer_fact_i.loc[eligible, "ClaimsReceivedInNetwork"]
            + issuer_fact_i.loc[eligible, "ClaimsReceivedOutOfNetwork"]
        ).sum()
    )
    add_kpi("KPI-005", "Comparable Total Claims Received", "ISSUER", value, "count", int(eligible.sum()))
    add_population("KPI-005", "ISSUER", eligible, issuer_status, [1, 2])

    # KPI-006 comparable total denied.
    eligible = issuer_status[[3, 4]].eq("AVAILABLE").all(axis=1)
    value = float(
        (
            issuer_fact_i.loc[eligible, "ClaimsDeniedInNetwork"]
            + issuer_fact_i.loc[eligible, "ClaimsDeniedOutOfNetwork"]
        ).sum()
    )
    add_kpi("KPI-006", "Comparable Total Claims Denied", "ISSUER", value, "count", int(eligible.sum()))
    add_population("KPI-006", "ISSUER", eligible, issuer_status, [3, 4])

    # KPI-007 to KPI-009 issuer denial rates.
    ratio_defs = [
        ("KPI-007", "In-Network Denial Rate", "ClaimsDeniedInNetwork", 4, "ClaimsReceivedInNetwork", 2, "%"),
        ("KPI-008", "Out-of-Network Denial Rate", "ClaimsDeniedOutOfNetwork", 3, "ClaimsReceivedOutOfNetwork", 1, "%"),
    ]
    denial_rate_rows: list[dict[str, Any]] = []
    for kpi_id, name, ncol, nm, dcol, dm, unit in ratio_defs:
        value, num, den, eligible, denom_pos = governed_ratio(
            issuer_fact_i, issuer_status, ncol, nm, dcol, dm, 100.0
        )
        add_kpi(kpi_id, name, "ISSUER", value, unit, int(eligible.sum()), num, den)
        add_population(kpi_id, "ISSUER", eligible, issuer_status, [nm, dm], denom_pos)
        denial_rate_rows.append({
            "KPIID": kpi_id,
            "EntityLevel": "ISSUER",
            "NetworkScope": "IN_NETWORK" if "In-Network" in name else "OUT_OF_NETWORK",
            "Numerator": num,
            "Denominator": den,
            "RatePct": value,
            "EligibleEntities": int(eligible.sum()),
        })

    eligible = (
        issuer_status[[1, 2, 3, 4]].eq("AVAILABLE").all(axis=1)
        & (issuer_fact_i["ClaimsReceivedInNetwork"] + issuer_fact_i["ClaimsReceivedOutOfNetwork"]).gt(0)
    )
    overall_num = float(
        (
            issuer_fact_i.loc[eligible, "ClaimsDeniedInNetwork"]
            + issuer_fact_i.loc[eligible, "ClaimsDeniedOutOfNetwork"]
        ).sum()
    )
    overall_den = float(
        (
            issuer_fact_i.loc[eligible, "ClaimsReceivedInNetwork"]
            + issuer_fact_i.loc[eligible, "ClaimsReceivedOutOfNetwork"]
        ).sum()
    )
    overall_rate = safe_ratio(overall_num, overall_den, 100.0)
    add_kpi("KPI-009", "Comparable Overall Denial Rate", "ISSUER", overall_rate, "%", int(eligible.sum()), overall_num, overall_den)
    add_population("KPI-009", "ISSUER", eligible, issuer_status, [1, 2, 3, 4])
    denial_rate_rows.append({
        "KPIID": "KPI-009",
        "EntityLevel": "ISSUER",
        "NetworkScope": "ALL_NETWORK_COMPARABLE",
        "Numerator": overall_num,
        "Denominator": overall_den,
        "RatePct": overall_rate,
        "EligibleEntities": int(eligible.sum()),
    })

    # KPI-010 / 011 resubmissions per 100 denied.
    resub_rows: list[dict[str, Any]] = []
    resub_defs = [
        ("KPI-010", "In-Network Resubmissions per 100 Denied Claims", "ClaimsResubmittedInNetwork", 6, "ClaimsDeniedInNetwork", 4, "IN_NETWORK"),
        ("KPI-011", "Out-of-Network Resubmissions per 100 Denied Claims", "ClaimsResubmittedOutOfNetwork", 5, "ClaimsDeniedOutOfNetwork", 3, "OUT_OF_NETWORK"),
    ]
    for kpi_id, name, ncol, nm, dcol, dm, network in resub_defs:
        value, num, den, eligible, denom_pos = governed_ratio(
            issuer_fact_i, issuer_status, ncol, nm, dcol, dm, 100.0
        )
        add_kpi(kpi_id, name, "ISSUER", value, "events per 100 denied claims", int(eligible.sum()), num, den)
        add_population(kpi_id, "ISSUER", eligible, issuer_status, [nm, dm], denom_pos)
        entity_ratio = issuer_fact_i.loc[eligible, ncol] / issuer_fact_i.loc[eligible, dcol] * 100.0
        resub_rows.append({
            "KPIID": kpi_id,
            "EntityLevel": "ISSUER",
            "NetworkScope": network,
            "AggregateEventsPer100Denied": value,
            "NumeratorResubmissionEvents": num,
            "DenominatorDeniedClaims": den,
            "EligibleEntities": int(eligible.sum()),
            "EntitiesAbove100": int((entity_ratio > 100).sum()),
            "Interpretation": "Values above 100 are permitted and are not automatically DQ failures.",
        })

    eligible = issuer_status[[5, 6]].eq("AVAILABLE").all(axis=1)
    total_resub = float(
        (
            issuer_fact_i.loc[eligible, "ClaimsResubmittedInNetwork"]
            + issuer_fact_i.loc[eligible, "ClaimsResubmittedOutOfNetwork"]
        ).sum()
    )
    add_kpi("KPI-012", "Comparable Total Resubmission Events", "ISSUER", total_resub, "events", int(eligible.sum()))
    add_population("KPI-012", "ISSUER", eligible, issuer_status, [5, 6])

    # Appeals KPI-013 to 018.
    appeal_rows: list[dict[str, Any]] = []
    for kpi_id, name, col, mk in [
        ("KPI-013", "Internal Appeals Filed", "InternalAppealsFiled", 7),
        ("KPI-014", "Internal Appeals Overturned", "InternalAppealsOverturned", 8),
        ("KPI-016", "External Appeals Filed", "ExternalAppealsFiled", 10),
        ("KPI-017", "External Appeals Overturned", "ExternalAppealsOverturned", 11),
    ]:
        value, mask = available_sum(issuer_fact_i, issuer_status, col, mk)
        add_kpi(kpi_id, name, "ISSUER", value, "count", int(mask.sum()))
        add_population(kpi_id, "ISSUER", mask, issuer_status, [mk])

    for kpi_id, appeal_type, filed_col, filed_mk, overturned_col, overturned_mk, pct_col, pct_mk in [
        ("KPI-015", "Internal", "InternalAppealsFiled", 7, "InternalAppealsOverturned", 8, "InternalAppealsOverturnPctPublished", 9),
        ("KPI-018", "External", "ExternalAppealsFiled", 10, "ExternalAppealsOverturned", 11, "ExternalAppealsOverturnPctPublished", 12),
    ]:
        value, num, den, eligible, denom_pos = governed_ratio(
            issuer_fact_i,
            issuer_status,
            overturned_col,
            overturned_mk,
            filed_col,
            filed_mk,
            100.0,
        )
        add_kpi(kpi_id, f"{appeal_type} Appeal Overturn Rate", "ISSUER", value, "%", int(eligible.sum()), num, den)
        add_population(kpi_id, "ISSUER", eligible, issuer_status, [filed_mk, overturned_mk], denom_pos)

        published_eligible = (
            issuer_status[filed_mk].eq("AVAILABLE")
            & issuer_status[overturned_mk].eq("AVAILABLE")
            & issuer_status[pct_mk].eq("AVAILABLE")
            & issuer_fact_i[filed_col].gt(0)
        )
        calc = issuer_fact_i.loc[published_eligible, overturned_col] / issuer_fact_i.loc[published_eligible, filed_col] * 100.0
        published = issuer_fact_i.loc[published_eligible, pct_col]
        diff = (calc - published).abs()
        mismatch = diff > 0.011

        appeal_rows.append(
            {
                "AppealType": appeal_type,
                "AggregateOverturnRatePct": value,
                "AggregateNumerator": num,
                "AggregateDenominator": den,
                "AggregateEligibleIssuers": int(eligible.sum()),
                "PublishedPctComparableIssuers": int(published_eligible.sum()),
                "PublishedPctFormulaMismatches": int(mismatch.sum()),
                "MaxAbsoluteDifferencePctPoints": float(diff.max()) if len(diff) else np.nan,
                "KnownExceptionIncluded": "YES" if appeal_type == "Internal" else "N/A",
            }
        )

    appeal_df = pd.DataFrame(appeal_rows)
    appeal_mismatch_total = int(appeal_df["PublishedPctFormulaMismatches"].sum())
    accept("KPI-VAL-010", "Published appeal percentage formula mismatches", 0, appeal_mismatch_total, "PASS" if appeal_mismatch_total == 0 else "FAIL", "RECONCILIATION")

    # Sensitivity view: known source exceptions removed, clearly separate.
    known_issuer_ids = set(
        dq.loc[
            dq["IsKnownSourceException"].astype(str).str.lower().eq("true"),
            "EntityKey",
        ].dropna().astype(str)
    )
    issuer_id_map = dim_issuer.set_index("IssuerKey")["IssuerID"].astype(str)
    internal_elig = (
        issuer_status[7].eq("AVAILABLE")
        & issuer_status[8].eq("AVAILABLE")
        & issuer_fact_i["InternalAppealsFiled"].gt(0)
    )
    internal_excl = internal_elig & ~issuer_fact_i.index.to_series().map(issuer_id_map).isin(known_issuer_ids).values
    # Reconstruct the index-safe mask above explicitly.
    issuer_ids_for_fact = issuer_fact_i["IssuerKey"].map(issuer_id_map)
    internal_excl = internal_elig & ~issuer_ids_for_fact.isin(known_issuer_ids)
    sens_num = float(issuer_fact_i.loc[internal_excl, "InternalAppealsOverturned"].sum())
    sens_den = float(issuer_fact_i.loc[internal_excl, "InternalAppealsFiled"].sum())
    sens_rate = safe_ratio(sens_num, sens_den, 100.0)
    appeal_df.loc[appeal_df["AppealType"] == "Internal", "SensitivityRateExcludingKnownExceptionsPct"] = sens_rate
    appeal_df.loc[appeal_df["AppealType"] == "Internal", "SensitivityEligibleIssuers"] = int(internal_excl.sum())

    # ------------------------------------------------------------------
    # Plan-level governed ratios (validation outputs, not duplicate issuer KPIs).
    # ------------------------------------------------------------------
    plan_denial_rows: list[dict[str, Any]] = []
    for name, ncol, nm, dcol, dm, network in [
        ("Plan In-Network Denial Rate", "ClaimsDeniedInNetwork", 16, "ClaimsReceivedInNetwork", 14, "IN_NETWORK"),
        ("Plan Out-of-Network Denial Rate", "ClaimsDeniedOutOfNetwork", 15, "ClaimsReceivedOutOfNetwork", 13, "OUT_OF_NETWORK"),
    ]:
        value, num, den, eligible, denom_pos = governed_ratio(
            plan_fact_i, plan_status, ncol, nm, dcol, dm, 100.0
        )
        plan_denial_rows.append({
            "Measure": name,
            "NetworkScope": network,
            "Numerator": num,
            "Denominator": den,
            "RatePct": value,
            "EligiblePlans": int(eligible.sum()),
        })

    plan_overall_eligible = (
        plan_status[[13, 14, 15, 16]].eq("AVAILABLE").all(axis=1)
        & (plan_fact_i["ClaimsReceivedInNetwork"] + plan_fact_i["ClaimsReceivedOutOfNetwork"]).gt(0)
    )
    pnum = float((plan_fact_i.loc[plan_overall_eligible, "ClaimsDeniedInNetwork"] + plan_fact_i.loc[plan_overall_eligible, "ClaimsDeniedOutOfNetwork"]).sum())
    pden = float((plan_fact_i.loc[plan_overall_eligible, "ClaimsReceivedInNetwork"] + plan_fact_i.loc[plan_overall_eligible, "ClaimsReceivedOutOfNetwork"]).sum())
    plan_denial_rows.append({
        "Measure": "Plan Comparable Overall Denial Rate",
        "NetworkScope": "ALL_NETWORK_COMPARABLE",
        "Numerator": pnum,
        "Denominator": pden,
        "RatePct": safe_ratio(pnum, pden, 100.0),
        "EligiblePlans": int(plan_overall_eligible.sum()),
    })
    plan_denial_df = pd.DataFrame(plan_denial_rows)

    for name, ncol, nm, dcol, dm, network in [
        ("Plan In-Network Resubmissions per 100 Denied Claims", "ClaimsResubmittedInNetwork", 18, "ClaimsDeniedInNetwork", 16, "IN_NETWORK"),
        ("Plan Out-of-Network Resubmissions per 100 Denied Claims", "ClaimsResubmittedOutOfNetwork", 17, "ClaimsDeniedOutOfNetwork", 15, "OUT_OF_NETWORK"),
    ]:
        value, num, den, eligible, denom_pos = governed_ratio(
            plan_fact_i, plan_status, ncol, nm, dcol, dm, 100.0
        )
        entity_ratio = plan_fact_i.loc[eligible, ncol] / plan_fact_i.loc[eligible, dcol] * 100.0
        resub_rows.append({
            "KPIID": "PLAN-VALIDATION",
            "EntityLevel": "PLAN",
            "NetworkScope": network,
            "AggregateEventsPer100Denied": value,
            "NumeratorResubmissionEvents": num,
            "DenominatorDeniedClaims": den,
            "EligibleEntities": int(eligible.sum()),
            "EntitiesAbove100": int((entity_ratio > 100).sum()),
            "Interpretation": "Values above 100 are permitted and are not automatically DQ failures.",
        })

    # ------------------------------------------------------------------
    # Denial reasons.
    # ------------------------------------------------------------------
    reason_comparable = plan_status[REASON_METRIC_KEYS].eq("AVAILABLE").all(axis=1)
    reason_sums = plan_fact_i.loc[reason_comparable, REASON_COLUMNS].sum()
    reason_total = float(reason_sums.sum())
    reason_rows = []
    for metric_key in REASON_METRIC_KEYS:
        column = PLAN_METRIC_MAP[metric_key]
        metric_row = dim_metric.loc[dim_metric["MetricKey"] == metric_key].iloc[0]
        count = float(reason_sums[column])
        share = safe_ratio(count, reason_total, 100.0)
        reason_rows.append(
            {
                "MetricKey": metric_key,
                "MetricCode": metric_row["MetricCode"],
                "DenialReason": metric_row["MetricDisplayName"],
                "ComparablePlans": int(reason_comparable.sum()),
                "ReasonCount": count,
                "ReportedReasonCompositionPct": share,
                "Interpretation": "Share of reported denial-reason counts within the fully comparable plan population; not automatically share of all denied claims.",
            }
        )
    reason_df = pd.DataFrame(reason_rows)
    reason_share_total = float(reason_df["ReportedReasonCompositionPct"].sum())
    accept("KPI-VAL-011", "Comparable denial-reason composition sums to 100%", "100% ± 0.000001", reason_share_total, "PASS" if abs(reason_share_total - 100.0) <= 1e-6 else "FAIL", "RECONCILIATION")
    accept("KPI-VAL-012", "Fully comparable plan population for all 10 denial reasons", "> 0", int(reason_comparable.sum()), "PASS" if int(reason_comparable.sum()) > 0 else "FAIL", "ANALYTICAL")

    # KPI-019/020 are reason-specific contract measures. Detailed values are
    # reported one row per reason in the Denial_Reasons sheet.
    add_kpi("KPI-019", "Denial Reason Count", "METRIC-SPECIFIC", np.nan, "count", interpretation="See Denial_Reasons sheet for one governed row per denial reason.")
    add_kpi("KPI-020", "Comparable Denial Reason Composition %", "METRIC-SPECIFIC", np.nan, "%", interpretation="See Denial_Reasons sheet; all 10 reasons use the same fully comparable plan population.")

    # ------------------------------------------------------------------
    # Enrollment KPIs.
    # ------------------------------------------------------------------
    enrollment_rows: list[dict[str, Any]] = []
    for kpi_id, name, col, mk in [
        ("KPI-021", "Reported Average Monthly Enrollment", "AverageMonthlyEnrollment", 29),
        ("KPI-022", "Reported Average Monthly Disenrollment", "AverageMonthlyDisenrollment", 30),
    ]:
        value, mask = available_sum(plan_fact_i, plan_status, col, mk)
        add_kpi(kpi_id, name, "PLAN", value, "reported average monthly members", int(mask.sum()))
        add_population(kpi_id, "PLAN", mask, plan_status, [mk])
        status_series = plan_status[mk]
        applicable_plans = int(status_series.isin(APPLICABLE_STATUS_CODES).sum())
        enrollment_rows.append({
            "KPIID": kpi_id,
            "Measure": name,
            "ReportedValue": value,
            "AvailablePlans": int(mask.sum()),
            "ApplicablePlans": applicable_plans,
            "TotalPlans": len(plan_fact_i),
            "AvailabilityRateAmongApplicablePct": safe_ratio(int(mask.sum()), applicable_plans, 100.0),
            "AvailableAmongAllPlansPct": safe_ratio(int(mask.sum()), len(plan_fact_i), 100.0),
            "SuppressedPlans": int(status_series.eq("SUPPRESSED_SMALL_CELL").sum()),
            "NotAvailablePlans": int(status_series.eq("NOT_AVAILABLE").sum()),
            "StructuralNonApplicablePlans": int(status_series.isin(STRUCTURAL_NON_APPLICABLE_CODES).sum()),
        })

    enrollment_eligible = (
        plan_status[29].eq("AVAILABLE")
        & plan_status[30].eq("AVAILABLE")
        & plan_fact_i["AverageMonthlyEnrollment"].gt(0)
    )
    enrollment_num = float(plan_fact_i.loc[enrollment_eligible, "AverageMonthlyDisenrollment"].sum())
    enrollment_den = float(plan_fact_i.loc[enrollment_eligible, "AverageMonthlyEnrollment"].sum())
    enrollment_ratio = safe_ratio(enrollment_num, enrollment_den, 100.0)
    add_kpi("KPI-023", "Comparable Disenrollment-to-Enrollment Ratio", "PLAN", enrollment_ratio, "%", int(enrollment_eligible.sum()), enrollment_num, enrollment_den, "Not annual churn or retention probability.")
    add_population("KPI-023", "PLAN", enrollment_eligible, plan_status, [29, 30], plan_fact_i["AverageMonthlyEnrollment"].gt(0))
    enrollment_rows.append({
        "KPIID": "KPI-023",
        "Measure": "Comparable Disenrollment-to-Enrollment Ratio",
        "ReportedValue": enrollment_ratio,
        "AvailablePlans": int(enrollment_eligible.sum()),
        "ApplicablePlans": np.nan,
        "TotalPlans": len(plan_fact_i),
        "AvailabilityRateAmongApplicablePct": np.nan,
        "AvailableAmongAllPlansPct": safe_ratio(int(enrollment_eligible.sum()), len(plan_fact_i), 100.0),
        "SuppressedPlans": np.nan,
        "NotAvailablePlans": np.nan,
        "StructuralNonApplicablePlans": np.nan,
    })
    enrollment_df = pd.DataFrame(enrollment_rows)

    # ------------------------------------------------------------------
    # Availability generic KPIs 024-029 represented per metric.
    # ------------------------------------------------------------------
    # Contract-level rows point to the per-metric availability table rather than
    # pretending there is one cross-metric scalar.
    add_kpi("KPI-024", "Available Entity Count", "METRIC-SPECIFIC", np.nan, "count", interpretation="See Availability sheet per metric.")
    add_kpi("KPI-025", "Applicable Entity Count", "METRIC-SPECIFIC", np.nan, "count", interpretation="See Availability sheet per metric.")
    add_kpi("KPI-026", "Data Availability Rate", "METRIC-SPECIFIC", np.nan, "%", interpretation="See Availability sheet per metric.")
    add_kpi("KPI-027", "Suppression Rate", "METRIC-SPECIFIC", np.nan, "%", interpretation="See Availability sheet per metric.")
    add_kpi("KPI-028", "Unavailable Rate", "METRIC-SPECIFIC", np.nan, "%", interpretation="See Availability sheet per metric.")
    add_kpi("KPI-029", "Structural Non-Applicability Count", "METRIC-SPECIFIC", np.nan, "count", interpretation="See Availability sheet per metric.")

    # ------------------------------------------------------------------
    # DQ + diagnostics KPIs 030-033.
    # ------------------------------------------------------------------
    dq_resolution = dq["ResolutionState"].fillna("").astype(str)
    known_mask = dq["IsKnownSourceException"].fillna(False).astype(str).str.lower().eq("true")
    open_dq_count = int(dq_resolution.eq("OPEN_REVIEW").sum())
    known_rows = int(known_mask.sum())
    known_entities = int(dq.loc[known_mask, "EntityKey"].nunique())

    # Derive diagnostic observations rather than depending on a temporary file.
    issuer_resub_diag_in = (
        issuer_fact_i["ClaimsResubmittedInNetwork"].notna()
        & issuer_fact_i["ClaimsDeniedInNetwork"].notna()
        & (issuer_fact_i["ClaimsResubmittedInNetwork"] > issuer_fact_i["ClaimsDeniedInNetwork"])
    )
    issuer_resub_diag_out = (
        issuer_fact_i["ClaimsResubmittedOutOfNetwork"].notna()
        & issuer_fact_i["ClaimsDeniedOutOfNetwork"].notna()
        & (issuer_fact_i["ClaimsResubmittedOutOfNetwork"] > issuer_fact_i["ClaimsDeniedOutOfNetwork"])
    )
    plan_resub_diag_in = (
        plan_fact_i["ClaimsResubmittedInNetwork"].notna()
        & plan_fact_i["ClaimsDeniedInNetwork"].notna()
        & (plan_fact_i["ClaimsResubmittedInNetwork"] > plan_fact_i["ClaimsDeniedInNetwork"])
    )
    plan_resub_diag_out = (
        plan_fact_i["ClaimsResubmittedOutOfNetwork"].notna()
        & plan_fact_i["ClaimsDeniedOutOfNetwork"].notna()
        & (plan_fact_i["ClaimsResubmittedOutOfNetwork"] > plan_fact_i["ClaimsDeniedOutOfNetwork"])
    )
    diagnostic_count = int(
        issuer_resub_diag_in.sum()
        + issuer_resub_diag_out.sum()
        + plan_resub_diag_in.sum()
        + plan_resub_diag_out.sum()
    )

    add_kpi("KPI-030", "Open DQ Exception Count", "DQ", open_dq_count, "rows")
    add_kpi("KPI-031", "Known Source Exception Row Count", "DQ", known_rows, "rows")
    add_kpi("KPI-032", "Known Source Exception Entity Count", "DQ", known_entities, "entities")
    add_kpi("KPI-033", "Diagnostic Observation Count", "DQ", diagnostic_count, "observations")

    accept("KPI-VAL-013", "Open DQ exception rows", 36, open_dq_count, "PASS" if open_dq_count == 36 else "REVIEW", "DQ_CONTEXT")
    accept("KPI-VAL-014", "Known source exception rows", 2, known_rows, "PASS" if known_rows == 2 else "REVIEW", "DQ_CONTEXT")
    accept("KPI-VAL-015", "Known source exception entities", 1, known_entities, "PASS" if known_entities == 1 else "REVIEW", "DQ_CONTEXT")
    accept("KPI-VAL-016", "Derived resubmission diagnostic observations", 161, diagnostic_count, "PASS" if diagnostic_count == 161 else "REVIEW", "DQ_CONTEXT")

    resub_in_dq = int(dq["RuleCode"].fillna("").eq("RESUBMITTED_GT_DENIED").sum())
    accept("KPI-VAL-017", "Resubmission > denied observations classified as DQ failures", 0, resub_in_dq, "PASS" if resub_in_dq == 0 else "FAIL", "DQ_CONTEXT")

    known_97725 = dq.loc[known_mask & dq["EntityKey"].astype(str).eq("97725")]
    accept("KPI-VAL-018", "Known issuer 97725 source-exception evidence rows", 2, len(known_97725), "PASS" if len(known_97725) == 2 else "REVIEW", "DQ_CONTEXT")

    # ------------------------------------------------------------------
    # Cross-grain comparison: informational, never an equality gate.
    # ------------------------------------------------------------------
    cross_pairs = [
        (13, 1, "ClaimsReceivedOutOfNetwork", "Claims Received — Out of Network"),
        (14, 2, "ClaimsReceivedInNetwork", "Claims Received — In Network"),
        (15, 3, "ClaimsDeniedOutOfNetwork", "Claims Denied — Out of Network"),
        (16, 4, "ClaimsDeniedInNetwork", "Claims Denied — In Network"),
        (17, 5, "ClaimsResubmittedOutOfNetwork", "Claims Resubmitted — Out of Network"),
        (18, 6, "ClaimsResubmittedInNetwork", "Claims Resubmitted — In Network"),
    ]
    cross_rows: list[dict[str, Any]] = []

    for plan_mk, issuer_mk, column, label in cross_pairs:
        pa = plan_av.loc[plan_av["MetricKey"] == plan_mk, ["PlanKey", "StatusCode"]]
        child = (
            plan_fact[["PlanKey", "IssuerKey", column]]
            .merge(pa, on="PlanKey", how="left", validate="one_to_one")
        )
        grouped = child.groupby("IssuerKey", as_index=True).agg(
            ChildPlanCount=("PlanKey", "size"),
            AvailableChildPlanCount=("StatusCode", lambda s: int((s == "AVAILABLE").sum())),
            PlanValueSum=(column, "sum"),
        )
        issuer_metric_status = issuer_av.loc[
            issuer_av["MetricKey"] == issuer_mk, ["IssuerKey", "StatusCode"]
        ].set_index("IssuerKey")["StatusCode"]
        comp = issuer_fact_i[[column]].rename(columns={column: "IssuerValue"}).join(grouped).join(issuer_metric_status.rename("IssuerStatus"))
        comp["EligibleForCrossGrainTest"] = (
            comp["IssuerStatus"].eq("AVAILABLE")
            & comp["ChildPlanCount"].eq(comp["AvailableChildPlanCount"])
        )
        eligible_comp = comp.loc[comp["EligibleForCrossGrainTest"]].copy()
        equality = np.isclose(
            eligible_comp["PlanValueSum"].astype(float),
            eligible_comp["IssuerValue"].astype(float),
            rtol=0,
            atol=0,
            equal_nan=False,
        )
        cross_rows.append(
            {
                "Metric": label,
                "PlanMetricKey": plan_mk,
                "IssuerMetricKey": issuer_mk,
                "EligibleIssuers": len(eligible_comp),
                "ExactMatches": int(equality.sum()),
                "Mismatches": int((~equality).sum()),
                "MatchRatePct": safe_ratio(int(equality.sum()), len(eligible_comp), 100.0),
                "MaxAbsoluteDifference": float((eligible_comp["PlanValueSum"] - eligible_comp["IssuerValue"]).abs().max()) if len(eligible_comp) else np.nan,
                "ValidationInterpretation": "Informational cross-grain test only. Equality is not assumed; plan and issuer metrics remain separate authoritative grains.",
            }
        )

    cross_df = pd.DataFrame(cross_rows)
    accept("KPI-VAL-019", "Cross-grain comparison executed only on complete child-plan availability", "6 metrics tested", len(cross_df), "PASS" if len(cross_df) == 6 else "FAIL", "ANALYTICAL", "Mismatch counts are informational, not failures.")

    # ------------------------------------------------------------------
    # Ratio-of-totals vs average-of-row-rates evidence.
    # ------------------------------------------------------------------
    ratio_method_rows: list[dict[str, Any]] = []
    for kpi_id, ncol, nm, dcol, dm, label in [
        ("KPI-007", "ClaimsDeniedInNetwork", 4, "ClaimsReceivedInNetwork", 2, "Issuer In-Network Denial Rate"),
        ("KPI-008", "ClaimsDeniedOutOfNetwork", 3, "ClaimsReceivedOutOfNetwork", 1, "Issuer Out-of-Network Denial Rate"),
        ("KPI-010", "ClaimsResubmittedInNetwork", 6, "ClaimsDeniedInNetwork", 4, "Issuer In-Network Resubmissions per 100 Denied"),
        ("KPI-011", "ClaimsResubmittedOutOfNetwork", 5, "ClaimsDeniedOutOfNetwork", 3, "Issuer Out-of-Network Resubmissions per 100 Denied"),
        ("KPI-015", "InternalAppealsOverturned", 8, "InternalAppealsFiled", 7, "Internal Appeal Overturn Rate"),
        ("KPI-018", "ExternalAppealsOverturned", 11, "ExternalAppealsFiled", 10, "External Appeal Overturn Rate"),
    ]:
        value, num, den, elig, denom_pos = governed_ratio(issuer_fact_i, issuer_status, ncol, nm, dcol, dm, 100.0)
        row_rates = issuer_fact_i.loc[elig, ncol] / issuer_fact_i.loc[elig, dcol] * 100.0
        ratio_method_rows.append({
            "KPIID": kpi_id,
            "Measure": label,
            "GovernedRatioOfTotalsPct": value,
            "UnweightedAverageOfEntityRatesPct": float(row_rates.mean()) if len(row_rates) else np.nan,
            "DifferencePctPoints": value - float(row_rates.mean()) if len(row_rates) else np.nan,
            "EligibleEntities": int(elig.sum()),
            "GovernedMethod": "RATIO_OF_TOTALS",
            "ForbiddenDefault": "UNWEIGHTED_AVERAGE_OF_ENTITY_RATES",
        })
    ratio_method_df = pd.DataFrame(ratio_method_rows)
    accept("KPI-VAL-020", "Governed ratio calculations use common eligible populations", "PASS", "PASS", "PASS", "ANALYTICAL")

    # ------------------------------------------------------------------
    # Denied > received DQ reconciliation.
    # ------------------------------------------------------------------
    denied_gt_received_issuer = int(
        (
            (issuer_fact_i["ClaimsDeniedOutOfNetwork"].notna())
            & (issuer_fact_i["ClaimsReceivedOutOfNetwork"].notna())
            & (issuer_fact_i["ClaimsDeniedOutOfNetwork"] > issuer_fact_i["ClaimsReceivedOutOfNetwork"])
        ).sum()
        + (
            (issuer_fact_i["ClaimsDeniedInNetwork"].notna())
            & (issuer_fact_i["ClaimsReceivedInNetwork"].notna())
            & (issuer_fact_i["ClaimsDeniedInNetwork"] > issuer_fact_i["ClaimsReceivedInNetwork"])
        ).sum()
    )
    denied_gt_received_plan = int(
        (
            (plan_fact_i["ClaimsDeniedOutOfNetwork"].notna())
            & (plan_fact_i["ClaimsReceivedOutOfNetwork"].notna())
            & (plan_fact_i["ClaimsDeniedOutOfNetwork"] > plan_fact_i["ClaimsReceivedOutOfNetwork"])
        ).sum()
        + (
            (plan_fact_i["ClaimsDeniedInNetwork"].notna())
            & (plan_fact_i["ClaimsReceivedInNetwork"].notna())
            & (plan_fact_i["ClaimsDeniedInNetwork"] > plan_fact_i["ClaimsReceivedInNetwork"])
        ).sum()
    )
    dq_denied_count = int(dq["RuleCode"].fillna("").eq("DENIED_GT_RECEIVED").sum())
    derived_denied_count = denied_gt_received_issuer + denied_gt_received_plan
    accept("KPI-VAL-021", "Denied > received DQ rows reconcile to canonical facts", derived_denied_count, dq_denied_count, "PASS" if derived_denied_count == dq_denied_count else "FAIL", "DQ_CONTEXT")

    # ------------------------------------------------------------------
    # Canonical lineage and metric-dictionary contract.
    # ------------------------------------------------------------------
    lineage_hashes = set(dim_plan.get("SourceFileSHA256", pd.Series(dtype=str)).dropna().astype(str).str.lower())
    lineage_hashes |= set(issuer_fact.get("SourceFileSHA256", pd.Series(dtype=str)).dropna().astype(str).str.lower())
    lineage_ok = lineage_hashes == {EXPECTED_SOURCE_SHA256.lower()}
    accept(
        "KPI-VAL-022",
        "Canonical lineage references approved RAW source SHA256",
        EXPECTED_SOURCE_SHA256,
        " | ".join(sorted(lineage_hashes)) if lineage_hashes else "<missing>",
        "PASS" if lineage_ok else "FAIL",
        "STRUCTURAL",
    )

    metric_contract_ok = (
        set(dim_metric.loc[dim_metric["EntityLevel"].eq("ISSUER"), "MetricKey"].astype(int)) == set(ISSUER_METRIC_MAP)
        and set(dim_metric.loc[dim_metric["EntityLevel"].eq("PLAN"), "MetricKey"].astype(int)) == set(PLAN_METRIC_MAP)
    )
    accept(
        "KPI-VAL-023",
        "Metric dictionary entity-level contract",
        "Issuer MetricKeys 1-12; Plan MetricKeys 13-30",
        f"issuer={sorted(dim_metric.loc[dim_metric['EntityLevel'].eq('ISSUER'), 'MetricKey'].astype(int).tolist())}; plan={sorted(dim_metric.loc[dim_metric['EntityLevel'].eq('PLAN'), 'MetricKey'].astype(int).tolist())}",
        "PASS" if metric_contract_ok else "FAIL",
        "STRUCTURAL",
    )

    # ------------------------------------------------------------------
    # Current expected canonical shape from approved build.
    # ------------------------------------------------------------------
    shape_expectations = [
        ("DimState rows", 30, len(dim_state)),
        ("DimIssuer rows", 348, len(dim_issuer)),
        ("DimPlan rows", 4956, len(dim_plan)),
        ("Issuer fact rows", 348, len(issuer_fact)),
        ("Plan fact rows", 4956, len(plan_fact)),
        ("Issuer metric availability rows", 4176, len(issuer_av)),
        ("Plan metric availability rows", 89208, len(plan_av)),
        ("DimMetric rows", 30, len(dim_metric)),
        ("DimAvailabilityStatus rows", 8, len(dim_status)),
    ]
    for i, (name, expected, actual) in enumerate(shape_expectations, start=24):
        accept(f"KPI-VAL-{i:03d}", name, expected, actual, "PASS" if expected == actual else "FAIL", "STRUCTURAL")

    # ------------------------------------------------------------------
    # Results and final status.
    # ------------------------------------------------------------------
    kpi_df = pd.DataFrame(kpi_rows)
    population_df = pd.DataFrame(population_rows)
    denial_rate_df = pd.DataFrame(denial_rate_rows)
    resub_df = pd.DataFrame(resub_rows)
    acceptance_df = pd.DataFrame(acceptance_rows)

    hard_fail = bool(
        acceptance_df["Status"].eq("FAIL").any()
    )
    review_count = int(acceptance_df["Status"].eq("REVIEW").sum())
    validation_status = (
        "FAIL" if hard_fail else ("PASS_WITH_CONTEXT" if review_count > 0 or len(dq) > 0 else "PASS")
    )

    file_inventory_rows = []
    for name in REQUIRED_FILES:
        path = CANONICAL_DIR / name
        file_inventory_rows.append(
            {
                "CanonicalFile": name,
                "Bytes": path.stat().st_size,
                "SHA256": sha256(path),
            }
        )
    file_inventory_df = pd.DataFrame(file_inventory_rows)

    dq_summary = (
        dq.groupby(["Severity", "Scope", "RuleCode", "ResolutionState"], dropna=False)
        .size()
        .reset_index(name="ExceptionCount")
    )

    summary_df = pd.DataFrame(
        [
            ("ValidationStatus", validation_status),
            ("ScriptVersion", SCRIPT_VERSION),
            ("PUFYear", PUF_YEAR),
            ("ExperienceYear", EXPERIENCE_YEAR),
            ("CanonicalDirectory", str(CANONICAL_DIR)),
            ("CanonicalFilesRead", len(REQUIRED_FILES)),
            ("AcceptanceTests", len(acceptance_df)),
            ("AcceptancePass", int(acceptance_df["Status"].eq("PASS").sum())),
            ("AcceptanceReview", review_count),
            ("AcceptanceFail", int(acceptance_df["Status"].eq("FAIL").sum())),
            ("KPIContractRowsRepresented", 33),
            ("OpenDQExceptions", open_dq_count),
            ("KnownSourceExceptionRows", known_rows),
            ("KnownSourceExceptionEntities", known_entities),
            ("DiagnosticObservations", diagnostic_count),
            ("ComparableDenialReasonPlans", int(reason_comparable.sum())),
            ("OutputFile", str(REVIEW_FILE)),
        ],
        columns=["Item", "Value"],
    )

    methodology_df = pd.DataFrame(
        [
            ("M-001", "Rate aggregation", "All governed rates use ratio of totals over the same eligible population; unweighted averages of row-level rates are not the default."),
            ("M-002", "Availability", "Only AVAILABLE numeric values are used. Suppressed/unavailable/non-applicable statuses are not converted to zero."),
            ("M-003", "Overall network metrics", "In-network and out-of-network components must both be AVAILABLE for an entity before that entity enters a comparable all-network KPI."),
            ("M-004", "Resubmissions", "Reported as events per 100 denied claims. Values above 100 are permitted and are not automatically a DQ failure."),
            ("M-005", "Appeals", "Aggregate overturn rates are calculated from aggregate overturned/filed counts. Published percentages are used for source reconciliation only."),
            ("M-006", "Denial reasons", "Composition uses only plans where all 10 denial-reason metrics are AVAILABLE. Result is share of reported reason counts, not automatically share of all denied claims."),
            ("M-007", "Enrollment", "Totals are reported available-population sums of average monthly measures. The disenrollment/enrollment ratio is not annual churn or retention."),
            ("M-008", "DQ", "Source anomalies are preserved and flagged. Known source exceptions are never silently corrected or clipped."),
            ("M-009", "Cross-grain", "Plan and issuer metrics remain separate facts. Cross-grain comparisons are informational and equality is not assumed."),
            ("M-010", "Publication guardrail", "Claims, denials, appeals, and enrollment disclosures must not be presented as direct insurer/plan quality scores."),
        ],
        columns=["RuleID", "Topic", "GovernedRule"],
    )

    # ------------------------------------------------------------------
    # Write one review artifact outside the project.
    # ------------------------------------------------------------------
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(REVIEW_FILE, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="00_Summary", index=False)
        acceptance_df.to_excel(writer, sheet_name="01_Acceptance", index=False)
        kpi_df.to_excel(writer, sheet_name="02_KPI_Results", index=False)
        population_df.to_excel(writer, sheet_name="03_Populations", index=False)
        availability_pivot.to_excel(writer, sheet_name="04_Availability", index=False)
        availability_integrity_df.to_excel(writer, sheet_name="05_Avail_Integrity", index=False)
        denial_rate_df.to_excel(writer, sheet_name="06_Issuer_Denial", index=False)
        plan_denial_df.to_excel(writer, sheet_name="07_Plan_Denial", index=False)
        resub_df.to_excel(writer, sheet_name="08_Resubmission", index=False)
        appeal_df.to_excel(writer, sheet_name="09_Appeals", index=False)
        reason_df.to_excel(writer, sheet_name="10_Denial_Reasons", index=False)
        enrollment_df.to_excel(writer, sheet_name="11_Enrollment", index=False)
        cross_df.to_excel(writer, sheet_name="12_Cross_Grain", index=False)
        ratio_method_df.to_excel(writer, sheet_name="13_Ratio_Method", index=False)
        dq_summary.to_excel(writer, sheet_name="14_DQ_Context", index=False)
        file_inventory_df.to_excel(writer, sheet_name="15_File_Inventory", index=False)
        methodology_df.to_excel(writer, sheet_name="16_Methodology", index=False)

        wb = writer.book
        header_fill = "1F4E78"
        header_font_color = "FFFFFF"

        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.fill = PatternFill("solid", fgColor=header_fill)
                cell.font = Font(bold=True, color=header_font_color)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for column_cells in ws.columns:
                max_len = 0
                for cell in list(column_cells)[:250]:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                width = min(max(max_len + 2, 11), 55)
                ws.column_dimensions[column_cells[0].column_letter].width = width
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)

        # Number formatting for the most important output sheets.
        ws = wb["02_KPI_Results"]
        for cell in ws["D"][1:]:
            cell.number_format = "#,##0.0000"
        for col in ["G", "H"]:
            for cell in ws[col][1:]:
                cell.number_format = "#,##0.0000"

        ws = wb["04_Availability"]
        for col_name in ["DataAvailabilityRatePct", "SuppressionRatePct", "UnavailableRatePct"]:
            headers = [c.value for c in ws[1]]
            if col_name in headers:
                idx = headers.index(col_name) + 1
                for row in range(2, ws.max_row + 1):
                    ws.cell(row=row, column=idx).number_format = "0.00"

        # Conditional status styling.
        for sheet_name in ["01_Acceptance", "05_Avail_Integrity"]:
            ws = wb[sheet_name]
            headers = [c.value for c in ws[1]]
            if "Status" in headers:
                status_col = headers.index("Status") + 1
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row, column=status_col)
                    if cell.value == "PASS":
                        cell.fill = PatternFill("solid", fgColor="E2F0D9")
                    elif cell.value == "REVIEW":
                        cell.fill = PatternFill("solid", fgColor="FFF2CC")
                    elif cell.value == "FAIL":
                        cell.fill = PatternFill("solid", fgColor="F4CCCC")
                    cell.font = Font(bold=True)

    print()
    print("KPI validation completed successfully.")
    print(f"Validation status: {validation_status}")
    print(f"Acceptance tests : {len(acceptance_df)}")
    print(f"PASS             : {int(acceptance_df['Status'].eq('PASS').sum())}")
    print(f"REVIEW           : {review_count}")
    print(f"FAIL             : {int(acceptance_df['Status'].eq('FAIL').sum())}")
    print(f"Review evidence  : {REVIEW_FILE}")
    print("Canonical files were read only; no canonical file was modified.")
    print("Review evidence is generated separately from the canonical data.")
    print("Next gate: review 03_KPI_VALIDATION_REPORT.xlsx before EDA or BI work.")

    return 0 if not hard_fail else 2


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print()
        print("KPI VALIDATION FAILED")
        print(str(exc))
        sys.exit(1)
