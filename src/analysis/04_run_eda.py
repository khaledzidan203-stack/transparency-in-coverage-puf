from __future__ import annotations

from pathlib import Path
from typing import Iterable
import hashlib
import os
import sys

import numpy as np
import pandas as pd


SCRIPT_VERSION = "1.0.0"
PUF_YEAR = 2026
EXPERIENCE_YEAR = 2024

PROJECT_ROOT = Path(__file__).resolve().parents[2]
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
REVIEW_FILE = REVIEW_DIR / "04_EDA_REPORT.xlsx"

FILES = {
    "period": "dim_reporting_period.csv",
    "state": "dim_state.csv",
    "issuer": "dim_issuer.csv",
    "plan": "dim_plan.csv",
    "metric": "dim_metric.csv",
    "status": "dim_availability_status.csv",
    "issuer_fact": "fact_issuer_transparency.csv",
    "plan_fact": "fact_plan_transparency.csv",
    "issuer_avail": "fact_issuer_metric_availability.csv",
    "plan_avail": "fact_plan_metric_availability.csv",
    "dq": "dq_exception_register.csv",
}

ISSUER_METRIC_KEYS = {
    "claims_received_oon": 1,
    "claims_received_in": 2,
    "claims_denied_oon": 3,
    "claims_denied_in": 4,
    "claims_resubmitted_oon": 5,
    "claims_resubmitted_in": 6,
    "internal_appeals_filed": 7,
    "internal_appeals_overturned": 8,
    "internal_appeals_pct": 9,
    "external_appeals_filed": 10,
    "external_appeals_overturned": 11,
    "external_appeals_pct": 12,
}

PLAN_METRIC_KEYS = {
    "claims_received_oon": 13,
    "claims_received_in": 14,
    "claims_denied_oon": 15,
    "claims_denied_in": 16,
    "claims_resubmitted_oon": 17,
    "claims_resubmitted_in": 18,
    "reason_referral": 19,
    "reason_oon": 20,
    "reason_excluded": 21,
    "reason_not_med_nec_non_bh": 22,
    "reason_not_med_nec_bh": 23,
    "reason_benefit_limit": 24,
    "reason_member_not_covered": 25,
    "reason_investigational": 26,
    "reason_administrative": 27,
    "reason_other": 28,
    "avg_monthly_enrollment": 29,
    "avg_monthly_disenrollment": 30,
}

PLAN_REASON_KEYS = list(range(19, 29))

PLAN_REASON_COLUMNS = {
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
}


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def read_all() -> dict[str, pd.DataFrame]:
    missing = [
        filename
        for filename in FILES.values()
        if not (CANONICAL_DIR / filename).exists()
    ]
    if missing:
        raise FileNotFoundError(
            "Missing canonical files:\n- " + "\n- ".join(missing)
        )

    return {
        key: pd.read_csv(CANONICAL_DIR / filename, low_memory=False)
        for key, filename in FILES.items()
    }


def available_entity_keys(
    availability: pd.DataFrame,
    status_dim: pd.DataFrame,
    metric_keys: Iterable[int],
    entity_columns: list[str],
) -> pd.DataFrame:
    status_map = status_dim[["StatusKey", "StatusCode"]]
    a = availability.merge(
        status_map,
        left_on="AvailabilityStatusKey",
        right_on="StatusKey",
        how="left",
        validate="many_to_one",
    )

    keys = list(metric_keys)
    subset = a[
        a["MetricKey"].isin(keys) & (a["StatusCode"] == "AVAILABLE")
    ].copy()

    if subset.empty:
        return pd.DataFrame(columns=entity_columns)

    complete = (
        subset.groupby(entity_columns)["MetricKey"]
        .nunique()
        .reset_index(name="_metric_count")
    )
    return complete.loc[
        complete["_metric_count"] == len(keys),
        entity_columns,
    ].copy()


def ratio_of_totals(
    df: pd.DataFrame,
    numerator_col: str,
    denominator_col: str,
) -> float:
    denominator = df[denominator_col].sum()
    if pd.isna(denominator) or denominator <= 0:
        return np.nan
    return float(df[numerator_col].sum() / denominator * 100.0)


def make_rate_frame(
    fact: pd.DataFrame,
    availability: pd.DataFrame,
    status_dim: pd.DataFrame,
    entity_columns: list[str],
    required_metric_keys: list[int],
    numerator_col: str,
    denominator_col: str,
) -> tuple[pd.DataFrame, float]:
    eligible_keys = available_entity_keys(
        availability,
        status_dim,
        required_metric_keys,
        entity_columns,
    )
    eligible = fact.merge(
        eligible_keys,
        on=entity_columns,
        how="inner",
        validate="one_to_one",
    )
    eligible = eligible.loc[
        eligible[denominator_col].notna()
        & (eligible[denominator_col] > 0)
    ].copy()
    rate = ratio_of_totals(eligible, numerator_col, denominator_col)
    return eligible, rate


def add_common_dimensions(
    issuer_fact: pd.DataFrame,
    plan_fact: pd.DataFrame,
    dim_state: pd.DataFrame,
    dim_issuer: pd.DataFrame,
    dim_plan: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    state_small = dim_state[["StateKey", "StateCode", "ExchangeType"]]
    issuer_small = dim_issuer[
        ["IssuerKey", "IssuerID", "IssuerName", "StateKey"]
    ].rename(columns={"StateKey": "IssuerStateKey"})

    issuer = (
        issuer_fact.merge(
            state_small,
            on="StateKey",
            how="left",
            validate="many_to_one",
        )
        .merge(
            issuer_small[["IssuerKey", "IssuerID", "IssuerName"]],
            on="IssuerKey",
            how="left",
            validate="many_to_one",
        )
    )

    plan_small = dim_plan[
        [
            "PlanKey",
            "PlanID",
            "IssuerKey",
            "StateKey",
            "MarketSegment",
            "PlanType",
            "PlanOfferingType",
            "MetalLevel",
        ]
    ].rename(
        columns={
            "IssuerKey": "PlanIssuerKey",
            "StateKey": "PlanStateKey",
        }
    )

    plan = (
        plan_fact.merge(
            plan_small,
            on="PlanKey",
            how="left",
            validate="one_to_one",
        )
        .merge(
            issuer_small[["IssuerKey", "IssuerID", "IssuerName"]],
            on="IssuerKey",
            how="left",
            validate="many_to_one",
        )
        .merge(
            state_small,
            on="StateKey",
            how="left",
            validate="many_to_one",
        )
    )

    return issuer, plan


def grouped_rate(
    eligible: pd.DataFrame,
    group_cols: list[str],
    numerator_col: str,
    denominator_col: str,
    entity_key_col: str,
    rate_name: str,
) -> pd.DataFrame:
    grouped = (
        eligible.groupby(group_cols, dropna=False)
        .agg(
            EligibleEntities=(entity_key_col, "nunique"),
            Numerator=(numerator_col, "sum"),
            Denominator=(denominator_col, "sum"),
        )
        .reset_index()
    )
    grouped[rate_name] = np.where(
        grouped["Denominator"] > 0,
        grouped["Numerator"] / grouped["Denominator"] * 100.0,
        np.nan,
    )
    return grouped


def style_workbook(path: Path) -> None:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = load_workbook(path)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        max_scan_rows = min(ws.max_row, 250)
        for col_idx in range(1, ws.max_column + 1):
            max_len = 0
            for row_idx in range(1, max_scan_rows + 1):
                value = ws.cell(row_idx, col_idx).value
                if value is not None:
                    max_len = max(max_len, len(str(value)))
            width = min(max(max_len + 2, 10), 45)
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True,
                )

    wb.save(path)


def main() -> int:
    print("=" * 78)
    print("Transparency in Coverage PUF — Governed EDA")
    print("=" * 78)
    print(f"Script version : {SCRIPT_VERSION}")
    print(f"Canonical dir  : {CANONICAL_DIR}")
    print(f"Review file    : {REVIEW_FILE}")

    data = read_all()

    period = data["period"]
    dim_state = data["state"]
    dim_issuer = data["issuer"]
    dim_plan = data["plan"]
    dim_metric = data["metric"]
    dim_status = data["status"]
    issuer_fact = data["issuer_fact"]
    plan_fact = data["plan_fact"]
    issuer_avail = data["issuer_avail"]
    plan_avail = data["plan_avail"]
    dq = data["dq"]

    issuer, plan = add_common_dimensions(
        issuer_fact,
        plan_fact,
        dim_state,
        dim_issuer,
        dim_plan,
    )

    acceptance: list[dict] = []

    def test(test_id, name, expected, actual, status, notes=""):
        acceptance.append(
            {
                "TestID": test_id,
                "TestName": name,
                "Expected": expected,
                "Actual": actual,
                "Status": status,
                "Notes": notes,
            }
        )

    # ------------------------------------------------------------------
    # Baseline validation.
    # ------------------------------------------------------------------

    period_ok = (
        len(period) == 1
        and int(period.iloc[0]["PUFYear"]) == PUF_YEAR
        and int(period.iloc[0]["ExperienceYear"]) == EXPERIENCE_YEAR
    )
    test(
        "EDA-001",
        "Reporting period",
        "PUF 2026 / Experience 2024",
        (
            f"PUF {int(period.iloc[0]['PUFYear'])} / "
            f"Experience {int(period.iloc[0]['ExperienceYear'])}"
        ),
        "PASS" if period_ok else "FAIL",
    )
    test("EDA-002", "Issuer fact rows", 348, len(issuer_fact), "PASS" if len(issuer_fact) == 348 else "FAIL")
    test("EDA-003", "Plan fact rows", 4956, len(plan_fact), "PASS" if len(plan_fact) == 4956 else "FAIL")
    test("EDA-004", "State rows", 30, len(dim_state), "PASS" if len(dim_state) == 30 else "FAIL")
    test("EDA-005", "Issuer rows", 348, len(dim_issuer), "PASS" if len(dim_issuer) == 348 else "FAIL")
    test("EDA-006", "Plan rows", 4956, len(dim_plan), "PASS" if len(dim_plan) == 4956 else "FAIL")

    # ------------------------------------------------------------------
    # Issuer denial EDA.
    # ------------------------------------------------------------------

    issuer_in, issuer_in_rate = make_rate_frame(
        issuer,
        issuer_avail,
        dim_status,
        ["ReportingPeriodKey", "StateKey", "IssuerKey"],
        [ISSUER_METRIC_KEYS["claims_denied_in"], ISSUER_METRIC_KEYS["claims_received_in"]],
        "ClaimsDeniedInNetwork",
        "ClaimsReceivedInNetwork",
    )

    issuer_oon, issuer_oon_rate = make_rate_frame(
        issuer,
        issuer_avail,
        dim_status,
        ["ReportingPeriodKey", "StateKey", "IssuerKey"],
        [ISSUER_METRIC_KEYS["claims_denied_oon"], ISSUER_METRIC_KEYS["claims_received_oon"]],
        "ClaimsDeniedOutOfNetwork",
        "ClaimsReceivedOutOfNetwork",
    )

    issuer_all_keys = available_entity_keys(
        issuer_avail,
        dim_status,
        [1, 2, 3, 4],
        ["ReportingPeriodKey", "StateKey", "IssuerKey"],
    )
    issuer_all = issuer.merge(
        issuer_all_keys,
        on=["ReportingPeriodKey", "StateKey", "IssuerKey"],
        how="inner",
        validate="one_to_one",
    ).copy()
    issuer_all["ClaimsReceivedTotal"] = (
        issuer_all["ClaimsReceivedInNetwork"]
        + issuer_all["ClaimsReceivedOutOfNetwork"]
    )
    issuer_all["ClaimsDeniedTotal"] = (
        issuer_all["ClaimsDeniedInNetwork"]
        + issuer_all["ClaimsDeniedOutOfNetwork"]
    )
    issuer_all = issuer_all.loc[
        issuer_all["ClaimsReceivedTotal"] > 0
    ].copy()
    issuer_all["OverallDenialRatePct"] = (
        issuer_all["ClaimsDeniedTotal"]
        / issuer_all["ClaimsReceivedTotal"]
        * 100.0
    )
    issuer_all["InNetworkDenialRatePct"] = np.where(
        issuer_all["ClaimsReceivedInNetwork"] > 0,
        issuer_all["ClaimsDeniedInNetwork"]
        / issuer_all["ClaimsReceivedInNetwork"]
        * 100.0,
        np.nan,
    )
    issuer_all["OutOfNetworkDenialRatePct"] = np.where(
        issuer_all["ClaimsReceivedOutOfNetwork"] > 0,
        issuer_all["ClaimsDeniedOutOfNetwork"]
        / issuer_all["ClaimsReceivedOutOfNetwork"]
        * 100.0,
        np.nan,
    )
    issuer_all["NetworkDenialRateGapPctPoints"] = (
        issuer_all["OutOfNetworkDenialRatePct"]
        - issuer_all["InNetworkDenialRatePct"]
    )

    issuer_overall_rate = ratio_of_totals(
        issuer_all,
        "ClaimsDeniedTotal",
        "ClaimsReceivedTotal",
    )

    issuer_state_in = grouped_rate(
        issuer_in,
        ["StateCode", "ExchangeType"],
        "ClaimsDeniedInNetwork",
        "ClaimsReceivedInNetwork",
        "IssuerKey",
        "InNetworkDenialRatePct",
    )
    issuer_state_oon = grouped_rate(
        issuer_oon,
        ["StateCode", "ExchangeType"],
        "ClaimsDeniedOutOfNetwork",
        "ClaimsReceivedOutOfNetwork",
        "IssuerKey",
        "OutOfNetworkDenialRatePct",
    )
    issuer_state_all = grouped_rate(
        issuer_all,
        ["StateCode", "ExchangeType"],
        "ClaimsDeniedTotal",
        "ClaimsReceivedTotal",
        "IssuerKey",
        "OverallDenialRatePct",
    )
    issuer_state = (
        issuer_state_all
        .merge(
            issuer_state_in[
                [
                    "StateCode",
                    "ExchangeType",
                    "EligibleEntities",
                    "InNetworkDenialRatePct",
                ]
            ].rename(columns={"EligibleEntities": "EligibleIssuersInNetwork"}),
            on=["StateCode", "ExchangeType"],
            how="outer",
        )
        .merge(
            issuer_state_oon[
                [
                    "StateCode",
                    "ExchangeType",
                    "EligibleEntities",
                    "OutOfNetworkDenialRatePct",
                ]
            ].rename(columns={"EligibleEntities": "EligibleIssuersOutOfNetwork"}),
            on=["StateCode", "ExchangeType"],
            how="outer",
        )
        .rename(columns={"EligibleEntities": "EligibleIssuersOverall"})
    )
    issuer_state["NetworkGapPctPoints"] = (
        issuer_state["OutOfNetworkDenialRatePct"]
        - issuer_state["InNetworkDenialRatePct"]
    )
    issuer_state = issuer_state.sort_values(
        ["OverallDenialRatePct", "StateCode"],
        ascending=[False, True],
    )

    issuer_detail = issuer_all[
        [
            "IssuerID",
            "IssuerName",
            "StateCode",
            "ExchangeType",
            "ClaimsReceivedTotal",
            "ClaimsDeniedTotal",
            "OverallDenialRatePct",
            "ClaimsReceivedInNetwork",
            "ClaimsDeniedInNetwork",
            "InNetworkDenialRatePct",
            "ClaimsReceivedOutOfNetwork",
            "ClaimsDeniedOutOfNetwork",
            "OutOfNetworkDenialRatePct",
            "NetworkDenialRateGapPctPoints",
        ]
    ].copy()

    dq_issuer_keys = set(
        dq.loc[dq["Scope"] == "ISSUER", "EntityKey"].astype(str)
    )
    issuer_detail["HasDQException"] = (
        issuer_detail["IssuerID"].astype(str).isin(dq_issuer_keys)
    )
    issuer_detail = issuer_detail.sort_values(
        ["ClaimsReceivedTotal", "IssuerID"],
        ascending=[False, True],
    )

    network_gap = issuer_detail.loc[
        issuer_detail["InNetworkDenialRatePct"].notna()
        & issuer_detail["OutOfNetworkDenialRatePct"].notna()
    ].copy()
    network_gap["AbsoluteGapPctPoints"] = (
        network_gap["NetworkDenialRateGapPctPoints"].abs()
    )
    network_gap = network_gap.sort_values(
        ["AbsoluteGapPctPoints", "ClaimsReceivedTotal"],
        ascending=[False, False],
    )

    # ------------------------------------------------------------------
    # Plan denial EDA.
    # ------------------------------------------------------------------

    plan_in, plan_in_rate = make_rate_frame(
        plan,
        plan_avail,
        dim_status,
        ["ReportingPeriodKey", "PlanKey"],
        [16, 14],
        "ClaimsDeniedInNetwork",
        "ClaimsReceivedInNetwork",
    )

    plan_oon, plan_oon_rate = make_rate_frame(
        plan,
        plan_avail,
        dim_status,
        ["ReportingPeriodKey", "PlanKey"],
        [15, 13],
        "ClaimsDeniedOutOfNetwork",
        "ClaimsReceivedOutOfNetwork",
    )

    plan_all_keys = available_entity_keys(
        plan_avail,
        dim_status,
        [13, 14, 15, 16],
        ["ReportingPeriodKey", "PlanKey"],
    )
    plan_all = plan.merge(
        plan_all_keys,
        on=["ReportingPeriodKey", "PlanKey"],
        how="inner",
        validate="one_to_one",
    ).copy()
    plan_all["ClaimsReceivedTotal"] = (
        plan_all["ClaimsReceivedInNetwork"]
        + plan_all["ClaimsReceivedOutOfNetwork"]
    )
    plan_all["ClaimsDeniedTotal"] = (
        plan_all["ClaimsDeniedInNetwork"]
        + plan_all["ClaimsDeniedOutOfNetwork"]
    )
    plan_all = plan_all.loc[plan_all["ClaimsReceivedTotal"] > 0].copy()
    plan_all["OverallDenialRatePct"] = (
        plan_all["ClaimsDeniedTotal"]
        / plan_all["ClaimsReceivedTotal"]
        * 100.0
    )
    plan_overall_rate = ratio_of_totals(
        plan_all,
        "ClaimsDeniedTotal",
        "ClaimsReceivedTotal",
    )

    plan_detail = plan_all[
        [
            "PlanID",
            "IssuerID",
            "IssuerName",
            "StateCode",
            "MarketSegment",
            "PlanOfferingType",
            "PlanType",
            "MetalLevel",
            "ClaimsReceivedTotal",
            "ClaimsDeniedTotal",
            "OverallDenialRatePct",
        ]
    ].copy()
    dq_plan_keys = set(
        dq.loc[dq["Scope"] == "PLAN", "EntityKey"].astype(str)
    )
    plan_detail["HasDQException"] = (
        plan_detail["PlanID"].astype(str).isin(dq_plan_keys)
    )
    plan_detail = plan_detail.sort_values(
        ["ClaimsReceivedTotal", "PlanID"],
        ascending=[False, True],
    )

    # ------------------------------------------------------------------
    # Segment-level EDA using plan comparable population.
    # ------------------------------------------------------------------

    segment_tables = []
    for label, col in [
        ("MarketSegment", "MarketSegment"),
        ("PlanOfferingType", "PlanOfferingType"),
        ("PlanType", "PlanType"),
        ("MetalLevel", "MetalLevel"),
        ("State", "StateCode"),
    ]:
        table = grouped_rate(
            plan_all,
            [col],
            "ClaimsDeniedTotal",
            "ClaimsReceivedTotal",
            "PlanKey",
            "OverallDenialRatePct",
        )
        table.insert(0, "Dimension", label)
        table = table.rename(columns={col: "Category"})
        segment_tables.append(table)

    segment_summary = pd.concat(segment_tables, ignore_index=True)
    segment_summary = segment_summary.sort_values(
        ["Dimension", "OverallDenialRatePct", "Category"],
        ascending=[True, False, True],
    )

    # ------------------------------------------------------------------
    # Denial reasons — fully comparable plans only.
    # ------------------------------------------------------------------

    reason_keys = available_entity_keys(
        plan_avail,
        dim_status,
        PLAN_REASON_KEYS,
        ["ReportingPeriodKey", "PlanKey"],
    )
    reason_plans = plan.merge(
        reason_keys,
        on=["ReportingPeriodKey", "PlanKey"],
        how="inner",
        validate="one_to_one",
    ).copy()

    reason_rows = []
    total_reason_counts = sum(
        float(reason_plans[col].sum())
        for col in PLAN_REASON_COLUMNS.values()
    )
    metric_lookup = dim_metric.set_index("MetricKey")

    for key, column in PLAN_REASON_COLUMNS.items():
        reason_count = float(reason_plans[column].sum())
        reason_rows.append(
            {
                "MetricKey": key,
                "MetricCode": metric_lookup.loc[key, "MetricCode"],
                "DenialReason": metric_lookup.loc[key, "MetricDisplayName"],
                "ComparablePlans": len(reason_plans),
                "ReasonCount": reason_count,
                "ReportedReasonCompositionPct": (
                    reason_count / total_reason_counts * 100.0
                    if total_reason_counts > 0
                    else np.nan
                ),
            }
        )
    denial_reasons = pd.DataFrame(reason_rows).sort_values(
        "ReasonCount",
        ascending=False,
    )

    reason_by_market_rows = []
    for market, g in reason_plans.groupby("MarketSegment", dropna=False):
        market_total = sum(
            float(g[col].sum())
            for col in PLAN_REASON_COLUMNS.values()
        )
        for key, column in PLAN_REASON_COLUMNS.items():
            reason_count = float(g[column].sum())
            reason_by_market_rows.append(
                {
                    "MarketSegment": market,
                    "MetricKey": key,
                    "DenialReason": metric_lookup.loc[key, "MetricDisplayName"],
                    "ComparablePlans": g["PlanKey"].nunique(),
                    "ReasonCount": reason_count,
                    "CompositionPct": (
                        reason_count / market_total * 100.0
                        if market_total > 0
                        else np.nan
                    ),
                }
            )
    denial_reasons_by_market = pd.DataFrame(reason_by_market_rows)

    # ------------------------------------------------------------------
    # Resubmission analysis.
    # ------------------------------------------------------------------

    resub_rows = []
    resub_configs = [
        ("ISSUER", "IN_NETWORK", issuer, issuer_avail, ["ReportingPeriodKey", "StateKey", "IssuerKey"], 6, 4, "ClaimsResubmittedInNetwork", "ClaimsDeniedInNetwork", "IssuerKey"),
        ("ISSUER", "OUT_OF_NETWORK", issuer, issuer_avail, ["ReportingPeriodKey", "StateKey", "IssuerKey"], 5, 3, "ClaimsResubmittedOutOfNetwork", "ClaimsDeniedOutOfNetwork", "IssuerKey"),
        ("PLAN", "IN_NETWORK", plan, plan_avail, ["ReportingPeriodKey", "PlanKey"], 18, 16, "ClaimsResubmittedInNetwork", "ClaimsDeniedInNetwork", "PlanKey"),
        ("PLAN", "OUT_OF_NETWORK", plan, plan_avail, ["ReportingPeriodKey", "PlanKey"], 17, 15, "ClaimsResubmittedOutOfNetwork", "ClaimsDeniedOutOfNetwork", "PlanKey"),
    ]
    for entity_level, network, fact, avail, key_cols, num_key, den_key, num_col, den_col, entity_key in resub_configs:
        eligible, rate = make_rate_frame(
            fact,
            avail,
            dim_status,
            key_cols,
            [num_key, den_key],
            num_col,
            den_col,
        )
        entity_ratio = np.where(
            eligible[den_col] > 0,
            eligible[num_col] / eligible[den_col] * 100.0,
            np.nan,
        )
        resub_rows.append(
            {
                "EntityLevel": entity_level,
                "NetworkScope": network,
                "EligibleEntities": len(eligible),
                "ResubmissionEvents": float(eligible[num_col].sum()),
                "DeniedClaims": float(eligible[den_col].sum()),
                "EventsPer100Denied": rate,
                "EntitiesAbove100": int(np.nansum(entity_ratio > 100)),
                "Interpretation": (
                    "Informational intensity measure; values above 100 "
                    "are permitted and are not automatically DQ failures."
                ),
            }
        )
    resubmission = pd.DataFrame(resub_rows)

    # ------------------------------------------------------------------
    # Appeals.
    # ------------------------------------------------------------------

    appeals_rows = []
    for appeal_type, filed_key, overturned_key, filed_col, overturned_col in [
        ("Internal", 7, 8, "InternalAppealsFiled", "InternalAppealsOverturned"),
        ("External", 10, 11, "ExternalAppealsFiled", "ExternalAppealsOverturned"),
    ]:
        eligible, rate = make_rate_frame(
            issuer,
            issuer_avail,
            dim_status,
            ["ReportingPeriodKey", "StateKey", "IssuerKey"],
            [filed_key, overturned_key],
            overturned_col,
            filed_col,
        )
        appeals_rows.append(
            {
                "AppealType": appeal_type,
                "EligibleIssuers": len(eligible),
                "AppealsFiled": float(eligible[filed_col].sum()),
                "AppealsOverturned": float(eligible[overturned_col].sum()),
                "AggregateOverturnRatePct": rate,
            }
        )
    appeals = pd.DataFrame(appeals_rows)

    internal_keys = available_entity_keys(
        issuer_avail,
        dim_status,
        [7, 8],
        ["ReportingPeriodKey", "StateKey", "IssuerKey"],
    )
    internal_detail = issuer.merge(
        internal_keys,
        on=["ReportingPeriodKey", "StateKey", "IssuerKey"],
        how="inner",
        validate="one_to_one",
    ).copy()
    internal_detail = internal_detail.loc[
        internal_detail["InternalAppealsFiled"] > 0
    ].copy()
    internal_detail["CalculatedOverturnRatePct"] = (
        internal_detail["InternalAppealsOverturned"]
        / internal_detail["InternalAppealsFiled"]
        * 100.0
    )
    internal_detail["HasDQException"] = (
        internal_detail["IssuerID"].astype(str).isin(dq_issuer_keys)
    )
    internal_detail = internal_detail[
        [
            "IssuerID",
            "IssuerName",
            "StateCode",
            "InternalAppealsFiled",
            "InternalAppealsOverturned",
            "CalculatedOverturnRatePct",
            "InternalAppealsOverturnPctPublished",
            "HasDQException",
        ]
    ].sort_values(
        ["InternalAppealsFiled", "IssuerID"],
        ascending=[False, True],
    )

    # ------------------------------------------------------------------
    # Enrollment.
    # ------------------------------------------------------------------

    enrollment_keys = available_entity_keys(
        plan_avail,
        dim_status,
        [29],
        ["ReportingPeriodKey", "PlanKey"],
    )
    disenrollment_keys = available_entity_keys(
        plan_avail,
        dim_status,
        [30],
        ["ReportingPeriodKey", "PlanKey"],
    )
    comparable_enrollment_keys = available_entity_keys(
        plan_avail,
        dim_status,
        [29, 30],
        ["ReportingPeriodKey", "PlanKey"],
    )

    enrollment = plan.merge(
        enrollment_keys,
        on=["ReportingPeriodKey", "PlanKey"],
        how="inner",
        validate="one_to_one",
    )
    disenrollment = plan.merge(
        disenrollment_keys,
        on=["ReportingPeriodKey", "PlanKey"],
        how="inner",
        validate="one_to_one",
    )
    enrollment_comp = plan.merge(
        comparable_enrollment_keys,
        on=["ReportingPeriodKey", "PlanKey"],
        how="inner",
        validate="one_to_one",
    )
    enrollment_comp = enrollment_comp.loc[
        enrollment_comp["AverageMonthlyEnrollment"] > 0
    ].copy()

    enrollment_summary = pd.DataFrame(
        [
            {
                "Measure": "Reported Average Monthly Enrollment",
                "AvailablePlans": len(enrollment),
                "ReportedValue": float(enrollment["AverageMonthlyEnrollment"].sum()),
            },
            {
                "Measure": "Reported Average Monthly Disenrollment",
                "AvailablePlans": len(disenrollment),
                "ReportedValue": float(disenrollment["AverageMonthlyDisenrollment"].sum()),
            },
            {
                "Measure": "Comparable Disenrollment-to-Enrollment Ratio",
                "AvailablePlans": len(enrollment_comp),
                "ReportedValue": ratio_of_totals(
                    enrollment_comp,
                    "AverageMonthlyDisenrollment",
                    "AverageMonthlyEnrollment",
                ),
            },
        ]
    )

    enrollment_by_market = (
        enrollment.groupby("MarketSegment", dropna=False)
        .agg(
            PlansWithEnrollment=("PlanKey", "nunique"),
            ReportedAverageMonthlyEnrollment=("AverageMonthlyEnrollment", "sum"),
        )
        .reset_index()
    )

    disenroll_by_market = (
        disenrollment.groupby("MarketSegment", dropna=False)
        .agg(
            PlansWithDisenrollment=("PlanKey", "nunique"),
            ReportedAverageMonthlyDisenrollment=("AverageMonthlyDisenrollment", "sum"),
        )
        .reset_index()
    )

    enrollment_market = enrollment_by_market.merge(
        disenroll_by_market,
        on="MarketSegment",
        how="outer",
    )

    # ------------------------------------------------------------------
    # Availability summary.
    # ------------------------------------------------------------------

    status_map = dim_status[["StatusKey", "StatusCode"]]

    issuer_avail_named = issuer_avail.merge(
        status_map,
        left_on="AvailabilityStatusKey",
        right_on="StatusKey",
        how="left",
    ).assign(EntityLevel="ISSUER")

    plan_avail_named = plan_avail.merge(
        status_map,
        left_on="AvailabilityStatusKey",
        right_on="StatusKey",
        how="left",
    ).assign(EntityLevel="PLAN")

    all_avail = pd.concat(
        [issuer_avail_named, plan_avail_named],
        ignore_index=True,
        sort=False,
    ).merge(
        dim_metric[
            ["MetricKey", "MetricCode", "MetricDisplayName"]
        ],
        on="MetricKey",
        how="left",
        validate="many_to_one",
    )

    availability = (
        all_avail.groupby(
            ["EntityLevel", "MetricKey", "MetricCode", "MetricDisplayName", "StatusCode"],
            dropna=False,
        )
        .size()
        .reset_index(name="RecordCount")
    )

    avail_pivot = (
        availability.pivot_table(
            index=["EntityLevel", "MetricKey", "MetricCode", "MetricDisplayName"],
            columns="StatusCode",
            values="RecordCount",
            fill_value=0,
            aggfunc="sum",
        )
        .reset_index()
    )
    avail_pivot.columns.name = None

    for col in [
        "AVAILABLE",
        "NOT_AVAILABLE",
        "SUPPRESSED_SMALL_CELL",
        "NOT_REQUIRED_PLAN_TYPE",
        "NOT_APPLICABLE_NEW_ENTITY",
        "SOURCE_MISSING_UNEXPECTED",
        "INVALID_NUMERIC_SOURCE",
    ]:
        if col not in avail_pivot.columns:
            avail_pivot[col] = 0

    avail_pivot["ApplicableEntities"] = (
        avail_pivot["AVAILABLE"]
        + avail_pivot["NOT_AVAILABLE"]
        + avail_pivot["SUPPRESSED_SMALL_CELL"]
        + avail_pivot["SOURCE_MISSING_UNEXPECTED"]
        + avail_pivot["INVALID_NUMERIC_SOURCE"]
    )
    avail_pivot["AvailabilityRatePct"] = np.where(
        avail_pivot["ApplicableEntities"] > 0,
        avail_pivot["AVAILABLE"] / avail_pivot["ApplicableEntities"] * 100.0,
        np.nan,
    )
    avail_pivot["SuppressionRatePct"] = np.where(
        avail_pivot["ApplicableEntities"] > 0,
        avail_pivot["SUPPRESSED_SMALL_CELL"]
        / avail_pivot["ApplicableEntities"]
        * 100.0,
        np.nan,
    )
    avail_pivot["UnavailableRatePct"] = np.where(
        avail_pivot["ApplicableEntities"] > 0,
        avail_pivot["NOT_AVAILABLE"]
        / avail_pivot["ApplicableEntities"]
        * 100.0,
        np.nan,
    )
    avail_pivot["StructuralNonApplicableCount"] = (
        avail_pivot["NOT_REQUIRED_PLAN_TYPE"]
        + avail_pivot["NOT_APPLICABLE_NEW_ENTITY"]
    )
    avail_pivot = avail_pivot.sort_values(
        ["EntityLevel", "MetricKey"]
    )

    # ------------------------------------------------------------------
    # Concentration analysis — volume concentration, not quality ranking.
    # ------------------------------------------------------------------

    issuer_concentration = issuer_all[
        [
            "IssuerID",
            "IssuerName",
            "StateCode",
            "ClaimsReceivedTotal",
            "ClaimsDeniedTotal",
        ]
    ].copy()
    issuer_concentration = issuer_concentration.sort_values(
        "ClaimsReceivedTotal",
        ascending=False,
    )
    total_claims = issuer_concentration["ClaimsReceivedTotal"].sum()
    issuer_concentration["ClaimsSharePct"] = np.where(
        total_claims > 0,
        issuer_concentration["ClaimsReceivedTotal"] / total_claims * 100.0,
        np.nan,
    )
    issuer_concentration["CumulativeClaimsSharePct"] = (
        issuer_concentration["ClaimsSharePct"].cumsum()
    )

    concentration_summary = pd.DataFrame(
        [
            {
                "Population": "Comparable issuers with all 4 claims/denial metrics available",
                "EligibleIssuers": len(issuer_concentration),
                "TotalComparableClaimsReceived": float(total_claims),
                "Top5IssuerClaimsSharePct": float(
                    issuer_concentration.head(5)["ClaimsSharePct"].sum()
                ),
                "Top10IssuerClaimsSharePct": float(
                    issuer_concentration.head(10)["ClaimsSharePct"].sum()
                ),
                "Top20IssuerClaimsSharePct": float(
                    issuer_concentration.head(20)["ClaimsSharePct"].sum()
                ),
            }
        ]
    )

    # ------------------------------------------------------------------
    # DQ context.
    # ------------------------------------------------------------------

    dq_summary = (
        dq.groupby(
            ["Severity", "Scope", "RuleCode", "ResolutionState"],
            dropna=False,
        )
        .size()
        .reset_index(name="ExceptionCount")
        .sort_values(["Scope", "RuleCode"])
    )

    # ------------------------------------------------------------------
    # Automated descriptive findings.
    # ------------------------------------------------------------------

    largest_reason = denial_reasons.iloc[0]
    min_availability = avail_pivot.loc[
        avail_pivot["AvailabilityRatePct"].idxmin()
    ]
    max_availability = avail_pivot.loc[
        avail_pivot["AvailabilityRatePct"].idxmax()
    ]

    findings = pd.DataFrame(
        [
            {
                "FindingID": "F-001",
                "Topic": "Issuer denial rates",
                "Finding": (
                    f"In-network denial rate = {issuer_in_rate:.2f}% "
                    f"across {len(issuer_in):,} eligible issuers; "
                    f"out-of-network denial rate = {issuer_oon_rate:.2f}% "
                    f"across {len(issuer_oon):,} eligible issuers."
                ),
                "InterpretationBoundary": (
                    "Descriptive reported rates only; not a quality score."
                ),
            },
            {
                "FindingID": "F-002",
                "Topic": "Comparable overall denial rate",
                "Finding": (
                    f"Comparable overall issuer denial rate = "
                    f"{issuer_overall_rate:.2f}% across "
                    f"{len(issuer_all):,} issuers with all four network "
                    f"claims/denial measures available."
                ),
                "InterpretationBoundary": (
                    "Uses a common eligible population and ratio of totals."
                ),
            },
            {
                "FindingID": "F-003",
                "Topic": "Plan comparable denial rate",
                "Finding": (
                    f"Comparable plan-level overall denial rate = "
                    f"{plan_overall_rate:.2f}% across {len(plan_all):,} "
                    f"plans with all four required measures available."
                ),
                "InterpretationBoundary": (
                    "Plan and issuer facts remain separate authoritative grains."
                ),
            },
            {
                "FindingID": "F-004",
                "Topic": "Denial reasons",
                "Finding": (
                    f"Largest reported reason-count category in the fully "
                    f"comparable {len(reason_plans):,}-plan population is "
                    f"'{largest_reason['DenialReason']}' at "
                    f"{largest_reason['ReportedReasonCompositionPct']:.2f}% "
                    f"of reported reason counts."
                ),
                "InterpretationBoundary": (
                    "Not automatically the share of all denied claims."
                ),
            },
            {
                "FindingID": "F-005",
                "Topic": "Enrollment availability",
                "Finding": (
                    f"Average monthly enrollment is numerically available "
                    f"for {len(enrollment):,} of {len(plan):,} plans; "
                    f"average monthly disenrollment for "
                    f"{len(disenrollment):,} plans."
                ),
                "InterpretationBoundary": (
                    "Reported available-population values only; suppressed "
                    "and unavailable plans are not imputed."
                ),
            },
            {
                "FindingID": "F-006",
                "Topic": "Metric availability",
                "Finding": (
                    f"Highest governed availability: "
                    f"{max_availability['MetricDisplayName']} "
                    f"({max_availability['AvailabilityRatePct']:.2f}% "
                    f"among applicable entities). Lowest: "
                    f"{min_availability['MetricDisplayName']} "
                    f"({min_availability['AvailabilityRatePct']:.2f}%)."
                ),
                "InterpretationBoundary": (
                    "Applicable denominator excludes structural non-applicability."
                ),
            },
            {
                "FindingID": "F-007",
                "Topic": "Data quality context",
                "Finding": (
                    f"Canonical DQ register contains {len(dq):,} rows: "
                    f"{int((dq['ResolutionState'] == 'OPEN_REVIEW').sum()):,} "
                    f"open-review rows and "
                    f"{int(dq['IsKnownSourceException'].fillna(False).astype(bool).sum()):,} "
                    f"known-source-exception rows."
                ),
                "InterpretationBoundary": (
                    "Source anomalies are preserved and flagged; not silently corrected."
                ),
            },
            {
                "FindingID": "F-008",
                "Topic": "Claims concentration",
                "Finding": (
                    f"Top 10 issuers by comparable reported claims volume "
                    f"account for "
                    f"{issuer_concentration.head(10)['ClaimsSharePct'].sum():.2f}% "
                    f"of claims in the comparable issuer population."
                ),
                "InterpretationBoundary": (
                    "Volume concentration only; not a performance or quality ranking."
                ),
            },
        ]
    )

    # ------------------------------------------------------------------
    # Acceptance checks.
    # ------------------------------------------------------------------

    test(
        "EDA-007",
        "Issuer in-network denial rate reproduces KPI validation",
        "18.79248% approx.",
        round(issuer_in_rate, 6),
        "PASS" if abs(issuer_in_rate - 18.79248008450311) < 1e-6 else "FAIL",
    )
    test(
        "EDA-008",
        "Issuer out-of-network denial rate reproduces KPI validation",
        "37.15701% approx.",
        round(issuer_oon_rate, 6),
        "PASS" if abs(issuer_oon_rate - 37.15701150241005) < 1e-6 else "FAIL",
    )
    test(
        "EDA-009",
        "Issuer overall denial rate reproduces KPI validation",
        "20.444815% approx.",
        round(issuer_overall_rate, 6),
        "PASS" if abs(issuer_overall_rate - 20.44481513775063) < 1e-6 else "FAIL",
    )
    test(
        "EDA-010",
        "Plan overall denial rate reproduces KPI validation",
        "20.268263% approx.",
        round(plan_overall_rate, 6),
        "PASS" if abs(plan_overall_rate - 20.26826345024872) < 1e-6 else "FAIL",
    )
    test(
        "EDA-011",
        "Fully comparable denial-reason plans",
        1089,
        len(reason_plans),
        "PASS" if len(reason_plans) == 1089 else "FAIL",
    )
    test(
        "EDA-012",
        "Denial reason composition sums to 100%",
        "100 ± 0.000001",
        float(denial_reasons["ReportedReasonCompositionPct"].sum()),
        (
            "PASS"
            if abs(denial_reasons["ReportedReasonCompositionPct"].sum() - 100.0)
            <= 1e-6
            else "FAIL"
        ),
    )
    test(
        "EDA-013",
        "Enrollment available plan count",
        2344,
        len(enrollment),
        "PASS" if len(enrollment) == 2344 else "FAIL",
    )
    test(
        "EDA-014",
        "Disenrollment available plan count",
        1340,
        len(disenrollment),
        "PASS" if len(disenrollment) == 1340 else "FAIL",
    )
    test(
        "EDA-015",
        "Open DQ rows",
        36,
        int((dq["ResolutionState"] == "OPEN_REVIEW").sum()),
        (
            "PASS"
            if int((dq["ResolutionState"] == "OPEN_REVIEW").sum()) == 36
            else "FAIL"
        ),
    )
    test(
        "EDA-016",
        "Known source exception rows",
        2,
        int(dq["IsKnownSourceException"].fillna(False).astype(bool).sum()),
        (
            "PASS"
            if int(dq["IsKnownSourceException"].fillna(False).astype(bool).sum()) == 2
            else "FAIL"
        ),
    )

    acceptance_df = pd.DataFrame(acceptance)
    fail_count = int((acceptance_df["Status"] == "FAIL").sum())
    validation_status = "PASS_WITH_CONTEXT" if fail_count == 0 else "HOLD"

    # ------------------------------------------------------------------
    # Inventory and summary.
    # ------------------------------------------------------------------

    inventory_rows = []
    for key, filename in FILES.items():
        path = CANONICAL_DIR / filename
        inventory_rows.append(
            {
                "LogicalName": key,
                "CanonicalFile": filename,
                "Bytes": path.stat().st_size,
                "SHA256": sha256(path),
            }
        )
    inventory = pd.DataFrame(inventory_rows)

    summary = pd.DataFrame(
        [
            ("EDAStatus", validation_status),
            ("ScriptVersion", SCRIPT_VERSION),
            ("PUFYear", PUF_YEAR),
            ("ExperienceYear", EXPERIENCE_YEAR),
            ("IssuerFactRows", len(issuer_fact)),
            ("PlanFactRows", len(plan_fact)),
            ("IssuerComparableOverallPopulation", len(issuer_all)),
            ("PlanComparableOverallPopulation", len(plan_all)),
            ("ComparableDenialReasonPlans", len(reason_plans)),
            ("IssuerOverallDenialRatePct", issuer_overall_rate),
            ("PlanOverallDenialRatePct", plan_overall_rate),
            ("OpenDQExceptions", int((dq["ResolutionState"] == "OPEN_REVIEW").sum())),
            ("KnownSourceExceptionRows", int(dq["IsKnownSourceException"].fillna(False).astype(bool).sum())),
            ("AcceptanceTests", len(acceptance_df)),
            ("AcceptancePass", int((acceptance_df["Status"] == "PASS").sum())),
            ("AcceptanceFail", fail_count),
            ("OutputFile", str(REVIEW_FILE)),
        ],
        columns=["Item", "Value"],
    )

    methodology = pd.DataFrame(
        [
            ("M-001", "Rates", "All rates use ratio of totals over one common eligible population."),
            ("M-002", "Availability", "Only AVAILABLE metric values enter calculations; suppressed/unavailable/non-applicable values are never zero-filled."),
            ("M-003", "Grain", "Issuer and plan facts remain separate; no fact-to-fact summation is performed."),
            ("M-004", "Network", "All-network KPIs require all relevant in-network and out-of-network components to be AVAILABLE for the same entity."),
            ("M-005", "Resubmission", "Resubmissions are events per 100 denied claims, not a bounded probability; values above 100 are informational."),
            ("M-006", "Appeals", "Overturn rates use aggregated overturned/filed counts; source percentages are non-additive."),
            ("M-007", "Denial reasons", "Reason composition uses only plans where all ten reason fields are AVAILABLE."),
            ("M-008", "Enrollment", "Enrollment/disenrollment totals are reported available-population sums; no imputation."),
            ("M-009", "DQ", "Source anomalies are preserved and disclosed; no silent clipping or correction."),
            ("M-010", "Interpretation", "Observed claims/denial/appeal patterns are not insurer or plan quality scores."),
        ],
        columns=["RuleID", "Topic", "GovernedRule"],
    )

    # ------------------------------------------------------------------
    # Write one review workbook outside the project.
    # ------------------------------------------------------------------

    REVIEW_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(REVIEW_FILE, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="00_Summary", index=False)
        acceptance_df.to_excel(writer, sheet_name="01_Acceptance", index=False)
        findings.to_excel(writer, sheet_name="02_Key_Findings", index=False)
        issuer_state.to_excel(writer, sheet_name="03_State_Denial", index=False)
        issuer_detail.to_excel(writer, sheet_name="04_Issuer_Denial", index=False)
        network_gap.to_excel(writer, sheet_name="05_Network_Gap", index=False)
        plan_detail.to_excel(writer, sheet_name="06_Plan_Denial", index=False)
        segment_summary.to_excel(writer, sheet_name="07_Segments", index=False)
        denial_reasons.to_excel(writer, sheet_name="08_Denial_Reasons", index=False)
        denial_reasons_by_market.to_excel(writer, sheet_name="09_Reasons_by_Market", index=False)
        resubmission.to_excel(writer, sheet_name="10_Resubmission", index=False)
        appeals.to_excel(writer, sheet_name="11_Appeals", index=False)
        internal_detail.to_excel(writer, sheet_name="12_Internal_Appeals", index=False)
        enrollment_summary.to_excel(writer, sheet_name="13_Enrollment", index=False)
        enrollment_market.to_excel(writer, sheet_name="14_Enrollment_Market", index=False)
        avail_pivot.to_excel(writer, sheet_name="15_Availability", index=False)
        concentration_summary.to_excel(writer, sheet_name="16_Concentration", index=False)
        issuer_concentration.to_excel(writer, sheet_name="17_Issuer_Volume", index=False)
        dq_summary.to_excel(writer, sheet_name="18_DQ_Context", index=False)
        dq.to_excel(writer, sheet_name="19_DQ_Detail", index=False)
        inventory.to_excel(writer, sheet_name="20_File_Inventory", index=False)
        methodology.to_excel(writer, sheet_name="21_Methodology", index=False)

    style_workbook(REVIEW_FILE)

    print()
    print("EDA completed successfully.")
    print(f"EDA status      : {validation_status}")
    print(f"Acceptance tests: {len(acceptance_df)}")
    print(f"PASS            : {int((acceptance_df['Status'] == 'PASS').sum())}")
    print(f"FAIL            : {fail_count}")
    print(f"Review evidence : {REVIEW_FILE}")
    print("Canonical files were read only; no canonical file was modified.")
    print("No temporary EDA files were written inside the project.")
    print("Next gate: review 04_EDA_REPORT.xlsx before BI architecture/design.")

    return 0 if fail_count == 0 else 2


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print()
        print("EDA FAILED")
        print(str(exc))
        sys.exit(1)
