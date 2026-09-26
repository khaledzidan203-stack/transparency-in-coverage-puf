from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any
import hashlib
import json
import os
import re
import sys

import numpy as np
import pandas as pd


SCRIPT_VERSION = "1.1.0"
PUF_YEAR = 2026
EXPERIENCE_YEAR = 2024

EXPECTED_SOURCE_SHA256 = (
    "27379dc76590027b0e6d21d736331f87376f0ac87810d68e9f7cca555eef4b1f"
)

EXPECTED_ANALYTICAL_SHEETS = [
    "Transparency 2026 - Ind QHP",
    "Transparency 2026 - Ind SADP",
    "Transparency 2026 - SHOP",
]

EXPECTED_COLUMNS = [
    "Individual/SHOP",
    "Exchange_Type",
    "State",
    "Issuer_Name",
    "Issuer_ID",
    "Is_Issuer_New_to_Exchange?(Yes_or_No)",
    "SADP_Only",
    "Plan_ID",
    "Plan_Type",
    "QHP or SADP?",
    "Metal_Level",
    "URL_Claims_Payment_Policies",
    "Issuer_Claims_Received_Out_of_Network",
    "Issuer_Claims_Received_In_Network",
    "Issuer_Claims_Denied_Out_of_Network",
    "Issuer_Claims_Denied_In_Network",
    "Issuer_Claims_Resubmitted_Out_of_Network",
    "Issuer_Claims_Resubmitted_In_Network",
    "Issuer_Internal_Appeals_Filed",
    "Issuer_Number_Internal_Appeals_Overturned",
    "Issuer_Percent_Internal_Appeals_Overturned",
    "Issuer_External_Appeals_Filed",
    "Issuer_Number_External_Appeals_Overturned",
    "Issuer_Percent_External_Appeals_Overturned",
    "Plan_Number_Claims_Received_Out_of_Network",
    "Plan_Number_Claims_Received_In_Network",
    "Plan_Number_Claims_Denied_Out_of_Network",
    "Plan_Number_Claims_Denied_In_Network",
    "Plan_Number_Claims_Resubmitted_Out_of_Network",
    "Plan_Number_Claims_Resubmitted_In_Network",
    "Plan_Number_Claims_Denied_Referral_Required",
    "Plan_Number_Claims_Denied_Due_To_Out_Of_Network",
    "Plan_Number_Claims_Denied_Services_Excluded",
    "Plan_Number_Claims_Denied_Not_Medically_Necessary_Excluding_Behavioral_Health",
    "Plan_Number_Claims_Denied_Not_Medically_Necessary_Behavioral_Health_Only",
    "Plan_Number_Claims_Denied_Due_To_Enrolle_Benefit_Limit_Reached",
    "Plan_Number_Claims_Denied_Due_To_Member_Not_Covered",
    "Plan_Number_Claims_Denied_Due_To_Investigational_Experimental_Cosmetic_Proceduce",
    "Plan_Number_Claims_Denied_Due_To_Administrative_Reason",
    "Plan_Number_Claims_Denied_Other",
    "Rate_Review",
    "Financial_Information",
    "Average Monthly Enrollment",
    "Average Monthly Disenrollment",
]

SPECIAL_STATUS_MAP = {
    "*": "NOT_AVAILABLE",
    "**": "SUPPRESSED_SMALL_CELL",
    "***": "NOT_REQUIRED_PLAN_TYPE",
    "N/A": "NOT_APPLICABLE_NEW_ENTITY",
}

AVAILABILITY_STATUS_ROWS = [
    (1, "AVAILABLE", True, False, False),
    (2, "NOT_AVAILABLE", False, True, False),
    (3, "SUPPRESSED_SMALL_CELL", False, True, False),
    (4, "NOT_REQUIRED_PLAN_TYPE", False, False, False),
    (5, "NOT_APPLICABLE_NEW_ENTITY", False, False, False),
    (6, "MISSING_URL", False, False, False),
    (7, "SOURCE_MISSING_UNEXPECTED", False, False, True),
    (8, "INVALID_NUMERIC_SOURCE", False, False, True),
]

STATUS_KEY_BY_CODE = {row[1]: row[0] for row in AVAILABILITY_STATUS_ROWS}


@dataclass(frozen=True)
class MetricDef:
    metric_key: int
    metric_code: str
    display_name: str
    entity_level: str
    metric_family: str
    network_scope: str
    is_percentage: bool
    is_additive: bool
    source_column: str
    value_kind: str  # COUNT | DECIMAL


METRICS = [
    MetricDef(1, "ISSUER_CLAIMS_RECEIVED_OON", "Issuer Claims Received - Out of Network", "ISSUER", "CLAIMS", "OUT_OF_NETWORK", False, True, "Issuer_Claims_Received_Out_of_Network", "COUNT"),
    MetricDef(2, "ISSUER_CLAIMS_RECEIVED_IN", "Issuer Claims Received - In Network", "ISSUER", "CLAIMS", "IN_NETWORK", False, True, "Issuer_Claims_Received_In_Network", "COUNT"),
    MetricDef(3, "ISSUER_CLAIMS_DENIED_OON", "Issuer Claims Denied - Out of Network", "ISSUER", "DENIALS", "OUT_OF_NETWORK", False, True, "Issuer_Claims_Denied_Out_of_Network", "COUNT"),
    MetricDef(4, "ISSUER_CLAIMS_DENIED_IN", "Issuer Claims Denied - In Network", "ISSUER", "DENIALS", "IN_NETWORK", False, True, "Issuer_Claims_Denied_In_Network", "COUNT"),
    MetricDef(5, "ISSUER_CLAIMS_RESUBMITTED_OON", "Issuer Claims Resubmitted - Out of Network", "ISSUER", "RESUBMISSIONS", "OUT_OF_NETWORK", False, True, "Issuer_Claims_Resubmitted_Out_of_Network", "COUNT"),
    MetricDef(6, "ISSUER_CLAIMS_RESUBMITTED_IN", "Issuer Claims Resubmitted - In Network", "ISSUER", "RESUBMISSIONS", "IN_NETWORK", False, True, "Issuer_Claims_Resubmitted_In_Network", "COUNT"),
    MetricDef(7, "ISSUER_INTERNAL_APPEALS_FILED", "Issuer Internal Appeals Filed", "ISSUER", "APPEALS", "ALL", False, True, "Issuer_Internal_Appeals_Filed", "COUNT"),
    MetricDef(8, "ISSUER_INTERNAL_APPEALS_OVERTURNED", "Issuer Internal Appeals Overturned", "ISSUER", "APPEALS", "ALL", False, True, "Issuer_Number_Internal_Appeals_Overturned", "COUNT"),
    MetricDef(9, "ISSUER_INTERNAL_APPEALS_OVERTURN_PCT", "Issuer Internal Appeals Overturn % - Published", "ISSUER", "APPEALS", "ALL", True, False, "Issuer_Percent_Internal_Appeals_Overturned", "DECIMAL"),
    MetricDef(10, "ISSUER_EXTERNAL_APPEALS_FILED", "Issuer External Appeals Filed", "ISSUER", "APPEALS", "ALL", False, True, "Issuer_External_Appeals_Filed", "COUNT"),
    MetricDef(11, "ISSUER_EXTERNAL_APPEALS_OVERTURNED", "Issuer External Appeals Overturned", "ISSUER", "APPEALS", "ALL", False, True, "Issuer_Number_External_Appeals_Overturned", "COUNT"),
    MetricDef(12, "ISSUER_EXTERNAL_APPEALS_OVERTURN_PCT", "Issuer External Appeals Overturn % - Published", "ISSUER", "APPEALS", "ALL", True, False, "Issuer_Percent_External_Appeals_Overturned", "DECIMAL"),

    MetricDef(13, "PLAN_CLAIMS_RECEIVED_OON", "Plan Claims Received - Out of Network", "PLAN", "CLAIMS", "OUT_OF_NETWORK", False, True, "Plan_Number_Claims_Received_Out_of_Network", "COUNT"),
    MetricDef(14, "PLAN_CLAIMS_RECEIVED_IN", "Plan Claims Received - In Network", "PLAN", "CLAIMS", "IN_NETWORK", False, True, "Plan_Number_Claims_Received_In_Network", "COUNT"),
    MetricDef(15, "PLAN_CLAIMS_DENIED_OON", "Plan Claims Denied - Out of Network", "PLAN", "DENIALS", "OUT_OF_NETWORK", False, True, "Plan_Number_Claims_Denied_Out_of_Network", "COUNT"),
    MetricDef(16, "PLAN_CLAIMS_DENIED_IN", "Plan Claims Denied - In Network", "PLAN", "DENIALS", "IN_NETWORK", False, True, "Plan_Number_Claims_Denied_In_Network", "COUNT"),
    MetricDef(17, "PLAN_CLAIMS_RESUBMITTED_OON", "Plan Claims Resubmitted - Out of Network", "PLAN", "RESUBMISSIONS", "OUT_OF_NETWORK", False, True, "Plan_Number_Claims_Resubmitted_Out_of_Network", "COUNT"),
    MetricDef(18, "PLAN_CLAIMS_RESUBMITTED_IN", "Plan Claims Resubmitted - In Network", "PLAN", "RESUBMISSIONS", "IN_NETWORK", False, True, "Plan_Number_Claims_Resubmitted_In_Network", "COUNT"),
    MetricDef(19, "PLAN_DENIED_REFERRAL_REQUIRED", "Plan Claims Denied - Referral/Prior Authorization Required", "PLAN", "DENIAL_REASONS", "ALL", False, True, "Plan_Number_Claims_Denied_Referral_Required", "COUNT"),
    MetricDef(20, "PLAN_DENIED_DUE_TO_OON", "Plan Claims Denied - Out of Network Provider", "PLAN", "DENIAL_REASONS", "ALL", False, True, "Plan_Number_Claims_Denied_Due_To_Out_Of_Network", "COUNT"),
    MetricDef(21, "PLAN_DENIED_SERVICES_EXCLUDED", "Plan Claims Denied - Services Excluded", "PLAN", "DENIAL_REASONS", "ALL", False, True, "Plan_Number_Claims_Denied_Services_Excluded", "COUNT"),
    MetricDef(22, "PLAN_DENIED_NOT_MED_NEC_EXCL_BH", "Plan Claims Denied - Not Medically Necessary Excluding Behavioral Health", "PLAN", "DENIAL_REASONS", "ALL", False, True, "Plan_Number_Claims_Denied_Not_Medically_Necessary_Excluding_Behavioral_Health", "COUNT"),
    MetricDef(23, "PLAN_DENIED_NOT_MED_NEC_BH_ONLY", "Plan Claims Denied - Not Medically Necessary Behavioral Health Only", "PLAN", "DENIAL_REASONS", "ALL", False, True, "Plan_Number_Claims_Denied_Not_Medically_Necessary_Behavioral_Health_Only", "COUNT"),
    MetricDef(24, "PLAN_DENIED_BENEFIT_LIMIT", "Plan Claims Denied - Benefit Limit Reached", "PLAN", "DENIAL_REASONS", "ALL", False, True, "Plan_Number_Claims_Denied_Due_To_Enrolle_Benefit_Limit_Reached", "COUNT"),
    MetricDef(25, "PLAN_DENIED_MEMBER_NOT_COVERED", "Plan Claims Denied - Member Not Covered", "PLAN", "DENIAL_REASONS", "ALL", False, True, "Plan_Number_Claims_Denied_Due_To_Member_Not_Covered", "COUNT"),
    MetricDef(26, "PLAN_DENIED_INVESTIGATIONAL", "Plan Claims Denied - Investigational/Experimental/Cosmetic", "PLAN", "DENIAL_REASONS", "ALL", False, True, "Plan_Number_Claims_Denied_Due_To_Investigational_Experimental_Cosmetic_Proceduce", "COUNT"),
    MetricDef(27, "PLAN_DENIED_ADMINISTRATIVE", "Plan Claims Denied - Administrative Reason", "PLAN", "DENIAL_REASONS", "ALL", False, True, "Plan_Number_Claims_Denied_Due_To_Administrative_Reason", "COUNT"),
    MetricDef(28, "PLAN_DENIED_OTHER", "Plan Claims Denied - Other", "PLAN", "DENIAL_REASONS", "ALL", False, True, "Plan_Number_Claims_Denied_Other", "COUNT"),
    MetricDef(29, "PLAN_AVG_MONTHLY_ENROLLMENT", "Average Monthly Enrollment", "PLAN", "ENROLLMENT", "NOT_APPLICABLE", False, False, "Average Monthly Enrollment", "DECIMAL"),
    MetricDef(30, "PLAN_AVG_MONTHLY_DISENROLLMENT", "Average Monthly Disenrollment", "PLAN", "ENROLLMENT", "NOT_APPLICABLE", False, False, "Average Monthly Disenrollment", "DECIMAL"),
]

ISSUER_METRICS = [m for m in METRICS if m.entity_level == "ISSUER"]
PLAN_METRICS = [m for m in METRICS if m.entity_level == "PLAN"]

ISSUER_FACT_NAME_BY_SOURCE = {
    "Issuer_Claims_Received_Out_of_Network": "ClaimsReceivedOutOfNetwork",
    "Issuer_Claims_Received_In_Network": "ClaimsReceivedInNetwork",
    "Issuer_Claims_Denied_Out_of_Network": "ClaimsDeniedOutOfNetwork",
    "Issuer_Claims_Denied_In_Network": "ClaimsDeniedInNetwork",
    "Issuer_Claims_Resubmitted_Out_of_Network": "ClaimsResubmittedOutOfNetwork",
    "Issuer_Claims_Resubmitted_In_Network": "ClaimsResubmittedInNetwork",
    "Issuer_Internal_Appeals_Filed": "InternalAppealsFiled",
    "Issuer_Number_Internal_Appeals_Overturned": "InternalAppealsOverturned",
    "Issuer_Percent_Internal_Appeals_Overturned": "InternalAppealsOverturnPctPublished",
    "Issuer_External_Appeals_Filed": "ExternalAppealsFiled",
    "Issuer_Number_External_Appeals_Overturned": "ExternalAppealsOverturned",
    "Issuer_Percent_External_Appeals_Overturned": "ExternalAppealsOverturnPctPublished",
}

PLAN_FACT_NAME_BY_SOURCE = {
    "Plan_Number_Claims_Received_Out_of_Network": "ClaimsReceivedOutOfNetwork",
    "Plan_Number_Claims_Received_In_Network": "ClaimsReceivedInNetwork",
    "Plan_Number_Claims_Denied_Out_of_Network": "ClaimsDeniedOutOfNetwork",
    "Plan_Number_Claims_Denied_In_Network": "ClaimsDeniedInNetwork",
    "Plan_Number_Claims_Resubmitted_Out_of_Network": "ClaimsResubmittedOutOfNetwork",
    "Plan_Number_Claims_Resubmitted_In_Network": "ClaimsResubmittedInNetwork",
    "Plan_Number_Claims_Denied_Referral_Required": "DeniedReferralOrPriorAuthorization",
    "Plan_Number_Claims_Denied_Due_To_Out_Of_Network": "DeniedDueToOutOfNetworkProvider",
    "Plan_Number_Claims_Denied_Services_Excluded": "DeniedServicesExcluded",
    "Plan_Number_Claims_Denied_Not_Medically_Necessary_Excluding_Behavioral_Health": "DeniedNotMedicallyNecessaryExclBH",
    "Plan_Number_Claims_Denied_Not_Medically_Necessary_Behavioral_Health_Only": "DeniedNotMedicallyNecessaryBHOnly",
    "Plan_Number_Claims_Denied_Due_To_Enrolle_Benefit_Limit_Reached": "DeniedBenefitLimitReached",
    "Plan_Number_Claims_Denied_Due_To_Member_Not_Covered": "DeniedMemberNotCovered",
    "Plan_Number_Claims_Denied_Due_To_Investigational_Experimental_Cosmetic_Proceduce": "DeniedInvestigationalExperimentalCosmetic",
    "Plan_Number_Claims_Denied_Due_To_Administrative_Reason": "DeniedAdministrativeReason",
    "Plan_Number_Claims_Denied_Other": "DeniedOther",
    "Average Monthly Enrollment": "AverageMonthlyEnrollment",
    "Average Monthly Disenrollment": "AverageMonthlyDisenrollment",
}

COUNT_CANONICAL_COLUMNS = {
    ISSUER_FACT_NAME_BY_SOURCE[m.source_column]
    for m in ISSUER_METRICS
    if m.value_kind == "COUNT"
} | {
    PLAN_FACT_NAME_BY_SOURCE[m.source_column]
    for m in PLAN_METRICS
    if m.value_kind == "COUNT"
}


def project_root() -> Path:
    return Path(__file__).resolve().parents[2]


PROJECT_ROOT = project_root()
SOURCE_FILE = PROJECT_ROOT / "data" / "raw" / "Transparency_in_Coverage_PUF.xlsx"
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
REVIEW_FILE = REVIEW_DIR / "02_CANONICAL_BUILD_REPORT.xlsx"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def text_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and np.isnan(value):
        return ""
    return str(value).strip()


def id_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, float):
        if np.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
    value = str(value).strip()
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]
    return value


def yes_no_to_bool(value: Any) -> bool | None:
    v = text_value(value)
    if v == "Yes":
        return True
    if v == "No":
        return False
    return None


def decimal_from_raw(value: Any) -> Decimal | None:
    raw = text_value(value)
    if raw == "":
        return None
    try:
        return Decimal(raw.replace(",", ""))
    except InvalidOperation:
        return None


def canonical_compare_value(value: Any) -> str:
    raw = text_value(value)
    if raw in SPECIAL_STATUS_MAP:
        return raw
    if raw == "":
        return "<BLANK>"
    dec = decimal_from_raw(raw)
    if dec is not None:
        return format(dec.normalize(), "f")
    return raw


def detect_header_row(sheet_name: str) -> int:
    preview = pd.read_excel(
        SOURCE_FILE,
        sheet_name=sheet_name,
        header=None,
        nrows=10,
        dtype=object,
        keep_default_na=False,
        na_filter=False,
    )
    for idx in preview.index:
        values = {text_value(v) for v in preview.loc[idx].tolist()}
        if "Plan_ID" in values and "Issuer_ID" in values and "State" in values:
            return int(idx)
    raise RuntimeError(f"Could not detect analytical header row in sheet: {sheet_name}")


def load_source() -> pd.DataFrame:
    frames: list[pd.DataFrame] = []

    workbook = pd.ExcelFile(SOURCE_FILE)
    actual_sheets = workbook.sheet_names

    missing_sheets = [
        s for s in EXPECTED_ANALYTICAL_SHEETS if s not in actual_sheets
    ]
    if missing_sheets:
        raise RuntimeError(f"Missing required source sheets: {missing_sheets}")

    for sheet_name in EXPECTED_ANALYTICAL_SHEETS:
        header_idx = detect_header_row(sheet_name)

        df = pd.read_excel(
            SOURCE_FILE,
            sheet_name=sheet_name,
            header=header_idx,
            dtype=object,
            keep_default_na=False,
            na_filter=False,
        )

        # Keep only rows that contain at least one nonblank analytical value.
        nonblank_mask = df.apply(
            lambda row: any(text_value(v) != "" for v in row),
            axis=1,
        )
        df = df.loc[nonblank_mask].copy()

        actual_columns = list(df.columns)
        if actual_columns != EXPECTED_COLUMNS:
            missing = [c for c in EXPECTED_COLUMNS if c not in actual_columns]
            extra = [c for c in actual_columns if c not in EXPECTED_COLUMNS]
            raise RuntimeError(
                "Source schema drift detected.\n"
                f"Sheet: {sheet_name}\n"
                f"Missing columns: {missing}\n"
                f"Unexpected columns: {extra}\n"
                "Column order or names no longer match the approved data contract."
            )

        df["_SourceSheet"] = sheet_name
        # Excel rows are 1-based; row after zero-based header index.
        df["_SourceExcelRow"] = np.arange(
            header_idx + 2,
            header_idx + 2 + len(df),
        )

        frames.append(df)

    combined = pd.concat(frames, ignore_index=True, sort=False)

    for column in EXPECTED_COLUMNS:
        combined[column] = combined[column].map(
            lambda v: text_value(v) if isinstance(v, str) else v
        )

    combined["Issuer_ID"] = combined["Issuer_ID"].map(id_value)
    combined["Plan_ID"] = combined["Plan_ID"].map(id_value)
    combined["State"] = combined["State"].map(text_value)

    return combined


def parse_metric_value(
    raw_value: Any,
    metric: MetricDef,
) -> tuple[int | float | None, str, str]:
    raw = text_value(raw_value)

    if raw in SPECIAL_STATUS_MAP:
        return None, SPECIAL_STATUS_MAP[raw], raw

    if raw == "":
        return None, "SOURCE_MISSING_UNEXPECTED", ""

    dec = decimal_from_raw(raw)
    if dec is None:
        return None, "INVALID_NUMERIC_SOURCE", raw

    if metric.value_kind == "COUNT":
        if dec != dec.to_integral_value():
            return None, "INVALID_NUMERIC_SOURCE", raw
        return int(dec), "AVAILABLE", ""

    return float(dec), "AVAILABLE", ""


def create_dq_row(
    dq_rows: list[dict[str, Any]],
    *,
    severity: str,
    scope: str,
    entity_key: str,
    source_field: str,
    rule_code: str,
    observed_value: str,
    expected_condition: str,
    source_sheet: str = "",
    source_excel_row: int | str = "",
    known_source_exception: bool = False,
    resolution_state: str = "OPEN_REVIEW",
) -> None:
    dq_rows.append(
        {
            "Severity": severity,
            "Scope": scope,
            "EntityKey": entity_key,
            "SourceField": source_field,
            "RuleCode": rule_code,
            "ObservedValue": observed_value,
            "ExpectedCondition": expected_condition,
            "SourceSheet": source_sheet,
            "SourceExcelRow": source_excel_row,
            "IsKnownSourceException": known_source_exception,
            "ResolutionState": resolution_state,
        }
    )


def write_review_workbook(
    summary_df: pd.DataFrame,
    acceptance_df: pd.DataFrame,
    inventory_df: pd.DataFrame,
    dq_summary_df: pd.DataFrame,
    dq_df: pd.DataFrame,
    availability_summary_df: pd.DataFrame,
    appeal_recon_df: pd.DataFrame,
    grain_recon_df: pd.DataFrame,
    source_mapping_df: pd.DataFrame,
    dim_metric: pd.DataFrame,
    special_status_df: pd.DataFrame,
    diagnostic_df: pd.DataFrame,
) -> None:
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(REVIEW_FILE, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="00_Summary", index=False)
        acceptance_df.to_excel(writer, sheet_name="01_Acceptance", index=False)
        inventory_df.to_excel(writer, sheet_name="02_Inventory", index=False)
        dq_summary_df.to_excel(writer, sheet_name="03_DQ_Summary", index=False)
        dq_df.to_excel(writer, sheet_name="04_DQ_Exceptions", index=False)
        availability_summary_df.to_excel(
            writer, sheet_name="05_Availability", index=False
        )
        appeal_recon_df.to_excel(
            writer, sheet_name="06_Appeal_Recon", index=False
        )
        grain_recon_df.to_excel(
            writer, sheet_name="07_Grain_Recon", index=False
        )
        source_mapping_df.to_excel(
            writer, sheet_name="08_Source_Mapping", index=False
        )
        dim_metric.to_excel(writer, sheet_name="09_Metric_Dict", index=False)
        special_status_df.to_excel(
            writer, sheet_name="10_Status_Dict", index=False
        )
        diagnostic_df.to_excel(
            writer, sheet_name="11_Diagnostic_Observations", index=False
        )

        workbook = writer.book
        for worksheet in workbook.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for column_cells in worksheet.columns:
                width = min(
                    max(
                        10,
                        max(
                            len(str(cell.value)) if cell.value is not None else 0
                            for cell in column_cells[:200]
                        )
                        + 2,
                    ),
                    60,
                )
                worksheet.column_dimensions[
                    column_cells[0].column_letter
                ].width = width


def main() -> int:
    print("=" * 78)
    print("Transparency in Coverage PUF — Canonical Dataset Build")
    print("=" * 78)
    print(f"Script version : {SCRIPT_VERSION}")
    print(f"Source         : {SOURCE_FILE}")
    print(f"Canonical dir  : {CANONICAL_DIR}")
    print(f"Review file    : {REVIEW_FILE}")

    if not SOURCE_FILE.exists():
        raise FileNotFoundError(f"Source file not found: {SOURCE_FILE}")

    source_hash = sha256(SOURCE_FILE)
    print(f"SHA256         : {source_hash}")

    if source_hash.lower() != EXPECTED_SOURCE_SHA256.lower():
        raise RuntimeError(
            "RAW source fingerprint does not match the approved project baseline.\n"
            f"Expected: {EXPECTED_SOURCE_SHA256}\n"
            f"Actual:   {source_hash}"
        )

    print("Source integrity: PASS")

    raw = load_source()
    source_row_count = len(raw)

    print(f"Analytical rows: {source_row_count:,}")

    acceptance: list[dict[str, Any]] = []
    dq_rows: list[dict[str, Any]] = []
    diagnostic_rows: list[dict[str, Any]] = []

    def add_acceptance(
        test_id: str,
        test_name: str,
        expected: Any,
        actual: Any,
        status: str,
        gate_type: str = "STRUCTURAL",
    ) -> None:
        acceptance.append(
            {
                "TestID": test_id,
                "TestName": test_name,
                "Expected": expected,
                "Actual": actual,
                "Status": status,
                "GateType": gate_type,
            }
        )

    # ------------------------------------------------------------------
    # Structural key / schema validation before canonical construction.
    # ------------------------------------------------------------------

    plan_id = raw["Plan_ID"].map(id_value)
    issuer_id = raw["Issuer_ID"].map(id_value)
    state = raw["State"].map(text_value)

    missing_plan = int((plan_id == "").sum())
    missing_issuer = int((issuer_id == "").sum())
    missing_state = int((state == "").sum())

    duplicate_plan_rows = int(plan_id.duplicated(keep=False).sum())

    invalid_issuer_pattern = int(
        (~issuer_id.str.fullmatch(r"\d{5}", na=False)).sum()
    )

    plan_match = plan_id.str.extract(r"^(\d{5})([A-Z]{2})(\d{7})$")
    invalid_plan_pattern = int(plan_match.isna().any(axis=1).sum())

    issuer_embedded_mismatch = int(
        (
            plan_match[0].notna()
            & (plan_match[0].astype(str) != issuer_id.astype(str))
        ).sum()
    )
    state_embedded_mismatch = int(
        (
            plan_match[1].notna()
            & (plan_match[1].astype(str) != state.astype(str))
        ).sum()
    )

    state_exchange_max = int(
        raw.groupby("State")["Exchange_Type"].nunique(dropna=False).max()
    )

    issuer_attr_columns = [
        "Issuer_Name",
        "State",
        "Exchange_Type",
        "Is_Issuer_New_to_Exchange?(Yes_or_No)",
        "SADP_Only",
        "URL_Claims_Payment_Policies",
        "Rate_Review",
        "Financial_Information",
    ]

    issuer_attr_inconsistency = 0
    for column in issuer_attr_columns:
        n_unique = raw.groupby("Issuer_ID")[column].apply(
            lambda s: len({text_value(v) for v in s})
        )
        issuer_attr_inconsistency += int((n_unique > 1).sum())

    issuer_metric_inconsistency = 0
    for metric in ISSUER_METRICS:
        n_unique = raw.groupby(["State", "Issuer_ID"])[metric.source_column].apply(
            lambda s: len({canonical_compare_value(v) for v in s})
        )
        issuer_metric_inconsistency += int((n_unique > 1).sum())

    add_acceptance("STR-001", "Source analytical rows", 4956, source_row_count, "PASS" if source_row_count == 4956 else "FAIL")
    add_acceptance("STR-002", "Missing Plan_ID", 0, missing_plan, "PASS" if missing_plan == 0 else "FAIL")
    add_acceptance("STR-003", "Missing Issuer_ID", 0, missing_issuer, "PASS" if missing_issuer == 0 else "FAIL")
    add_acceptance("STR-004", "Missing State", 0, missing_state, "PASS" if missing_state == 0 else "FAIL")
    add_acceptance("STR-005", "Duplicate Plan_ID rows", 0, duplicate_plan_rows, "PASS" if duplicate_plan_rows == 0 else "FAIL")
    add_acceptance("STR-006", "Invalid Issuer_ID pattern", 0, invalid_issuer_pattern, "PASS" if invalid_issuer_pattern == 0 else "FAIL")
    add_acceptance("STR-007", "Invalid Plan_ID pattern", 0, invalid_plan_pattern, "PASS" if invalid_plan_pattern == 0 else "FAIL")
    add_acceptance("STR-008", "Plan_ID embedded issuer mismatch", 0, issuer_embedded_mismatch, "PASS" if issuer_embedded_mismatch == 0 else "FAIL")
    add_acceptance("STR-009", "Plan_ID embedded state mismatch", 0, state_embedded_mismatch, "PASS" if state_embedded_mismatch == 0 else "FAIL")
    add_acceptance("STR-010", "Maximum Exchange_Type values per State", 1, state_exchange_max, "PASS" if state_exchange_max == 1 else "FAIL")
    add_acceptance("STR-011", "Issuer attribute inconsistencies", 0, issuer_attr_inconsistency, "PASS" if issuer_attr_inconsistency == 0 else "FAIL")
    add_acceptance("STR-012", "Issuer metric inconsistencies", 0, issuer_metric_inconsistency, "PASS" if issuer_metric_inconsistency == 0 else "FAIL")

    structural_failed = any(
        row["GateType"] == "STRUCTURAL" and row["Status"] == "FAIL"
        for row in acceptance
    )

    if structural_failed:
        acceptance_df = pd.DataFrame(acceptance)
        summary_df = pd.DataFrame(
            [
                ("BuildStatus", "HOLD_STRUCTURAL_FAILURE"),
                ("ScriptVersion", SCRIPT_VERSION),
                ("SourceFile", SOURCE_FILE.name),
                ("SourceSHA256", source_hash),
                ("SourceRows", source_row_count),
                ("CanonicalWritten", "NO"),
            ],
            columns=["Item", "Value"],
        )
        empty = pd.DataFrame()
        write_review_workbook(
            summary_df,
            acceptance_df,
            empty,
            empty,
            empty,
            empty,
            empty,
            empty,
            empty,
            empty,
            pd.DataFrame(
                AVAILABILITY_STATUS_ROWS,
                columns=[
                    "StatusKey",
                    "StatusCode",
                    "IsNumericAvailable",
                    "IsDisclosureLimited",
                    "IsDQIssue",
                ],
            ),
            pd.DataFrame(),
        )
        print("Build status: HOLD — structural validation failed.")
        print(f"Review evidence: {REVIEW_FILE}")
        return 2

    # ------------------------------------------------------------------
    # Conformed dimensions.
    # ------------------------------------------------------------------

    dim_reporting_period = pd.DataFrame(
        [
            {
                "ReportingPeriodKey": EXPERIENCE_YEAR,
                "PUFYear": PUF_YEAR,
                "ExperienceYear": EXPERIENCE_YEAR,
                "PeriodStartDate": f"{EXPERIENCE_YEAR}-01-01",
                "PeriodEndDate": f"{EXPERIENCE_YEAR}-12-31",
            }
        ]
    )

    state_rows = (
        raw[["State", "Exchange_Type"]]
        .drop_duplicates()
        .sort_values("State")
        .reset_index(drop=True)
    )
    state_rows.insert(0, "StateKey", np.arange(1, len(state_rows) + 1))
    dim_state = state_rows.rename(
        columns={"State": "StateCode", "Exchange_Type": "ExchangeType"}
    )

    state_key_map = dict(
        zip(dim_state["StateCode"], dim_state["StateKey"])
    )

    issuer_source = (
        raw.sort_values(["Issuer_ID", "State", "_SourceSheet", "_SourceExcelRow"])
        .groupby(["Issuer_ID", "State"], as_index=False, sort=True)
        .first()
    )

    issuer_source["IssuerKey"] = np.arange(1, len(issuer_source) + 1)
    issuer_source["StateKey"] = issuer_source["State"].map(state_key_map)

    financial_status = issuer_source["Financial_Information"].map(
        lambda v: "MISSING_URL"
        if text_value(v) == "Missing URL"
        else ("AVAILABLE" if text_value(v) != "" else "SOURCE_MISSING_UNEXPECTED")
    )
    financial_url = issuer_source["Financial_Information"].map(
        lambda v: ""
        if text_value(v) == "Missing URL"
        else text_value(v)
    )

    dim_issuer = pd.DataFrame(
        {
            "IssuerKey": issuer_source["IssuerKey"],
            "IssuerID": issuer_source["Issuer_ID"].map(id_value),
            "IssuerName": issuer_source["Issuer_Name"].map(text_value),
            "StateKey": issuer_source["StateKey"],
            "IsIssuerNewToExchange": issuer_source[
                "Is_Issuer_New_to_Exchange?(Yes_or_No)"
            ].map(yes_no_to_bool),
            "IsSADPOnlyIssuer": issuer_source["SADP_Only"].map(yes_no_to_bool),
            "ClaimsPaymentPoliciesURL": issuer_source[
                "URL_Claims_Payment_Policies"
            ].map(text_value),
            "RateReviewURL": issuer_source["Rate_Review"].map(text_value),
            "FinancialInformationURL": financial_url,
            "FinancialInformationStatusCode": financial_status,
        }
    )

    issuer_key_map = dict(
        zip(dim_issuer["IssuerID"], dim_issuer["IssuerKey"])
    )

    plan_source = raw.sort_values("Plan_ID").copy()
    plan_source["PlanKey"] = np.arange(1, len(plan_source) + 1)
    plan_source["StateKey"] = plan_source["State"].map(state_key_map)
    plan_source["IssuerKey"] = plan_source["Issuer_ID"].map(issuer_key_map)

    dim_plan = pd.DataFrame(
        {
            "PlanKey": plan_source["PlanKey"],
            "PlanID": plan_source["Plan_ID"].map(id_value),
            "IssuerKey": plan_source["IssuerKey"],
            "StateKey": plan_source["StateKey"],
            "MarketSegment": plan_source["Individual/SHOP"].map(text_value),
            "PlanType": plan_source["Plan_Type"].map(text_value),
            "PlanOfferingType": plan_source["QHP or SADP?"].map(text_value),
            "MetalLevel": plan_source["Metal_Level"].map(text_value),
            "SourceSheet": plan_source["_SourceSheet"],
            "SourceExcelRow": plan_source["_SourceExcelRow"],
            "SourceFileName": SOURCE_FILE.name,
            "SourceFileSHA256": source_hash,
        }
    )

    dim_metric = pd.DataFrame(
        [
            {
                "MetricKey": m.metric_key,
                "MetricCode": m.metric_code,
                "MetricDisplayName": m.display_name,
                "EntityLevel": m.entity_level,
                "MetricFamily": m.metric_family,
                "NetworkScope": m.network_scope,
                "IsPercentage": m.is_percentage,
                "IsAdditive": m.is_additive,
                "SourceColumnName": m.source_column,
                "ExperienceYear": EXPERIENCE_YEAR,
            }
            for m in METRICS
        ]
    )

    dim_availability_status = pd.DataFrame(
        AVAILABILITY_STATUS_ROWS,
        columns=[
            "StatusKey",
            "StatusCode",
            "IsNumericAvailable",
            "IsDisclosureLimited",
            "IsDQIssue",
        ],
    )

    # ------------------------------------------------------------------
    # Plan fact + plan availability.
    # ------------------------------------------------------------------

    fact_plan = pd.DataFrame(
        {
            "ReportingPeriodKey": EXPERIENCE_YEAR,
            "StateKey": plan_source["StateKey"],
            "IssuerKey": plan_source["IssuerKey"],
            "PlanKey": plan_source["PlanKey"],
        }
    )

    plan_availability_rows: list[dict[str, Any]] = []

    for metric in PLAN_METRICS:
        canonical_name = PLAN_FACT_NAME_BY_SOURCE[metric.source_column]
        values: list[int | float | None] = []

        for _, row in plan_source.iterrows():
            value, status_code, source_token = parse_metric_value(
                row[metric.source_column], metric
            )
            values.append(value)

            plan_availability_rows.append(
                {
                    "ReportingPeriodKey": EXPERIENCE_YEAR,
                    "PlanKey": int(row["PlanKey"]),
                    "MetricKey": metric.metric_key,
                    "AvailabilityStatusKey": STATUS_KEY_BY_CODE[status_code],
                    "SourceStatusToken": source_token,
                }
            )

            if status_code in {
                "SOURCE_MISSING_UNEXPECTED",
                "INVALID_NUMERIC_SOURCE",
            }:
                create_dq_row(
                    dq_rows,
                    severity="ERROR",
                    scope="PLAN",
                    entity_key=id_value(row["Plan_ID"]),
                    source_field=metric.source_column,
                    rule_code=status_code,
                    observed_value=text_value(row[metric.source_column]),
                    expected_condition="Numeric value or governed CMS special-status token",
                    source_sheet=row["_SourceSheet"],
                    source_excel_row=int(row["_SourceExcelRow"]),
                )

        fact_plan[canonical_name] = values

    fact_plan_metric_availability = pd.DataFrame(plan_availability_rows)

    # ------------------------------------------------------------------
    # Issuer fact + issuer availability.
    # ------------------------------------------------------------------

    issuer_groups = raw.groupby(["State", "Issuer_ID"], sort=True)
    issuer_fact_rows: list[dict[str, Any]] = []
    issuer_availability_rows: list[dict[str, Any]] = []

    for (state_code, issuer_id_value), group in issuer_groups:
        issuer_id_text = id_value(issuer_id_value)
        issuer_key = issuer_key_map[issuer_id_text]
        state_key = state_key_map[text_value(state_code)]

        fact_row: dict[str, Any] = {
            "ReportingPeriodKey": EXPERIENCE_YEAR,
            "StateKey": state_key,
            "IssuerKey": issuer_key,
        }

        for metric in ISSUER_METRICS:
            canonical_name = ISSUER_FACT_NAME_BY_SOURCE[metric.source_column]

            normalized_values = {
                canonical_compare_value(v)
                for v in group[metric.source_column]
            }
            if len(normalized_values) != 1:
                # Already protected by structural gate, but keep defensive guard.
                raise RuntimeError(
                    f"Issuer metric inconsistency after gate: "
                    f"{issuer_id_text} / {metric.source_column}"
                )

            raw_value = group.iloc[0][metric.source_column]
            value, status_code, source_token = parse_metric_value(
                raw_value, metric
            )
            fact_row[canonical_name] = value

            issuer_availability_rows.append(
                {
                    "ReportingPeriodKey": EXPERIENCE_YEAR,
                    "StateKey": state_key,
                    "IssuerKey": issuer_key,
                    "MetricKey": metric.metric_key,
                    "AvailabilityStatusKey": STATUS_KEY_BY_CODE[status_code],
                    "SourceStatusToken": source_token,
                }
            )

            if status_code in {
                "SOURCE_MISSING_UNEXPECTED",
                "INVALID_NUMERIC_SOURCE",
            }:
                first_row = group.iloc[0]
                create_dq_row(
                    dq_rows,
                    severity="ERROR",
                    scope="ISSUER",
                    entity_key=issuer_id_text,
                    source_field=metric.source_column,
                    rule_code=status_code,
                    observed_value=text_value(raw_value),
                    expected_condition="Numeric value or governed CMS special-status token",
                    source_sheet=first_row["_SourceSheet"],
                    source_excel_row=int(first_row["_SourceExcelRow"]),
                )

        fact_row["SourcePlanRowCount"] = len(group)
        fact_row["SourceSheetCount"] = group["_SourceSheet"].nunique()
        fact_row["SourceSheets"] = " | ".join(
            sorted(group["_SourceSheet"].unique())
        )
        fact_row["SourceFileName"] = SOURCE_FILE.name
        fact_row["SourceFileSHA256"] = source_hash

        issuer_fact_rows.append(fact_row)

    fact_issuer = pd.DataFrame(issuer_fact_rows)
    fact_issuer_metric_availability = pd.DataFrame(
        issuer_availability_rows
    )

    # Apply nullable integer types to count metrics.
    for column in fact_plan.columns:
        if column in COUNT_CANONICAL_COLUMNS:
            fact_plan[column] = pd.array(
                fact_plan[column], dtype="Int64"
            )

    for column in fact_issuer.columns:
        if column in COUNT_CANONICAL_COLUMNS:
            fact_issuer[column] = pd.array(
                fact_issuer[column], dtype="Int64"
            )

    # ------------------------------------------------------------------
    # Logical DQ tests. These preserve data and create evidence only.
    # ------------------------------------------------------------------

    issuer_id_by_key = dict(
        zip(dim_issuer["IssuerKey"], dim_issuer["IssuerID"])
    )
    plan_id_by_key = dict(zip(dim_plan["PlanKey"], dim_plan["PlanID"]))

    def compare_fact_columns(
        df: pd.DataFrame,
        *,
        scope: str,
        left: str,
        right: str,
        rule_code: str,
        expected: str,
        severity: str = "WARNING",
    ) -> None:
        mask = (
            df[left].notna()
            & df[right].notna()
            & (df[left] > df[right])
        )
        for _, row in df.loc[mask].iterrows():
            if scope == "ISSUER":
                entity = issuer_id_by_key[int(row["IssuerKey"])]
            else:
                entity = plan_id_by_key[int(row["PlanKey"])]

            known = (
                scope == "ISSUER"
                and entity == "97725"
                and rule_code == "APPEALS_OVERTURNED_GT_FILED"
                and left == "InternalAppealsOverturned"
            )
            create_dq_row(
                dq_rows,
                severity=severity,
                scope=scope,
                entity_key=entity,
                source_field=f"{left} vs {right}",
                rule_code=rule_code,
                observed_value=f"{left}={row[left]}; {right}={row[right]}",
                expected_condition=expected,
                known_source_exception=known,
                resolution_state=(
                    "KNOWN_SOURCE_EXCEPTION_PRESERVED"
                    if known
                    else "OPEN_REVIEW"
                ),
            )

    compare_fact_columns(
        fact_issuer,
        scope="ISSUER",
        left="ClaimsDeniedInNetwork",
        right="ClaimsReceivedInNetwork",
        rule_code="DENIED_GT_RECEIVED",
        expected="Denied claims <= received claims for comparable available values",
    )
    compare_fact_columns(
        fact_issuer,
        scope="ISSUER",
        left="ClaimsDeniedOutOfNetwork",
        right="ClaimsReceivedOutOfNetwork",
        rule_code="DENIED_GT_RECEIVED",
        expected="Denied claims <= received claims for comparable available values",
    )
    compare_fact_columns(
        fact_plan,
        scope="PLAN",
        left="ClaimsDeniedInNetwork",
        right="ClaimsReceivedInNetwork",
        rule_code="DENIED_GT_RECEIVED",
        expected="Denied claims <= received claims for comparable available values",
    )
    compare_fact_columns(
        fact_plan,
        scope="PLAN",
        left="ClaimsDeniedOutOfNetwork",
        right="ClaimsReceivedOutOfNetwork",
        rule_code="DENIED_GT_RECEIVED",
        expected="Denied claims <= received claims for comparable available values",
    )

    def observe_resubmissions_gt_denied(
        df: pd.DataFrame,
        *,
        scope: str,
        resubmitted_col: str,
        denied_col: str,
        network_scope: str,
    ) -> None:
        mask = (
            df[resubmitted_col].notna()
            & df[denied_col].notna()
            & (df[resubmitted_col] > df[denied_col])
        )
        for _, row in df.loc[mask].iterrows():
            if scope == "ISSUER":
                entity = issuer_id_by_key[int(row["IssuerKey"])]
            else:
                entity = plan_id_by_key[int(row["PlanKey"])]

            diagnostic_rows.append(
                {
                    "ObservationType": "RESUBMISSION_EVENTS_GT_DENIED_CLAIMS",
                    "Scope": scope,
                    "EntityKey": entity,
                    "NetworkScope": network_scope,
                    "ResubmittedValue": row[resubmitted_col],
                    "DeniedValue": row[denied_col],
                    "Interpretation": (
                        "Informational only. CMS defines this field as the "
                        "number of claim resubmissions for claims previously "
                        "denied; the published definition does not establish "
                        "a one-resubmission-per-denied-claim constraint."
                    ),
                    "DQStatus": "NOT_A_DQ_FAILURE",
                }
            )

    observe_resubmissions_gt_denied(
        fact_issuer,
        scope="ISSUER",
        resubmitted_col="ClaimsResubmittedInNetwork",
        denied_col="ClaimsDeniedInNetwork",
        network_scope="IN_NETWORK",
    )
    observe_resubmissions_gt_denied(
        fact_issuer,
        scope="ISSUER",
        resubmitted_col="ClaimsResubmittedOutOfNetwork",
        denied_col="ClaimsDeniedOutOfNetwork",
        network_scope="OUT_OF_NETWORK",
    )
    observe_resubmissions_gt_denied(
        fact_plan,
        scope="PLAN",
        resubmitted_col="ClaimsResubmittedInNetwork",
        denied_col="ClaimsDeniedInNetwork",
        network_scope="IN_NETWORK",
    )
    observe_resubmissions_gt_denied(
        fact_plan,
        scope="PLAN",
        resubmitted_col="ClaimsResubmittedOutOfNetwork",
        denied_col="ClaimsDeniedOutOfNetwork",
        network_scope="OUT_OF_NETWORK",
    )

    compare_fact_columns(
        fact_issuer,
        scope="ISSUER",
        left="InternalAppealsOverturned",
        right="InternalAppealsFiled",
        rule_code="APPEALS_OVERTURNED_GT_FILED",
        expected="Appeals overturned <= appeals filed",
    )
    compare_fact_columns(
        fact_issuer,
        scope="ISSUER",
        left="ExternalAppealsOverturned",
        right="ExternalAppealsFiled",
        rule_code="APPEALS_OVERTURNED_GT_FILED",
        expected="Appeals overturned <= appeals filed",
    )

    compare_fact_columns(
        fact_plan,
        scope="PLAN",
        left="AverageMonthlyDisenrollment",
        right="AverageMonthlyEnrollment",
        rule_code="DISENROLLMENT_GT_ENROLLMENT",
        expected="Average monthly disenrollment <= average monthly enrollment",
    )

    # Negative count tests.
    for scope, df, key_column, key_lookup in [
        ("ISSUER", fact_issuer, "IssuerKey", issuer_id_by_key),
        ("PLAN", fact_plan, "PlanKey", plan_id_by_key),
    ]:
        for column in [
            c for c in df.columns if c in COUNT_CANONICAL_COLUMNS
        ]:
            mask = df[column].notna() & (df[column] < 0)
            for _, row in df.loc[mask].iterrows():
                create_dq_row(
                    dq_rows,
                    severity="WARNING",
                    scope=scope,
                    entity_key=key_lookup[int(row[key_column])],
                    source_field=column,
                    rule_code="NEGATIVE_COUNT",
                    observed_value=str(row[column]),
                    expected_condition="Count >= 0",
                )

    # Published appeal percentage formula reconciliation.
    appeal_recon_rows: list[dict[str, Any]] = []

    for label, filed_col, overturned_col, pct_col in [
        (
            "Internal",
            "InternalAppealsFiled",
            "InternalAppealsOverturned",
            "InternalAppealsOverturnPctPublished",
        ),
        (
            "External",
            "ExternalAppealsFiled",
            "ExternalAppealsOverturned",
            "ExternalAppealsOverturnPctPublished",
        ),
    ]:
        comparable = (
            fact_issuer[filed_col].notna()
            & fact_issuer[overturned_col].notna()
            & fact_issuer[pct_col].notna()
            & (fact_issuer[filed_col] != 0)
        )

        calc = (
            fact_issuer.loc[comparable, overturned_col].astype(float)
            / fact_issuer.loc[comparable, filed_col].astype(float)
            * 100.0
        )
        published = fact_issuer.loc[comparable, pct_col].astype(float)
        diff = (calc - published).abs()
        mismatch = diff > 0.011

        appeal_recon_rows.append(
            {
                "AppealType": label,
                "ComparableIssuerRows": int(comparable.sum()),
                "FormulaMatchesWithinTolerance": int((~mismatch).sum()),
                "FormulaMismatches": int(mismatch.sum()),
                "TolerancePercentagePoints": 0.011,
                "MaxAbsoluteDifference": (
                    float(diff.max()) if len(diff) else np.nan
                ),
                "Status": "PASS" if int(mismatch.sum()) == 0 else "REVIEW",
            }
        )

        if int(mismatch.sum()) > 0:
            mismatch_indices = diff.index[mismatch]
            for idx in mismatch_indices:
                row = fact_issuer.loc[idx]
                entity = issuer_id_by_key[int(row["IssuerKey"])]
                create_dq_row(
                    dq_rows,
                    severity="WARNING",
                    scope="ISSUER",
                    entity_key=entity,
                    source_field=pct_col,
                    rule_code="PUBLISHED_APPEAL_PCT_MISMATCH",
                    observed_value=(
                        f"published={row[pct_col]}; "
                        f"calculated={calc.loc[idx]:.6f}"
                    ),
                    expected_condition=(
                        "Published percentage matches overturned/filed × 100 "
                        "within 0.011 percentage points"
                    ),
                )

    # Percentage range.
    for column in [
        "InternalAppealsOverturnPctPublished",
        "ExternalAppealsOverturnPctPublished",
    ]:
        mask = fact_issuer[column].notna() & (
            (fact_issuer[column] < 0) | (fact_issuer[column] > 100)
        )
        for _, row in fact_issuer.loc[mask].iterrows():
            entity = issuer_id_by_key[int(row["IssuerKey"])]
            known = (
                entity == "97725"
                and column == "InternalAppealsOverturnPctPublished"
            )
            create_dq_row(
                dq_rows,
                severity="WARNING",
                scope="ISSUER",
                entity_key=entity,
                source_field=column,
                rule_code="PERCENT_OUTSIDE_0_100",
                observed_value=str(row[column]),
                expected_condition="Expected percentage range 0–100 unless source anomaly",
                known_source_exception=known,
                resolution_state=(
                    "KNOWN_SOURCE_EXCEPTION_PRESERVED"
                    if known
                    else "OPEN_REVIEW"
                ),
            )

    # ------------------------------------------------------------------
    # Acceptance tests for canonical shape and relationships.
    # ------------------------------------------------------------------

    add_acceptance("CAN-001", "DimReportingPeriod rows", 1, len(dim_reporting_period), "PASS" if len(dim_reporting_period) == 1 else "FAIL")
    add_acceptance("CAN-002", "DimState rows", 30, len(dim_state), "PASS" if len(dim_state) == 30 else "FAIL")
    add_acceptance("CAN-003", "DimIssuer rows", 348, len(dim_issuer), "PASS" if len(dim_issuer) == 348 else "FAIL")
    add_acceptance("CAN-004", "DimPlan rows", 4956, len(dim_plan), "PASS" if len(dim_plan) == 4956 else "FAIL")
    add_acceptance("CAN-005", "FactIssuerTransparency rows", 348, len(fact_issuer), "PASS" if len(fact_issuer) == 348 else "FAIL")
    add_acceptance("CAN-006", "FactPlanTransparency rows", 4956, len(fact_plan), "PASS" if len(fact_plan) == 4956 else "FAIL")
    add_acceptance("CAN-007", "FactIssuerMetricAvailability rows", 4176, len(fact_issuer_metric_availability), "PASS" if len(fact_issuer_metric_availability) == 4176 else "FAIL")
    add_acceptance("CAN-008", "FactPlanMetricAvailability rows", 89208, len(fact_plan_metric_availability), "PASS" if len(fact_plan_metric_availability) == 89208 else "FAIL")
    add_acceptance("CAN-009", "DimMetric rows", 30, len(dim_metric), "PASS" if len(dim_metric) == 30 else "FAIL")
    add_acceptance("CAN-010", "Duplicate DimPlan PlanID", 0, int(dim_plan["PlanID"].duplicated().sum()), "PASS" if int(dim_plan["PlanID"].duplicated().sum()) == 0 else "FAIL")
    add_acceptance("CAN-011", "Duplicate DimIssuer IssuerID", 0, int(dim_issuer["IssuerID"].duplicated().sum()), "PASS" if int(dim_issuer["IssuerID"].duplicated().sum()) == 0 else "FAIL")

    broken_plan_issuer = int(
        (~fact_plan["IssuerKey"].isin(dim_issuer["IssuerKey"])).sum()
    )
    broken_plan_state = int(
        (~fact_plan["StateKey"].isin(dim_state["StateKey"])).sum()
    )
    broken_plan_plan = int(
        (~fact_plan["PlanKey"].isin(dim_plan["PlanKey"])).sum()
    )
    broken_issuer_state = int(
        (~fact_issuer["StateKey"].isin(dim_state["StateKey"])).sum()
    )
    broken_issuer_issuer = int(
        (~fact_issuer["IssuerKey"].isin(dim_issuer["IssuerKey"])).sum()
    )

    add_acceptance("CAN-012", "Broken FactPlan -> DimIssuer keys", 0, broken_plan_issuer, "PASS" if broken_plan_issuer == 0 else "FAIL")
    add_acceptance("CAN-013", "Broken FactPlan -> DimState keys", 0, broken_plan_state, "PASS" if broken_plan_state == 0 else "FAIL")
    add_acceptance("CAN-014", "Broken FactPlan -> DimPlan keys", 0, broken_plan_plan, "PASS" if broken_plan_plan == 0 else "FAIL")
    add_acceptance("CAN-015", "Broken FactIssuer -> DimState keys", 0, broken_issuer_state, "PASS" if broken_issuer_state == 0 else "FAIL")
    add_acceptance("CAN-016", "Broken FactIssuer -> DimIssuer keys", 0, broken_issuer_issuer, "PASS" if broken_issuer_issuer == 0 else "FAIL")

    appeal_recon_df = pd.DataFrame(appeal_recon_rows)
    appeal_mismatches = int(appeal_recon_df["FormulaMismatches"].sum())
    add_acceptance(
        "REC-001",
        "Published appeal percentage formula mismatches",
        0,
        appeal_mismatches,
        "PASS" if appeal_mismatches == 0 else "REVIEW",
        gate_type="RECONCILIATION",
    )

    # ------------------------------------------------------------------
    # DQ register and review summaries.
    # ------------------------------------------------------------------

    dq_df = pd.DataFrame(
        dq_rows,
        columns=[
            "Severity",
            "Scope",
            "EntityKey",
            "SourceField",
            "RuleCode",
            "ObservedValue",
            "ExpectedCondition",
            "SourceSheet",
            "SourceExcelRow",
            "IsKnownSourceException",
            "ResolutionState",
        ],
    )
    if len(dq_df):
        dq_df.insert(
            0, "DQExceptionID", np.arange(1, len(dq_df) + 1)
        )
    else:
        dq_df.insert(0, "DQExceptionID", pd.Series(dtype="int64"))

    diagnostic_df = pd.DataFrame(
        diagnostic_rows,
        columns=[
            "ObservationType",
            "Scope",
            "EntityKey",
            "NetworkScope",
            "ResubmittedValue",
            "DeniedValue",
            "Interpretation",
            "DQStatus",
        ],
    )

    dq_summary_df = (
        dq_df.groupby(
            ["Severity", "Scope", "RuleCode", "ResolutionState"],
            dropna=False,
        )
        .size()
        .reset_index(name="ExceptionCount")
        if len(dq_df)
        else pd.DataFrame(
            columns=[
                "Severity",
                "Scope",
                "RuleCode",
                "ResolutionState",
                "ExceptionCount",
            ]
        )
    )

    availability_combined = pd.concat(
        [
            fact_issuer_metric_availability.assign(EntityLevel="ISSUER"),
            fact_plan_metric_availability.assign(EntityLevel="PLAN"),
        ],
        ignore_index=True,
        sort=False,
    )

    availability_summary_df = (
        availability_combined
        .merge(
            dim_metric[
                ["MetricKey", "MetricCode", "MetricDisplayName"]
            ],
            on="MetricKey",
            how="left",
        )
        .merge(
            dim_availability_status[
                ["StatusKey", "StatusCode"]
            ],
            left_on="AvailabilityStatusKey",
            right_on="StatusKey",
            how="left",
        )
        .groupby(
            [
                "EntityLevel",
                "MetricKey",
                "MetricCode",
                "MetricDisplayName",
                "StatusCode",
            ],
            dropna=False,
        )
        .size()
        .reset_index(name="RecordCount")
        .sort_values(
            ["EntityLevel", "MetricKey", "StatusCode"]
        )
    )

    inventory_df = pd.DataFrame(
        [
            ("dim_reporting_period.csv", len(dim_reporting_period)),
            ("dim_state.csv", len(dim_state)),
            ("dim_issuer.csv", len(dim_issuer)),
            ("dim_plan.csv", len(dim_plan)),
            ("dim_metric.csv", len(dim_metric)),
            ("dim_availability_status.csv", len(dim_availability_status)),
            ("fact_issuer_transparency.csv", len(fact_issuer)),
            ("fact_plan_transparency.csv", len(fact_plan)),
            (
                "fact_issuer_metric_availability.csv",
                len(fact_issuer_metric_availability),
            ),
            (
                "fact_plan_metric_availability.csv",
                len(fact_plan_metric_availability),
            ),
            ("dq_exception_register.csv", len(dq_df)),
        ],
        columns=["CanonicalFile", "Rows"],
    )

    grain_recon_df = pd.DataFrame(
        [
            ("RAW analytical rows", 4956, source_row_count),
            ("Distinct Plan_ID", 4956, raw["Plan_ID"].nunique()),
            ("Distinct Issuer_ID", 348, raw["Issuer_ID"].nunique()),
            ("Distinct State", 30, raw["State"].nunique()),
            ("Canonical plan fact rows", 4956, len(fact_plan)),
            ("Canonical issuer fact rows", 348, len(fact_issuer)),
        ],
        columns=["Check", "Expected", "Actual"],
    )
    grain_recon_df["Status"] = np.where(
        grain_recon_df["Expected"] == grain_recon_df["Actual"],
        "PASS",
        "FAIL",
    )

    source_mapping_rows: list[dict[str, Any]] = []
    for column in EXPECTED_COLUMNS:
        if column in ISSUER_FACT_NAME_BY_SOURCE:
            owner = "FactIssuerTransparency"
            target = ISSUER_FACT_NAME_BY_SOURCE[column]
        elif column in PLAN_FACT_NAME_BY_SOURCE:
            owner = "FactPlanTransparency"
            target = PLAN_FACT_NAME_BY_SOURCE[column]
        else:
            mapping = {
                "Individual/SHOP": ("DimPlan", "MarketSegment"),
                "Exchange_Type": ("DimState", "ExchangeType"),
                "State": ("DimState", "StateCode"),
                "Issuer_Name": ("DimIssuer", "IssuerName"),
                "Issuer_ID": ("DimIssuer", "IssuerID"),
                "Is_Issuer_New_to_Exchange?(Yes_or_No)": (
                    "DimIssuer",
                    "IsIssuerNewToExchange",
                ),
                "SADP_Only": ("DimIssuer", "IsSADPOnlyIssuer"),
                "Plan_ID": ("DimPlan", "PlanID"),
                "Plan_Type": ("DimPlan", "PlanType"),
                "QHP or SADP?": ("DimPlan", "PlanOfferingType"),
                "Metal_Level": ("DimPlan", "MetalLevel"),
                "URL_Claims_Payment_Policies": (
                    "DimIssuer",
                    "ClaimsPaymentPoliciesURL",
                ),
                "Rate_Review": ("DimIssuer", "RateReviewURL"),
                "Financial_Information": (
                    "DimIssuer",
                    "FinancialInformationURL",
                ),
            }
            owner, target = mapping[column]

        source_mapping_rows.append(
            {
                "SourceOrdinal": EXPECTED_COLUMNS.index(column) + 1,
                "SourceColumn": column,
                "CanonicalOwner": owner,
                "CanonicalField": target,
            }
        )

    source_mapping_df = pd.DataFrame(source_mapping_rows)

    acceptance_df = pd.DataFrame(acceptance)
    canonical_structural_fail = bool(
        (
            (acceptance_df["GateType"] == "STRUCTURAL")
            & (acceptance_df["Status"] == "FAIL")
        ).any()
    )

    if canonical_structural_fail:
        build_status = "HOLD_STRUCTURAL_FAILURE"
    elif len(dq_df) > 0:
        build_status = "PASS_WITH_DQ_EXCEPTIONS"
    else:
        build_status = "PASS"

    summary_df = pd.DataFrame(
        [
            ("BuildStatus", build_status),
            ("ScriptVersion", SCRIPT_VERSION),
            ("PUFYear", PUF_YEAR),
            ("ExperienceYear", EXPERIENCE_YEAR),
            ("SourceFile", SOURCE_FILE.name),
            ("SourceSHA256", source_hash),
            ("SourceRows", source_row_count),
            ("States", len(dim_state)),
            ("Issuers", len(dim_issuer)),
            ("Plans", len(dim_plan)),
            ("IssuerMetrics", len(ISSUER_METRICS)),
            ("PlanMetrics", len(PLAN_METRICS)),
            ("DQExceptions", len(dq_df)),
            (
                "KnownSourceExceptionRows",
                int(dq_df["IsKnownSourceException"].sum())
                if len(dq_df)
                else 0,
            ),
            (
                "KnownSourceExceptionEntities",
                int(
                    dq_df.loc[
                        dq_df["IsKnownSourceException"], "EntityKey"
                    ].nunique()
                )
                if len(dq_df)
                else 0,
            ),
            ("DiagnosticObservations", len(diagnostic_df)),
            ("CanonicalWritten", "YES"),
        ],
        columns=["Item", "Value"],
    )

    # ------------------------------------------------------------------
    # Write final project artifacts only after structural gates pass.
    # ------------------------------------------------------------------

    CANONICAL_DIR.mkdir(parents=True, exist_ok=True)

    outputs = {
        "dim_reporting_period.csv": dim_reporting_period,
        "dim_state.csv": dim_state,
        "dim_issuer.csv": dim_issuer,
        "dim_plan.csv": dim_plan,
        "dim_metric.csv": dim_metric,
        "dim_availability_status.csv": dim_availability_status,
        "fact_issuer_transparency.csv": fact_issuer,
        "fact_plan_transparency.csv": fact_plan,
        "fact_issuer_metric_availability.csv": fact_issuer_metric_availability,
        "fact_plan_metric_availability.csv": fact_plan_metric_availability,
        "dq_exception_register.csv": dq_df,
    }

    for filename, dataframe in outputs.items():
        dataframe.to_csv(
            CANONICAL_DIR / filename,
            index=False,
            encoding="utf-8-sig",
        )

    special_status_df = dim_availability_status.copy()

    write_review_workbook(
        summary_df,
        acceptance_df,
        inventory_df,
        dq_summary_df,
        dq_df,
        availability_summary_df,
        appeal_recon_df,
        grain_recon_df,
        source_mapping_df,
        dim_metric,
        special_status_df,
        diagnostic_df,
    )

    print()
    print("Canonical build completed successfully.")
    print(f"Build status   : {build_status}")
    print(f"Canonical files: {CANONICAL_DIR}")
    print(f"Review evidence: {REVIEW_FILE}")
    print("No temporary transformation files were written inside the project.")
    print("Next gate: review 02_CANONICAL_BUILD_REPORT.xlsx before KPI/EDA work.")

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print()
        print("BUILD FAILED")
        print(str(exc))
        sys.exit(1)
