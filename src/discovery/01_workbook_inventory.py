from __future__ import annotations

from pathlib import Path
from datetime import datetime
import hashlib
import json
import os
import re
import sys

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


SCRIPT_VERSION = "2.0.0"
EXPECTED_SHA256 = "27379dc76590027b0e6d21d736331f87376f0ac87810d68e9f7cca555eef4b1f"
SOURCE_FILENAME = "Transparency_in_Coverage_PUF.xlsx"
REVIEW_FOLDER_NAME = "Transparency_PUF_Review"
OUTPUT_FILENAME = "01_SOURCE_DISCOVERY_REPORT.xlsx"

SPECIAL_CODES = {"*", "**", "***", "N/A", "Missing URL"}
DATA_SHEET_PREFIX = "Transparency 2026 - "
REQUIRED_HEADER_MARKERS = {"Plan_ID", "Issuer_ID", "Issuer_Name"}

KEY_COLUMNS = ["Plan_ID", "Issuer_ID", "State"]
CATEGORICAL_COLUMNS = [
    "Individual/SHOP",
    "Exchange_Type",
    "State",
    "Is_Issuer_New_to_Exchange?(Yes_or_No)",
    "SADP_Only",
    "Plan_Type",
    "QHP or SADP?",
    "Metal_Level",
]

APPEAL_RECONCILIATIONS = [
    (
        "Internal Appeals Overturn %",
        "Issuer_Internal_Appeals_Filed",
        "Issuer_Number_Internal_Appeals_Overturned",
        "Issuer_Percent_Internal_Appeals_Overturned",
    ),
    (
        "External Appeals Overturn %",
        "Issuer_External_Appeals_Filed",
        "Issuer_Number_External_Appeals_Overturned",
        "Issuer_Percent_External_Appeals_Overturned",
    ),
]


# -----------------------------------------------------------------------------
# Paths / source integrity
# -----------------------------------------------------------------------------

def project_root() -> Path:
    return Path(__file__).resolve().parents[2]


def review_dir() -> Path:
    configured = os.environ.get("TRANSPARENCY_PUF_REVIEW_DIR")
    if configured:
        return Path(configured).expanduser().resolve()
    return project_root() / ".local-review"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def clean_header(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and np.isnan(value):
        return True
    return isinstance(value, str) and value.strip() == ""


def text_value(value) -> str:
    if is_blank(value):
        return ""
    return str(value).strip()


def normalized_id(value) -> str:
    s = text_value(value)
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def normalized_compare(value) -> str:
    """Normalize display differences without changing business meaning."""
    s = text_value(value)
    if s == "":
        return "<BLANK>"
    if s in SPECIAL_CODES:
        return s
    try:
        n = float(s.replace(",", ""))
        if np.isfinite(n):
            if n.is_integer():
                return str(int(n))
            return format(n, ".15g")
    except (ValueError, TypeError):
        pass
    return s


def numeric_value(value):
    s = text_value(value)
    if s == "" or s in SPECIAL_CODES:
        return np.nan
    try:
        return float(s.replace(",", ""))
    except (ValueError, TypeError):
        return np.nan


def classify_value(value) -> str:
    s = text_value(value)
    if s == "":
        return "BLANK"
    if s in SPECIAL_CODES:
        return s
    try:
        float(s.replace(",", ""))
        return "NUMERIC"
    except (ValueError, TypeError):
        return "TEXT"


def detect_header_row(ws, max_scan_rows: int = 20) -> int:
    """Return zero-based row index for the analytical header."""
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_scan_rows), values_only=True)
    ):
        values = {clean_header(v) for v in row if not is_blank(v)}
        if REQUIRED_HEADER_MARKERS.issubset(values):
            return row_idx
    raise RuntimeError(
        f"Could not identify analytical header in sheet '{ws.title}'. "
        f"Expected markers: {sorted(REQUIRED_HEADER_MARKERS)}"
    )


# -----------------------------------------------------------------------------
# Workbook source discovery
# -----------------------------------------------------------------------------

def inspect_workbook(source: Path):
    wb = openpyxl.load_workbook(source, read_only=False, data_only=False)

    sheet_inventory = []
    preamble_rows = []
    formula_rows = []
    header_rows = {}

    for ws in wb.worksheets:
        is_data = ws.title.startswith(DATA_SHEET_PREFIX)
        header_idx = None
        header_excel_row = None

        if is_data:
            header_idx = detect_header_row(ws)
            header_excel_row = header_idx + 1
            header_rows[ws.title] = header_idx

            for r in range(1, header_excel_row):
                values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                nonblank = [text_value(v) for v in values if not is_blank(v)]
                if nonblank:
                    preamble_rows.append(
                        {
                            "sheet_name": ws.title,
                            "excel_row": r,
                            "content": " | ".join(nonblank),
                        }
                    )

        formula_count = 0
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    formula_count += 1
                    if len(formula_rows) < 500:
                        formula_rows.append(
                            {
                                "sheet_name": ws.title,
                                "cell": cell.coordinate,
                                "formula": str(cell.value),
                            }
                        )

        sheet_inventory.append(
            {
                "sheet_name": ws.title,
                "sheet_state": ws.sheet_state,
                "is_analytical_sheet": is_data,
                "worksheet_max_rows": ws.max_row,
                "worksheet_max_columns": ws.max_column,
                "detected_header_excel_row": header_excel_row,
                "merged_ranges": len(ws.merged_cells.ranges),
                "formula_cells": formula_count,
            }
        )

    props = wb.properties
    metadata = {
        "workbook_title": props.title or "",
        "subject": props.subject or "",
        "creator": props.creator or "",
        "last_modified_by": props.lastModifiedBy or "",
        "created": props.created,
        "modified": props.modified,
        "calculation_mode": getattr(wb.calculation, "calcMode", ""),
    }

    wb.close()
    return sheet_inventory, preamble_rows, formula_rows, header_rows, metadata


def load_data_sheets(source: Path, header_rows: dict[str, int]):
    frames = {}
    for sheet, header_idx in header_rows.items():
        df = pd.read_excel(
            source,
            sheet_name=sheet,
            header=header_idx,
            dtype=object,
            keep_default_na=False,
        )
        df.columns = [clean_header(c) for c in df.columns]
        df = df.loc[:, [c != "" for c in df.columns]].copy()
        blank_rows = df.apply(lambda r: all(is_blank(v) for v in r), axis=1)
        df = df.loc[~blank_rows].reset_index(drop=True)
        frames[sheet] = df
    return frames


# -----------------------------------------------------------------------------
# Profiling
# -----------------------------------------------------------------------------

def semantic_type(series: pd.Series) -> str:
    classes = series.map(classify_value)
    substantive = classes[~classes.isin(["BLANK", "*", "**", "***", "N/A", "Missing URL"])]
    if substantive.empty:
        return "NO_SUBSTANTIVE_VALUES"
    unique_classes = set(substantive.tolist())
    if unique_classes == {"NUMERIC"}:
        return "NUMERIC_WITH_OPTIONAL_STATUS_CODES" if (classes.isin(SPECIAL_CODES).any()) else "NUMERIC"
    if unique_classes == {"TEXT"}:
        return "TEXT_WITH_OPTIONAL_STATUS_CODES" if (classes.isin(SPECIAL_CODES).any()) else "TEXT"
    return "MIXED_NUMERIC_TEXT"


def column_profiles(frames: dict[str, pd.DataFrame]):
    rows = []
    special_rows = []

    for sheet, df in frames.items():
        for pos, column in enumerate(df.columns, start=1):
            s = df[column]
            classes = s.map(classify_value)
            numeric = s.map(numeric_value)
            numeric_nonnull = numeric.dropna()
            substantive_values = [
                text_value(v)
                for v in s
                if text_value(v) != "" and text_value(v) not in SPECIAL_CODES
            ]
            sample_values = []
            seen = set()
            for value in substantive_values:
                if value not in seen:
                    seen.add(value)
                    sample_values.append(value)
                if len(sample_values) == 5:
                    break

            row = {
                "sheet_name": sheet,
                "column_position": pos,
                "column_name": column,
                "semantic_type_candidate": semantic_type(s),
                "rows": len(df),
                "blank_count": int((classes == "BLANK").sum()),
                "star_count": int((classes == "*").sum()),
                "double_star_count": int((classes == "**").sum()),
                "triple_star_count": int((classes == "***").sum()),
                "na_code_count": int((classes == "N/A").sum()),
                "missing_url_count": int((classes == "Missing URL").sum()),
                "numeric_count": int((classes == "NUMERIC").sum()),
                "text_count": int((classes == "TEXT").sum()),
                "distinct_nonblank_raw": int(s.map(text_value).replace("", np.nan).nunique(dropna=True)),
                "numeric_min": float(numeric_nonnull.min()) if not numeric_nonnull.empty else None,
                "numeric_max": float(numeric_nonnull.max()) if not numeric_nonnull.empty else None,
                "numeric_sum": float(numeric_nonnull.sum()) if not numeric_nonnull.empty else None,
                "sample_values": " | ".join(sample_values),
            }
            rows.append(row)

            if any(row[k] > 0 for k in [
                "star_count", "double_star_count", "triple_star_count", "na_code_count", "missing_url_count"
            ]):
                special_rows.append(row.copy())

    return pd.DataFrame(rows), pd.DataFrame(special_rows)


def schema_comparison(frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    all_columns = []
    for df in frames.values():
        for c in df.columns:
            if c not in all_columns:
                all_columns.append(c)

    rows = []
    for position, column in enumerate(all_columns, start=1):
        row = {"column_position_reference": position, "column_name": column}
        types = []
        for sheet, df in frames.items():
            present = column in df.columns
            row[f"present__{sheet}"] = present
            stype = semantic_type(df[column]) if present else "NOT_PRESENT"
            row[f"type__{sheet}"] = stype
            if present:
                types.append(stype)
        row["schema_presence_consistent"] = all(column in df.columns for df in frames.values())
        row["semantic_type_consistent"] = len(set(types)) <= 1
        rows.append(row)
    return pd.DataFrame(rows)


def categorical_profile(frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows = []
    for sheet, df in frames.items():
        for column in CATEGORICAL_COLUMNS:
            if column not in df.columns:
                continue
            values = df[column].map(text_value).replace("", "<BLANK>")
            counts = values.value_counts(dropna=False)
            total = len(df)
            for rank, (value, count) in enumerate(counts.head(50).items(), start=1):
                rows.append(
                    {
                        "sheet_name": sheet,
                        "column_name": column,
                        "rank": rank,
                        "value": value,
                        "count": int(count),
                        "percent_of_sheet_rows": round((count / total) * 100, 4) if total else None,
                    }
                )
    return pd.DataFrame(rows)


# -----------------------------------------------------------------------------
# Keys, grain and DQ
# -----------------------------------------------------------------------------

def add_test(tests, scope, test, value, status, note=""):
    tests.append(
        {
            "scope": scope,
            "test": test,
            "value": value,
            "status": status,
            "note": note,
        }
    )


def key_tests(frames: dict[str, pd.DataFrame]):
    tests = []
    details = []

    for sheet, df in frames.items():
        add_test(tests, sheet, "TOTAL_ANALYTICAL_ROWS", len(df), "INFO")
        exact_dupes = int(df.astype(str).duplicated(keep=False).sum())
        add_test(
            tests,
            sheet,
            "EXACT_DUPLICATE_ROWS",
            exact_dupes,
            "PASS" if exact_dupes == 0 else "REVIEW",
        )

        for key in KEY_COLUMNS:
            if key not in df.columns:
                add_test(tests, sheet, f"{key}_PRESENT", False, "FAIL")
                continue
            vals = df[key].map(normalized_id)
            missing = int((vals == "").sum())
            distinct = int(vals[vals != ""].nunique())
            duplicate_rows = int(vals[(vals != "")].duplicated(keep=False).sum())
            add_test(tests, sheet, f"{key}_MISSING", missing, "PASS" if missing == 0 else "REVIEW")
            add_test(tests, sheet, f"{key}_DISTINCT", distinct, "INFO")
            if key == "Plan_ID":
                add_test(
                    tests,
                    sheet,
                    "Plan_ID_DUPLICATE_ROWS",
                    duplicate_rows,
                    "PASS" if duplicate_rows == 0 else "REVIEW",
                )
                if duplicate_rows:
                    bad = df.loc[vals.duplicated(keep=False) & (vals != "")].copy()
                    bad.insert(0, "normalized_plan_id", vals.loc[bad.index].values)
                    bad.insert(0, "source_sheet", sheet)
                    details.extend(bad.head(500).to_dict("records"))

    combined = pd.concat(
        [df.assign(_source_sheet=sheet) for sheet, df in frames.items()],
        ignore_index=True,
        sort=False,
    )

    plan_ids = combined["Plan_ID"].map(normalized_id)
    issuer_ids = combined["Issuer_ID"].map(normalized_id)
    add_test(tests, "ALL_SHEETS", "TOTAL_ANALYTICAL_ROWS", len(combined), "INFO")
    add_test(tests, "ALL_SHEETS", "DISTINCT_PLAN_ID", int(plan_ids[plan_ids != ""].nunique()), "INFO")
    plan_dup = int(plan_ids[(plan_ids != "")].duplicated(keep=False).sum())
    add_test(tests, "ALL_SHEETS", "PLAN_ID_DUPLICATE_ROWS", plan_dup, "PASS" if plan_dup == 0 else "REVIEW")
    add_test(tests, "ALL_SHEETS", "DISTINCT_ISSUER_ID", int(issuer_ids[issuer_ids != ""].nunique()), "INFO")
    add_test(tests, "ALL_SHEETS", "DISTINCT_STATE", int(combined["State"].map(text_value).replace("", np.nan).nunique()), "INFO")

    if plan_dup:
        bad = combined.loc[plan_ids.duplicated(keep=False) & (plan_ids != "")].copy()
        bad.insert(0, "normalized_plan_id", plan_ids.loc[bad.index].values)
        details.extend(bad.head(500).to_dict("records"))

    detail_df = pd.DataFrame(details)
    return pd.DataFrame(tests), detail_df, combined


def field_role_candidates(columns) -> pd.DataFrame:
    rows = []
    for c in columns:
        if c in {"Individual/SHOP", "Exchange_Type", "State", "Issuer_Name", "Issuer_ID", "Plan_ID", "Plan_Type", "QHP or SADP?", "Metal_Level", "SADP_Only", "Is_Issuer_New_to_Exchange?(Yes_or_No)"}:
            role = "IDENTIFIER_OR_DIMENSION"
            grain = "PLAN ROW / DIMENSION ATTRIBUTE"
        elif str(c).startswith("Issuer_"):
            role = "ISSUER_LEVEL_MEASURE_CANDIDATE"
            grain = "ISSUER (TO BE PROVEN)"
        elif str(c).startswith("Plan_"):
            role = "PLAN_LEVEL_MEASURE_CANDIDATE"
            grain = "PLAN_ID"
        elif c in {"Average Monthly Enrollment", "Average Monthly Disenrollment"}:
            role = "PLAN_LEVEL_MEASURE_CANDIDATE"
            grain = "PLAN_ID"
        elif "URL" in str(c) or c in {"Rate_Review", "Financial_Information"}:
            role = "REFERENCE_OR_URL"
            grain = "PLAN / ISSUER REFERENCE (TO BE CONFIRMED)"
        else:
            role = "UNCLASSIFIED_CANDIDATE"
            grain = "TO BE CONFIRMED"
        rows.append(
            {
                "column_name": c,
                "candidate_role": role,
                "candidate_grain": grain,
                "evidence_state": "DISCOVERY CANDIDATE - NOT YET APPROVED",
            }
        )
    return pd.DataFrame(rows)


def issuer_grain_tests(combined: pd.DataFrame):
    issuer_cols = [
        c
        for c in combined.columns
        if str(c).startswith("Issuer_") and c not in {"Issuer_ID", "Issuer_Name"}
    ]

    work = combined.copy()
    work["_issuer_id_norm"] = work["Issuer_ID"].map(normalized_id)

    summaries = []
    details = []

    for col in issuer_cols:
        temp = work[["_issuer_id_norm", "Issuer_Name", col]].copy()
        temp["_value_norm"] = temp[col].map(normalized_compare)
        grouped = temp.groupby("_issuer_id_norm", dropna=False)

        inconsistent_groups = 0
        max_distinct = 0
        for issuer_id, g in grouped:
            values = sorted(set(g["_value_norm"].tolist()))
            n = len(values)
            max_distinct = max(max_distinct, n)
            if n > 1:
                inconsistent_groups += 1
                details.append(
                    {
                        "Issuer_ID": issuer_id,
                        "Issuer_Name": text_value(g["Issuer_Name"].iloc[0]),
                        "metric": col,
                        "distinct_values": n,
                        "values": " | ".join(values[:20]),
                        "rows": len(g),
                        "status": "REVIEW",
                    }
                )

        summaries.append(
            {
                "metric": col,
                "issuer_groups": int(work["_issuer_id_norm"].nunique()),
                "inconsistent_issuer_groups": inconsistent_groups,
                "max_distinct_values_within_issuer": max_distinct,
                "status": "PASS" if inconsistent_groups == 0 else "REVIEW",
                "interpretation": (
                    "Value is constant within Issuer_ID across plan rows"
                    if inconsistent_groups == 0
                    else "Issuer metric varies within at least one Issuer_ID"
                ),
            }
        )

    return pd.DataFrame(summaries), pd.DataFrame(details)


def dq_findings(frames: dict[str, pd.DataFrame], combined: pd.DataFrame, grain_summary: pd.DataFrame):
    findings = []
    details = []

    def record(scope, rule, count, severity, note, detail_df=None):
        findings.append(
            {
                "scope": scope,
                "rule": rule,
                "count": int(count),
                "severity": severity,
                "status": "PASS" if count == 0 else "REVIEW",
                "note": note,
            }
        )
        if count and detail_df is not None and not detail_df.empty:
            d = detail_df.copy().head(500)
            d.insert(0, "dq_rule", rule)
            d.insert(0, "dq_scope", scope)
            details.extend(d.to_dict("records"))

    # Key completeness / Plan ID internal consistency
    for sheet, df in frames.items():
        for key in ["Plan_ID", "Issuer_ID", "State"]:
            vals = df[key].map(text_value)
            bad = vals == ""
            record(sheet, f"MISSING_{key}", int(bad.sum()), "HIGH", f"{key} should be populated for analytical rows.", df.loc[bad, [key]].copy())

        plan = df["Plan_ID"].map(normalized_id)
        issuer = df["Issuer_ID"].map(normalized_id)
        state = df["State"].map(text_value)

        issuer_prefix_match = pd.Series(
            [p.startswith(i) if p and i else True for p, i in zip(plan, issuer)],
            index=df.index,
            dtype=bool,
        )
        bad = (plan != "") & (issuer != "") & (~issuer_prefix_match)
        record(
            sheet,
            "PLAN_ID_ISSUER_PREFIX_MISMATCH",
            int(bad.sum()),
            "HIGH",
            "Plan_ID should begin with the normalized Issuer_ID in this source structure.",
            df.loc[bad, ["Plan_ID", "Issuer_ID", "Issuer_Name"]].copy(),
        )

        state_from_plan = plan.str.extract(r"^\d+([A-Z]{2})", expand=False).fillna("")
        bad = (plan != "") & (state != "") & (state_from_plan != "") & (state_from_plan != state)
        record(
            sheet,
            "PLAN_ID_STATE_CODE_MISMATCH",
            int(bad.sum()),
            "HIGH",
            "State code embedded in Plan_ID does not match State.",
            df.loc[bad, ["Plan_ID", "State", "Issuer_ID"]].copy(),
        )

        # Negative counts / out-of-range percentages
        for col in df.columns:
            nums = df[col].map(numeric_value)
            numeric_mask = nums.notna()
            if not numeric_mask.any():
                continue
            if any(token in str(col).lower() for token in ["claims", "appeals", "enrollment", "disenrollment", "number_"]):
                bad = numeric_mask & (nums < 0)
                record(
                    sheet,
                    f"NEGATIVE_VALUE::{col}",
                    int(bad.sum()),
                    "HIGH",
                    "Count-like field contains negative numeric values.",
                    df.loc[bad, ["Plan_ID", "Issuer_ID", col]].copy() if bad.any() else None,
                )
            if "percent" in str(col).lower():
                bad = numeric_mask & ((nums < 0) | (nums > 100))
                record(
                    sheet,
                    f"PERCENT_OUTSIDE_0_100::{col}",
                    int(bad.sum()),
                    "HIGH",
                    "Published percentage is outside the expected 0-100 display range.",
                    df.loc[bad, ["Plan_ID", "Issuer_ID", col]].copy() if bad.any() else None,
                )

    # Issuer name consistency by Issuer_ID
    temp = combined[["Issuer_ID", "Issuer_Name"]].copy()
    temp["_issuer"] = temp["Issuer_ID"].map(normalized_id)
    temp["_name"] = temp["Issuer_Name"].map(text_value)
    name_counts = temp.groupby("_issuer")["_name"].nunique(dropna=False)
    bad_issuers = name_counts[name_counts > 1].index
    detail = temp[temp["_issuer"].isin(bad_issuers)].drop_duplicates()
    record(
        "ALL_SHEETS",
        "ISSUER_ID_MULTIPLE_NAMES",
        len(bad_issuers),
        "HIGH",
        "One Issuer_ID maps to multiple Issuer_Name values.",
        detail,
    )

    inconsistent = int((grain_summary["inconsistent_issuer_groups"] > 0).sum()) if not grain_summary.empty else 0
    record(
        "ALL_SHEETS",
        "ISSUER_METRIC_GRAIN_INCONSISTENCY",
        inconsistent,
        "HIGH",
        "At least one issuer-level measure varies across plan rows for the same Issuer_ID.",
    )

    return pd.DataFrame(findings), pd.DataFrame(details)


# -----------------------------------------------------------------------------
# Reconciliation
# -----------------------------------------------------------------------------

def unique_issuer_frame(combined: pd.DataFrame) -> pd.DataFrame:
    work = combined.copy()
    work["_issuer_id_norm"] = work["Issuer_ID"].map(normalized_id)
    work = work.sort_values(["_issuer_id_norm", "_source_sheet", "Plan_ID"], kind="stable")
    return work.groupby("_issuer_id_norm", as_index=False).first()


def appeal_reconciliation(combined: pd.DataFrame):
    issuer = unique_issuer_frame(combined)
    summary_rows = []
    detail_rows = []

    for label, filed_col, overturned_col, published_col in APPEAL_RECONCILIATIONS:
        for _, row in issuer.iterrows():
            filed = numeric_value(row.get(filed_col))
            overturned = numeric_value(row.get(overturned_col))
            published = numeric_value(row.get(published_col))
            raw_published = text_value(row.get(published_col))

            if np.isnan(filed) or np.isnan(overturned) or np.isnan(published) or filed == 0:
                calc = np.nan
                diff = np.nan
                status = "NOT_COMPARABLE"
            else:
                calc = (overturned / filed) * 100
                diff = abs(calc - published)
                status = "PASS" if diff <= 0.011 else "REVIEW"

            logical_status = "NOT_COMPARABLE"
            if not np.isnan(filed) and not np.isnan(overturned):
                logical_status = "PASS" if overturned <= filed else "REVIEW"

            detail_rows.append(
                {
                    "reconciliation": label,
                    "Issuer_ID": row.get("_issuer_id_norm"),
                    "Issuer_Name": row.get("Issuer_Name"),
                    "filed": filed if not np.isnan(filed) else None,
                    "overturned": overturned if not np.isnan(overturned) else None,
                    "published_percent_raw": raw_published,
                    "published_percent_numeric": published if not np.isnan(published) else None,
                    "calculated_percent": round(calc, 6) if not np.isnan(calc) else None,
                    "absolute_difference_pp": round(diff, 6) if not np.isnan(diff) else None,
                    "formula_status": status,
                    "overturned_le_filed_status": logical_status,
                }
            )

        detail = pd.DataFrame([r for r in detail_rows if r["reconciliation"] == label])
        comparable = detail[detail["formula_status"].isin(["PASS", "REVIEW"])]
        summary_rows.append(
            {
                "reconciliation": label,
                "issuer_population": len(detail),
                "formula_comparable_issuers": len(comparable),
                "formula_pass": int((comparable["formula_status"] == "PASS").sum()),
                "formula_review": int((comparable["formula_status"] == "REVIEW").sum()),
                "non_comparable": int((detail["formula_status"] == "NOT_COMPARABLE").sum()),
                "overturned_gt_filed_review": int((detail["overturned_le_filed_status"] == "REVIEW").sum()),
                "status": "PASS" if int((comparable["formula_status"] == "REVIEW").sum()) == 0 else "REVIEW",
                "formula": "overturned / filed * 100",
                "tolerance_percentage_points": 0.011,
            }
        )

    return pd.DataFrame(summary_rows), pd.DataFrame(detail_rows)


# -----------------------------------------------------------------------------
# Report composition / formatting
# -----------------------------------------------------------------------------

def dataframe_or_message(df: pd.DataFrame, message: str) -> pd.DataFrame:
    return df if not df.empty else pd.DataFrame([{"message": message}])


def build_readme(source: Path, source_hash: str, output: Path) -> pd.DataFrame:
    rows = [
        ("Report", "Transparency in Coverage PUF — Source Discovery Report"),
        ("Evidence State", "DISCOVERY — NOT YET CANONICAL"),
        ("Script Version", SCRIPT_VERSION),
        ("Source File", str(source)),
        ("Source SHA256", source_hash),
        ("Expected SHA256", EXPECTED_SHA256),
        ("Integrity", "PASS" if source_hash == EXPECTED_SHA256 else "FAIL"),
        ("Output File", str(output)),
        ("Generated At", datetime.now().astimezone().isoformat(timespec="seconds")),
        ("Python", sys.version.split()[0]),
        ("pandas", pd.__version__),
        ("openpyxl", openpyxl.__version__),
        ("numpy", np.__version__),
        ("Scope", "Read-only source discovery, profiling, grain tests, DQ screening and reconciliation."),
        ("Not Performed", "No cleaning, imputation, destructive deduplication, KPI approval, canonical transformation, SQL modeling or BI development."),
        ("Special Value Rule", "*, **, ***, N/A, Missing URL and blank are preserved as distinct source states."),
        ("Grain Rule", "Issuer-level and plan-level fields are tested separately; no issuer measure is summed across plan rows."),
    ]
    return pd.DataFrame(rows, columns=["item", "value"])


def build_exec_summary(frames, combined, key_df, grain_df, dq_df, recon_df) -> pd.DataFrame:
    rows = []
    rows.append(("Analytical sheets", len(frames), "INFO"))
    rows.append(("Total analytical rows", len(combined), "INFO"))
    rows.append(("Distinct Plan_ID", int(combined["Plan_ID"].map(normalized_id).replace("", np.nan).nunique()), "INFO"))
    rows.append(("Distinct Issuer_ID", int(combined["Issuer_ID"].map(normalized_id).replace("", np.nan).nunique()), "INFO"))
    rows.append(("Distinct states", int(combined["State"].map(text_value).replace("", np.nan).nunique()), "INFO"))
    rows.append(("Exact schema across analytical sheets", bool(all(list(df.columns) == list(next(iter(frames.values())).columns) for df in frames.values())), "PASS" if all(list(df.columns) == list(next(iter(frames.values())).columns) for df in frames.values()) else "REVIEW"))
    rows.append(("Cross-sheet duplicate Plan_ID rows", int(key_df.loc[(key_df["scope"] == "ALL_SHEETS") & (key_df["test"] == "PLAN_ID_DUPLICATE_ROWS"), "value"].iloc[0]), str(key_df.loc[(key_df["scope"] == "ALL_SHEETS") & (key_df["test"] == "PLAN_ID_DUPLICATE_ROWS"), "status"].iloc[0])))
    rows.append(("Issuer metrics with within-issuer inconsistency", int((grain_df["status"] == "REVIEW").sum()) if not grain_df.empty else 0, "PASS" if grain_df.empty or (grain_df["status"] == "REVIEW").sum() == 0 else "REVIEW"))
    rows.append(("DQ rules requiring review", int((dq_df["status"] == "REVIEW").sum()) if not dq_df.empty else 0, "PASS" if dq_df.empty or (dq_df["status"] == "REVIEW").sum() == 0 else "REVIEW"))
    rows.append(("Reconciliation checks requiring review", int((recon_df["status"] == "REVIEW").sum()) if not recon_df.empty else 0, "PASS" if recon_df.empty or (recon_df["status"] == "REVIEW").sum() == 0 else "REVIEW"))
    rows.append(("Automated checkpoint state", "DISCOVERY COMPLETE — HUMAN REVIEW REQUIRED", "INFO"))
    return pd.DataFrame(rows, columns=["metric", "value", "status"])


def style_workbook(path: Path):
    wb = openpyxl.load_workbook(path)

    header_fill = PatternFill("solid", fgColor="17365D")
    header_font = Font(color="FFFFFF", bold=True)
    thin_gray = Side(style="thin", color="D9E2F3")
    border = Border(bottom=thin_gray)
    status_fills = {
        "PASS": PatternFill("solid", fgColor="E2F0D9"),
        "REVIEW": PatternFill("solid", fgColor="FFF2CC"),
        "FAIL": PatternFill("solid", fgColor="F4CCCC"),
        "INFO": PatternFill("solid", fgColor="D9EAF7"),
        "NOT_COMPARABLE": PatternFill("solid", fgColor="E7E6E6"),
    }

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        if ws.max_row >= 1:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        # Practical widths, capped to avoid unusably wide sheets.
        for col_idx in range(1, ws.max_column + 1):
            letter = get_column_letter(col_idx)
            max_len = 0
            for row_idx in range(1, min(ws.max_row, 250) + 1):
                value = ws.cell(row_idx, col_idx).value
                if value is None:
                    continue
                max_len = max(max_len, len(str(value)))
            ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 55)

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if cell.value in status_fills:
                    cell.fill = status_fills[cell.value]
                    cell.font = Font(bold=True)
                if isinstance(cell.value, str) and len(cell.value) > 80:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)

        ws.row_dimensions[1].height = 32

    wb.save(path)


def write_report(
    output: Path,
    report_sheets: list[tuple[str, pd.DataFrame]],
):
    if output.exists():
        try:
            output.unlink()
        except PermissionError as exc:
            raise PermissionError(
                f"Cannot overwrite '{output}'. Close the Excel report if it is open, then run again."
            ) from exc

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in report_sheets:
            dataframe_or_message(df, "No records for this section.").to_excel(
                writer,
                sheet_name=sheet_name[:31],
                index=False,
            )

    style_workbook(output)


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------

def main():
    root = project_root()
    source = root / "data" / "raw" / SOURCE_FILENAME
    output_dir = review_dir()
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / OUTPUT_FILENAME

    print("=" * 78)
    print("Transparency in Coverage PUF — Governed Source Discovery")
    print("=" * 78)
    print(f"Script version : {SCRIPT_VERSION}")
    print(f"Source         : {source}")
    print(f"Review folder  : {output_dir}")

    if not source.exists():
        raise FileNotFoundError(f"RAW source not found: {source}")

    source_hash = sha256(source)
    print(f"SHA256         : {source_hash}")
    if source_hash.lower() != EXPECTED_SHA256.lower():
        raise RuntimeError(
            "RAW source integrity check FAILED. The current RAW workbook hash does not match "
            "the approved baseline. Stop and review the source before continuing."
        )
    print("Source integrity: PASS")

    sheet_inventory, preamble_rows, formula_rows, header_rows, workbook_metadata = inspect_workbook(source)
    frames = load_data_sheets(source, header_rows)

    if not frames:
        raise RuntimeError("No analytical sheets were detected.")

    print("Analytical sheets detected:")
    for sheet, df in frames.items():
        print(f"  - {sheet}: {len(df):,} rows x {len(df.columns)} columns")

    col_profile, special_profile = column_profiles(frames)
    schema_df = schema_comparison(frames)
    categorical_df = categorical_profile(frames)
    key_df, key_detail_df, combined = key_tests(frames)
    role_df = field_role_candidates(next(iter(frames.values())).columns)
    grain_df, grain_detail_df = issuer_grain_tests(combined)
    dq_df, dq_detail_df = dq_findings(frames, combined, grain_df)
    recon_summary_df, recon_detail_df = appeal_reconciliation(combined)

    workbook_meta_df = pd.DataFrame(
        [
            {"item": "file_name", "value": source.name},
            {"item": "file_size_bytes", "value": source.stat().st_size},
            {"item": "source_sha256", "value": source_hash},
            {"item": "last_modified", "value": datetime.fromtimestamp(source.stat().st_mtime).astimezone().isoformat(timespec="seconds")},
        ]
        + [{"item": k, "value": v} for k, v in workbook_metadata.items()]
    )

    sheet_inventory_df = pd.DataFrame(sheet_inventory)
    preamble_df = pd.DataFrame(preamble_rows)
    formula_df = pd.DataFrame(formula_rows)

    disclaimer_df = pd.read_excel(
        source,
        sheet_name="PUF Data Disclaimer",
        header=None,
        dtype=object,
        keep_default_na=False,
    )
    disclaimer_df.columns = ["disclaimer"]
    disclaimer_df = disclaimer_df[disclaimer_df["disclaimer"].map(text_value) != ""].reset_index(drop=True)

    sample_frames = []
    for sheet, df in frames.items():
        sample = df.head(5).copy()
        sample.insert(0, "source_sheet", sheet)
        sample_frames.append(sample)
    samples_df = pd.concat(sample_frames, ignore_index=True, sort=False)

    readme_df = build_readme(source, source_hash, output)
    exec_df = build_exec_summary(frames, combined, key_df, grain_df, dq_df, recon_summary_df)

    report_sheets = [
        ("00_README", readme_df),
        ("01_EXEC_SUMMARY", exec_df),
        ("02_WORKBOOK_META", workbook_meta_df),
        ("03_SHEET_INVENTORY", sheet_inventory_df),
        ("04_SOURCE_PREAMBLE", preamble_df),
        ("05_SCHEMA_COMPARE", schema_df),
        ("06_FIELD_ROLES", role_df),
        ("07_COLUMN_PROFILE", col_profile),
        ("08_SPECIAL_VALUES", special_profile),
        ("09_CATEGORY_PROFILE", categorical_df),
        ("10_KEY_TESTS", key_df),
        ("11_KEY_DETAILS", key_detail_df),
        ("12_ISSUER_GRAIN", grain_df),
        ("13_GRAIN_DETAILS", grain_detail_df),
        ("14_DQ_FINDINGS", dq_df),
        ("15_DQ_DETAILS", dq_detail_df),
        ("16_RECON_SUMMARY", recon_summary_df),
        ("17_RECON_DETAIL", recon_detail_df),
        ("18_FORMULA_CELLS", formula_df),
        ("19_SOURCE_SAMPLES", samples_df),
        ("20_DISCLAIMER", disclaimer_df),
    ]

    write_report(output, report_sheets)

    print()
    print("Source discovery completed successfully.")
    print(f"Output: {output}")
    print("No discovery output files were written inside the project.")
    print("Evidence state: DISCOVERY — NOT YET CANONICAL")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print()
        print("SOURCE DISCOVERY FAILED")
        print(f"{type(exc).__name__}: {exc}")
        sys.exit(1)
