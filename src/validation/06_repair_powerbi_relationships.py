from __future__ import annotations

from pathlib import Path
from datetime import datetime
import hashlib
import os
import re
import shutil
import subprocess
import sys
import uuid
import zipfile

import pandas as pd


SCRIPT_VERSION = "1.0.0"

PROJECT_ROOT = Path(__file__).resolve().parents[2]
POWERBI_ROOT = PROJECT_ROOT / "powerbi"


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


REVIEW_DIR = get_windows_desktop() / "Transparency_PUF_Review"
REPORT_FILE = REVIEW_DIR / "06_POWER_BI_RELATIONSHIP_REPAIR_REPORT.xlsx"
BACKUP_FILE = REVIEW_DIR / "06_PRE_RELATIONSHIP_MODEL_BACKUP.zip"


EXPECTED_BUSINESS_TABLES = {
    "DimAvailabilityStatus",
    "DimIssuer",
    "DimMetric",
    "DimPlan",
    "DimReportingPeriod",
    "DimState",
    "DQExceptionRegister",
    "FactIssuerMetricAvailability",
    "FactIssuerTransparency",
    "FactPlanMetricAvailability",
    "FactPlanTransparency",
}

AUTO_DATE_PREFIXES = (
    "DateTableTemplate_",
    "LocalDateTable_",
)

# Child/many-side column -> parent/one-side column.
# All relationships are active and single-direction by TMDL defaults.
GOVERNED_RELATIONSHIPS = [
    (
        "State filters Issuers",
        "DimIssuer.StateKey",
        "DimState.StateKey",
    ),
    (
        "Issuer filters Plans",
        "DimPlan.IssuerKey",
        "DimIssuer.IssuerKey",
    ),
    (
        "Issuer filters Issuer Transparency",
        "FactIssuerTransparency.IssuerKey",
        "DimIssuer.IssuerKey",
    ),
    (
        "Issuer filters Issuer Metric Availability",
        "FactIssuerMetricAvailability.IssuerKey",
        "DimIssuer.IssuerKey",
    ),
    (
        "Plan filters Plan Transparency",
        "FactPlanTransparency.PlanKey",
        "DimPlan.PlanKey",
    ),
    (
        "Plan filters Plan Metric Availability",
        "FactPlanMetricAvailability.PlanKey",
        "DimPlan.PlanKey",
    ),
    (
        "Period filters Issuer Transparency",
        "FactIssuerTransparency.ReportingPeriodKey",
        "DimReportingPeriod.ReportingPeriodKey",
    ),
    (
        "Period filters Plan Transparency",
        "FactPlanTransparency.ReportingPeriodKey",
        "DimReportingPeriod.ReportingPeriodKey",
    ),
    (
        "Period filters Issuer Metric Availability",
        "FactIssuerMetricAvailability.ReportingPeriodKey",
        "DimReportingPeriod.ReportingPeriodKey",
    ),
    (
        "Period filters Plan Metric Availability",
        "FactPlanMetricAvailability.ReportingPeriodKey",
        "DimReportingPeriod.ReportingPeriodKey",
    ),
    (
        "Metric filters Issuer Metric Availability",
        "FactIssuerMetricAvailability.MetricKey",
        "DimMetric.MetricKey",
    ),
    (
        "Metric filters Plan Metric Availability",
        "FactPlanMetricAvailability.MetricKey",
        "DimMetric.MetricKey",
    ),
    (
        "Availability Status filters Issuer Metric Availability",
        "FactIssuerMetricAvailability.AvailabilityStatusKey",
        "DimAvailabilityStatus.StatusKey",
    ),
    (
        "Availability Status filters Plan Metric Availability",
        "FactPlanMetricAvailability.AvailabilityStatusKey",
        "DimAvailabilityStatus.StatusKey",
    ),
]


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8")


def find_semantic_model() -> Path:
    candidates = sorted(
        p for p in POWERBI_ROOT.glob("*.SemanticModel") if p.is_dir()
    )
    if len(candidates) != 1:
        raise RuntimeError(
            "Expected exactly one *.SemanticModel folder under "
            f"{POWERBI_ROOT}; found {len(candidates)}."
        )
    return candidates[0]


def power_bi_desktop_running() -> bool:
    if os.name != "nt":
        return False
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq PBIDesktop.exe"],
        capture_output=True,
        text=True,
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    return "PBIDesktop.exe" in result.stdout


def extract_ref_tables(model_text: str) -> list[str]:
    return re.findall(r"(?m)^ref table (.+?)\s*$", model_text)


def extract_relationships(text: str) -> list[dict[str, str]]:
    rows = []
    blocks = re.split(r"(?m)(?=^relationship\s+)", text.strip())
    for block in blocks:
        block = block.strip()
        if not block.startswith("relationship "):
            continue

        name_match = re.match(r"relationship\s+(.+)", block)
        from_match = re.search(r"(?m)^\s*fromColumn:\s*(.+?)\s*$", block)
        to_match = re.search(r"(?m)^\s*toColumn:\s*(.+?)\s*$", block)
        active_match = re.search(r"(?m)^\s*isActive:\s*(.+?)\s*$", block)
        cross_match = re.search(
            r"(?m)^\s*crossFilteringBehavior:\s*(.+?)\s*$",
            block,
        )
        from_card_match = re.search(
            r"(?m)^\s*fromCardinality:\s*(.+?)\s*$",
            block,
        )
        to_card_match = re.search(
            r"(?m)^\s*toCardinality:\s*(.+?)\s*$",
            block,
        )

        rows.append(
            {
                "RelationshipName": (
                    name_match.group(1).strip() if name_match else ""
                ),
                "FromColumn": (
                    from_match.group(1).strip() if from_match else ""
                ),
                "ToColumn": to_match.group(1).strip() if to_match else "",
                "IsActive": (
                    active_match.group(1).strip()
                    if active_match
                    else "true (default)"
                ),
                "CrossFilteringBehavior": (
                    cross_match.group(1).strip()
                    if cross_match
                    else "oneDirection (default)"
                ),
                "FromCardinality": (
                    from_card_match.group(1).strip()
                    if from_card_match
                    else "many (default)"
                ),
                "ToCardinality": (
                    to_card_match.group(1).strip()
                    if to_card_match
                    else "one (default)"
                ),
                "RawDefinition": block,
            }
        )
    return rows


def remove_indented_block(
    text: str,
    start_pattern: re.Pattern[str],
) -> tuple[str, int]:
    lines = text.splitlines()
    output = []
    removed = 0
    i = 0

    while i < len(lines):
        line = lines[i]
        if not start_pattern.match(line):
            output.append(line)
            i += 1
            continue

        base_indent = len(
            line[: len(line) - len(line.lstrip("\t "))]
            .replace("\t", "    ")
        )
        removed += 1
        i += 1

        # Remove children of this TMDL block.
        while i < len(lines):
            current = lines[i]
            if current.strip() == "":
                # Blank lines immediately inside the removed block are dropped.
                i += 1
                continue

            current_indent = len(
                current[: len(current) - len(current.lstrip("\t "))]
                .replace("\t", "    ")
            )
            if current_indent <= base_indent:
                break
            i += 1

    return "\n".join(output).rstrip() + "\n", removed


def deterministic_relationship_id(label: str) -> str:
    namespace = uuid.UUID("f860d7f4-2b1c-4ac5-8108-cbd6a3657269")
    return str(uuid.uuid5(namespace, label))


def governed_relationship_text() -> str:
    blocks = []
    for label, child, parent in GOVERNED_RELATIONSHIPS:
        rel_id = deterministic_relationship_id(label)
        blocks.append(
            "\n".join(
                [
                    f"relationship {rel_id}",
                    f"\tfromColumn: {child}",
                    f"\ttoColumn: {parent}",
                ]
            )
        )
    return "\n\n".join(blocks) + "\n"


def create_backup(semantic_model: Path) -> None:
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)
    if BACKUP_FILE.exists():
        BACKUP_FILE.unlink()

    with zipfile.ZipFile(
        BACKUP_FILE,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as zf:
        for path in sorted(p for p in semantic_model.rglob("*") if p.is_file()):
            arcname = Path(semantic_model.name) / path.relative_to(semantic_model)
            zf.write(path, arcname=str(arcname))


def style_report(path: Path) -> None:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = load_workbook(path)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        if ws.max_row and ws.max_column:
            ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        for col_idx in range(1, ws.max_column + 1):
            max_len = 0
            for row_idx in range(1, min(ws.max_row, 200) + 1):
                value = ws.cell(row_idx, col_idx).value
                if value is not None:
                    max_len = max(max_len, len(str(value)))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(
                max(max_len + 2, 10),
                60,
            )

        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    wb.save(path)


def main() -> int:
    print("=" * 78)
    print("Transparency in Coverage PUF — Power BI Relationship Governance Repair")
    print("=" * 78)
    print(f"Script version : {SCRIPT_VERSION}")
    print(f"Project root   : {PROJECT_ROOT}")
    print(f"Power BI root  : {POWERBI_ROOT}")
    print(f"Report         : {REPORT_FILE}")
    print(f"Backup         : {BACKUP_FILE}")

    if power_bi_desktop_running():
        raise RuntimeError(
            "Power BI Desktop is currently running. "
            "Save the PBIP, close Power BI Desktop completely, "
            "then run this script again."
        )

    semantic_model = find_semantic_model()
    definition_dir = semantic_model / "definition"
    tables_dir = definition_dir / "tables"
    model_file = definition_dir / "model.tmdl"
    relationships_file = definition_dir / "relationships.tmdl"
    reporting_period_file = tables_dir / "DimReportingPeriod.tmdl"

    required_files = [
        model_file,
        relationships_file,
        reporting_period_file,
    ]
    for path in required_files:
        if not path.exists():
            raise FileNotFoundError(f"Required PBIP file missing: {path}")

    before_relationship_text = read_text(relationships_file)
    before_relationships = extract_relationships(before_relationship_text)
    before_relationship_df = pd.DataFrame(before_relationships)

    model_text = read_text(model_file)
    refs_before = extract_ref_tables(model_text)
    business_refs_before = {
        r for r in refs_before if not r.startswith(AUTO_DATE_PREFIXES)
    }

    if business_refs_before != EXPECTED_BUSINESS_TABLES:
        missing = EXPECTED_BUSINESS_TABLES - business_refs_before
        extra = business_refs_before - EXPECTED_BUSINESS_TABLES
        raise RuntimeError(
            "PBIP business-table inventory differs from the approved model.\n"
            f"Missing: {sorted(missing)}\n"
            f"Unexpected: {sorted(extra)}\n"
            "No changes were made."
        )

    current_relationship_pairs = {
        (r["FromColumn"], r["ToColumn"])
        for r in before_relationships
    }

    expected_current_markers = {
        ("DimIssuer.StateKey", "DimState.StateKey"),
        ("DimPlan.IssuerKey", "DimIssuer.IssuerKey"),
        ("FactPlanTransparency.PlanKey", "DimPlan.PlanKey"),
        ("FactIssuerMetricAvailability.MetricKey", "DimMetric.MetricKey"),
    }
    if not expected_current_markers.issubset(current_relationship_pairs):
        raise RuntimeError(
            "Current relationship source does not match the reviewed "
            "05_POWER_BI_MODEL_STATE baseline. No changes were made."
        )

    create_backup(semantic_model)

    before_files = {
        str(p.relative_to(semantic_model)): sha256(p)
        for p in semantic_model.rglob("*")
        if p.is_file()
    }

    actions = []

    # 1. Disable Auto Date/Time in the semantic source.
    if re.search(
        r"(?m)^annotation __PBI_TimeIntelligenceEnabled = 1\s*$",
        model_text,
    ):
        model_text = re.sub(
            r"(?m)^annotation __PBI_TimeIntelligenceEnabled = 1\s*$",
            "annotation __PBI_TimeIntelligenceEnabled = 0",
            model_text,
        )
        actions.append(
            ("MODEL", "Disable Auto Date/Time", "APPLIED")
        )
    else:
        actions.append(
            (
                "MODEL",
                "Disable Auto Date/Time",
                "ALREADY_DISABLED_OR_NOT_PRESENT",
            )
        )

    # 2. Remove auto-date table refs from model.tmdl.
    original_model_text = model_text
    model_lines = []
    removed_model_refs = []
    for line in model_text.splitlines():
        match = re.match(r"^ref table (.+?)\s*$", line)
        if match and match.group(1).startswith(AUTO_DATE_PREFIXES):
            removed_model_refs.append(match.group(1))
            continue
        model_lines.append(line)

    model_text = "\n".join(model_lines).rstrip() + "\n"
    write_text(model_file, model_text)

    actions.append(
        (
            "MODEL",
            "Remove Auto Date/Time table references",
            f"REMOVED {len(removed_model_refs)}",
        )
    )

    # 3. Remove the local-date variation metadata from period date columns.
    reporting_text = read_text(reporting_period_file)
    reporting_text, variations_removed = remove_indented_block(
        reporting_text,
        re.compile(r"^\t\tvariation\s+Variation\s*$"),
    )
    write_text(reporting_period_file, reporting_text)
    actions.append(
        (
            "TABLE",
            "Remove LocalDate variation blocks from DimReportingPeriod",
            f"REMOVED {variations_removed}",
        )
    )

    # 4. Delete only known auto-date TMDL table files.
    deleted_auto_date_files = []
    for path in sorted(tables_dir.glob("*.tmdl")):
        if path.stem.startswith(AUTO_DATE_PREFIXES):
            deleted_auto_date_files.append(path.name)
            path.unlink()

    actions.append(
        (
            "TABLE",
            "Delete auto-generated date table TMDL files",
            f"DELETED {len(deleted_auto_date_files)}",
        )
    )

    # 5. Replace all auto-detected relationships with governed relationships.
    governed_text = governed_relationship_text()
    write_text(relationships_file, governed_text)
    actions.append(
        (
            "RELATIONSHIP",
            "Replace auto-detected model relationships",
            (
                f"{len(before_relationships)} BEFORE -> "
                f"{len(GOVERNED_RELATIONSHIPS)} GOVERNED"
            ),
        )
    )

    # -----------------------------------------------------------------
    # Post-change validation
    # -----------------------------------------------------------------

    after_model_text = read_text(model_file)
    after_refs = extract_ref_tables(after_model_text)

    after_relationship_text = read_text(relationships_file)
    after_relationships = extract_relationships(after_relationship_text)
    after_relationship_df = pd.DataFrame(after_relationships)

    expected_pairs = {
        (child, parent)
        for _, child, parent in GOVERNED_RELATIONSHIPS
    }
    actual_pairs = {
        (r["FromColumn"], r["ToColumn"])
        for r in after_relationships
    }

    validation_rows = []

    def check(check_id: str, test_name: str, expected, actual, passed: bool):
        validation_rows.append(
            {
                "CheckID": check_id,
                "TestName": test_name,
                "Expected": expected,
                "Actual": actual,
                "Status": "PASS" if passed else "FAIL",
            }
        )

    check(
        "REL-001",
        "Governed relationship count",
        len(GOVERNED_RELATIONSHIPS),
        len(after_relationships),
        len(after_relationships) == len(GOVERNED_RELATIONSHIPS),
    )
    check(
        "REL-002",
        "Relationship endpoint set",
        "Exact governed set",
        "Exact" if actual_pairs == expected_pairs else "Mismatch",
        actual_pairs == expected_pairs,
    )
    check(
        "REL-003",
        "Inactive relationships",
        0,
        sum(
            str(r["IsActive"]).lower() == "false"
            for r in after_relationships
        ),
        all(
            str(r["IsActive"]).lower() != "false"
            for r in after_relationships
        ),
    )
    check(
        "REL-004",
        "Bidirectional relationships",
        0,
        sum(
            "both" in str(r["CrossFilteringBehavior"]).lower()
            for r in after_relationships
        ),
        all(
            "both" not in str(r["CrossFilteringBehavior"]).lower()
            for r in after_relationships
        ),
    )
    check(
        "REL-005",
        "Auto-date table refs",
        0,
        sum(r.startswith(AUTO_DATE_PREFIXES) for r in after_refs),
        not any(r.startswith(AUTO_DATE_PREFIXES) for r in after_refs),
    )
    check(
        "REL-006",
        "Auto-date TMDL files",
        0,
        sum(
            p.stem.startswith(AUTO_DATE_PREFIXES)
            for p in tables_dir.glob("*.tmdl")
        ),
        not any(
            p.stem.startswith(AUTO_DATE_PREFIXES)
            for p in tables_dir.glob("*.tmdl")
        ),
    )
    check(
        "REL-007",
        "Auto Date/Time annotation",
        0,
        1
        if "__PBI_TimeIntelligenceEnabled = 1" in after_model_text
        else 0,
        "__PBI_TimeIntelligenceEnabled = 1" not in after_model_text,
    )
    check(
        "REL-008",
        "DQExceptionRegister relationships",
        0,
        sum(
            "DQExceptionRegister" in r["FromColumn"]
            or "DQExceptionRegister" in r["ToColumn"]
            for r in after_relationships
        ),
        all(
            "DQExceptionRegister" not in r["FromColumn"]
            and "DQExceptionRegister" not in r["ToColumn"]
            for r in after_relationships
        ),
    )

    validation_df = pd.DataFrame(validation_rows)
    failures = int((validation_df["Status"] == "FAIL").sum())

    after_files = {
        str(p.relative_to(semantic_model)): sha256(p)
        for p in semantic_model.rglob("*")
        if p.is_file()
    }

    file_changes = []
    all_paths = sorted(set(before_files) | set(after_files))
    for rel_path in all_paths:
        before_hash = before_files.get(rel_path, "")
        after_hash = after_files.get(rel_path, "")
        if before_hash == after_hash:
            state = "UNCHANGED"
        elif before_hash and after_hash:
            state = "MODIFIED"
        elif before_hash and not after_hash:
            state = "DELETED"
        else:
            state = "CREATED"

        if state != "UNCHANGED":
            file_changes.append(
                {
                    "RelativePath": rel_path,
                    "ChangeState": state,
                    "BeforeSHA256": before_hash,
                    "AfterSHA256": after_hash,
                }
            )

    governed_df = pd.DataFrame(
        [
            {
                "RelationshipPurpose": label,
                "ManySide_FromColumn": child,
                "OneSide_ToColumn": parent,
                "Cardinality": "Many-to-One",
                "CrossFilter": "Single direction",
                "Active": True,
            }
            for label, child, parent in GOVERNED_RELATIONSHIPS
        ]
    )

    actions_df = pd.DataFrame(
        actions,
        columns=["Area", "Action", "Result"],
    )

    summary_df = pd.DataFrame(
        [
            ("RepairStatus", "PASS" if failures == 0 else "FAIL"),
            ("ScriptVersion", SCRIPT_VERSION),
            ("ExecutedAtLocal", datetime.now().isoformat(timespec="seconds")),
            ("SemanticModel", str(semantic_model)),
            ("RelationshipsBefore", len(before_relationships)),
            ("RelationshipsAfter", len(after_relationships)),
            ("InactiveRelationshipsAfter", 0),
            ("BidirectionalRelationshipsAfter", 0),
            ("AutoDateTablesAfter", 0),
            ("DQRegisterConnected", "NO"),
            ("ValidationChecks", len(validation_df)),
            ("ValidationFailures", failures),
            ("BackupFile", str(BACKUP_FILE)),
            ("EvidenceState", "PBIP SOURCE MODIFIED + OFFLINE VALIDATED"),
        ],
        columns=["Item", "Value"],
    )

    REVIEW_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(REPORT_FILE, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="00_Summary", index=False)
        validation_df.to_excel(
            writer,
            sheet_name="01_Validation",
            index=False,
        )
        governed_df.to_excel(
            writer,
            sheet_name="02_Governed_Relationships",
            index=False,
        )
        before_relationship_df.to_excel(
            writer,
            sheet_name="03_Before_Relationships",
            index=False,
        )
        after_relationship_df.to_excel(
            writer,
            sheet_name="04_After_Relationships",
            index=False,
        )
        actions_df.to_excel(
            writer,
            sheet_name="05_Actions",
            index=False,
        )
        pd.DataFrame(file_changes).to_excel(
            writer,
            sheet_name="06_File_Changes",
            index=False,
        )
        pd.DataFrame(
            {"DeletedAutoDateFile": deleted_auto_date_files}
        ).to_excel(
            writer,
            sheet_name="07_Deleted_AutoDate",
            index=False,
        )

    style_report(REPORT_FILE)

    print()
    print("Power BI relationship governance repair completed.")
    print(f"Repair status          : {'PASS' if failures == 0 else 'FAIL'}")
    print(f"Relationships before   : {len(before_relationships)}")
    print(f"Relationships after    : {len(after_relationships)}")
    print(f"Inactive after         : 0")
    print(f"Bidirectional after    : 0")
    print(f"Auto-date tables after : 0")
    print(f"Validation failures    : {failures}")
    print(f"Backup                 : {BACKUP_FILE}")
    print(f"Review evidence        : {REPORT_FILE}")

    if failures:
        print()
        print("VALIDATION FAILED. Do not open/save the PBIP until reviewed.")
        return 2

    print()
    print("Next step: open the PBIP in Power BI Desktop, allow it to load,")
    print("then Save (Ctrl+S). Do not create measures or visuals yet.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print()
        print("REPAIR FAILED")
        print(str(exc))
        sys.exit(1)
