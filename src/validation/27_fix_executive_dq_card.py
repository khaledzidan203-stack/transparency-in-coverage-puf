from __future__ import annotations

from datetime import datetime
from pathlib import Path
import copy
import json
import os
import subprocess
import sys
import zipfile

import pandas as pd


SCRIPT_VERSION = "1.0.0"

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_DIR = (
    PROJECT_ROOT
    / "powerbi"
    / "TransparencyInCoverage.Report"
)
PAGES_DIR = REPORT_DIR / "definition" / "pages"

EXPECTED_VISUALS = 48
EXPECTED_BOUND_VISUALS = 20

TARGET_CARD = "p01_dq_global"
TARGET_BAND = "p01_band_dq_global"
NEXT_ROW_ANCHOR = "p01_band_kpi_issuers"

BAND_HEIGHT = 24.0
BODY_GAP = 2.0
BOTTOM_CLEARANCE = 8.0
VALUE_FONT_SIZE = 14.0

CANVAS_WIDTH = 1920
CANVAS_HEIGHT = 1080

COLORS = {
    "amber": "#C8872C",
    "amber_dark": "#9C641A",
    "white": "#FFFFFF",
    "border": "#E7ECF2",
}

UNSUPPORTED_PAGE_ROOT_PROPERTIES = {
    "horizontalAlignment",
    "verticalAlignment",
}


def get_windows_desktop() -> Path:
    if os.name == "nt":
        try:
            import winreg
            key_path = (
                r"Software\Microsoft\Windows\CurrentVersion"
                r"\Explorer\User Shell Folders"
            )
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                key_path,
            ) as key:
                raw_value, _ = winreg.QueryValueEx(
                    key,
                    "Desktop",
                )
            return Path(
                os.path.expandvars(raw_value)
            ).resolve()
        except Exception:
            pass

    return (Path.home() / "Desktop").resolve()


REVIEW_DIR = (
    get_windows_desktop()
    / "Transparency_PUF_Review"
)

BACKUP_FILE = (
    REVIEW_DIR
    / "27_PRE_EXECUTIVE_DQ_CARD_FIX_BACKUP.zip"
)

REPORT_FILE = (
    REVIEW_DIR
    / "27_EXECUTIVE_DQ_CARD_FIX_REPORT.xlsx"
)


def power_bi_desktop_running() -> bool:
    if os.name != "nt":
        return False

    result = subprocess.run(
        [
            "tasklist",
            "/FI",
            "IMAGENAME eq PBIDesktop.exe",
        ],
        capture_output=True,
        text=True,
        check=False,
        creationflags=getattr(
            subprocess,
            "CREATE_NO_WINDOW",
            0,
        ),
    )

    return "PBIDesktop.exe" in result.stdout


def read_json(path: Path):
    return json.loads(
        path.read_text(
            encoding="utf-8-sig",
        )
    )


def write_json(
    path: Path,
    payload,
) -> None:
    path.write_text(
        json.dumps(
            payload,
            indent=2,
            ensure_ascii=False,
        )
        + "\n",
        encoding="utf-8",
    )


def expr_bool(value: bool) -> dict:
    return {
        "expr": {
            "Literal": {
                "Value": "true" if value else "false"
            }
        }
    }


def expr_double(value: float | int) -> dict:
    return {
        "expr": {
            "Literal": {
                "Value": f"{value}D"
            }
        }
    }


def expr_string(value: str) -> dict:
    escaped = value.replace(
        "'",
        "''",
    )
    return {
        "expr": {
            "Literal": {
                "Value": f"'{escaped}'"
            }
        }
    }


def solid_color(color: str) -> dict:
    return {
        "solid": {
            "color": expr_string(
                color
            )
        }
    }


def find_executive_page() -> tuple[Path, dict]:
    matches = []

    for page_json in PAGES_DIR.glob(
        "*/page.json"
    ):
        payload = read_json(
            page_json
        )
        if payload.get(
            "displayName"
        ) == "01 Executive Overview":
            matches.append(
                (
                    page_json.parent,
                    payload,
                )
            )

    if len(matches) != 1:
        raise RuntimeError(
            "Expected exactly one '01 Executive Overview' page; "
            f"found {len(matches)}."
        )

    return matches[0]


def visual_path(
    page_dir: Path,
    visual_name: str,
) -> Path:
    return (
        page_dir
        / "visuals"
        / visual_name
        / "visual.json"
    )


def capture_positions(
    page_dir: Path,
) -> dict[str, dict]:
    result = {}

    for path in (
        page_dir
        / "visuals"
    ).rglob(
        "visual.json"
    ):
        payload = read_json(
            path
        )
        name = payload.get(
            "name"
        )
        if name:
            result[
                name
            ] = copy.deepcopy(
                payload.get(
                    "position",
                    {},
                )
            )

    return result


def count_bound_visuals(
    page_dir: Path,
) -> int:
    count = 0

    for path in (
        page_dir
        / "visuals"
    ).rglob(
        "visual.json"
    ):
        payload = read_json(
            path
        )
        if (
            payload.get(
                "visual",
                {},
            ).get(
                "query"
            )
        ):
            count += 1

    return count


def create_backup() -> None:
    REVIEW_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    if BACKUP_FILE.exists():
        BACKUP_FILE.unlink()

    with zipfile.ZipFile(
        BACKUP_FILE,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as zf:
        for path in sorted(
            p
            for p in REPORT_DIR.rglob("*")
            if p.is_file()
        ):
            archive_name = (
                Path(REPORT_DIR.name)
                / path.relative_to(
                    REPORT_DIR
                )
            )
            zf.write(
                path,
                str(
                    archive_name
                ),
            )


def patch_dq_block(
    page_dir: Path,
) -> list[dict]:
    card_path = visual_path(
        page_dir,
        TARGET_CARD,
    )
    band_path = visual_path(
        page_dir,
        TARGET_BAND,
    )
    anchor_path = visual_path(
        page_dir,
        NEXT_ROW_ANCHOR,
    )

    for required in (
        card_path,
        band_path,
        anchor_path,
    ):
        if not required.exists():
            raise FileNotFoundError(
                f"Required visual not found: {required}"
            )

    card = read_json(
        card_path
    )
    band = read_json(
        band_path
    )
    anchor = read_json(
        anchor_path
    )

    if (
        card.get(
            "visual",
            {},
        ).get(
            "visualType"
        )
        != "card"
    ):
        raise RuntimeError(
            f"{TARGET_CARD} is not a card visual."
        )

    old_card_pos = copy.deepcopy(
        card[
            "position"
        ]
    )
    old_band_pos = copy.deepcopy(
        band[
            "position"
        ]
    )

    band_y = float(
        band[
            "position"
        ][
            "y"
        ]
    )
    next_row_y = float(
        anchor[
            "position"
        ][
            "y"
        ]
    )

    # Keep x/width/z untouched. Only make the title strip compact and
    # give the value body enough real height below it to render.
    band[
        "position"
    ][
        "height"
    ] = BAND_HEIGHT

    body_y = (
        band_y
        + BAND_HEIGHT
        + BODY_GAP
    )

    body_height = (
        next_row_y
        - BOTTOM_CLEARANCE
        - body_y
    )

    if body_height < 38:
        raise RuntimeError(
            "Not enough vertical room to render the DQ value safely. "
            f"Computed body height: {body_height:.2f}px."
        )

    card[
        "position"
    ][
        "y"
    ] = body_y

    card[
        "position"
    ][
        "height"
    ] = body_height

    visual = card[
        "visual"
    ]

    labels = (
        visual.setdefault(
            "objects",
            {},
        ).setdefault(
            "labels",
            [
                {
                    "properties": {}
                }
            ],
        )
    )

    if not labels:
        labels.append(
            {
                "properties": {}
            }
        )

    label_props = labels[0].setdefault(
        "properties",
        {},
    )

    label_props[
        "fontSize"
    ] = expr_double(
        VALUE_FONT_SIZE
    )
    label_props[
        "bold"
    ] = expr_bool(
        True
    )
    label_props[
        "color"
    ] = solid_color(
        COLORS["amber_dark"]
    )

    category_labels = (
        visual.setdefault(
            "objects",
            {},
        ).setdefault(
            "categoryLabels",
            [
                {
                    "properties": {}
                }
            ],
        )
    )

    if not category_labels:
        category_labels.append(
            {
                "properties": {}
            }
        )

    category_labels[0].setdefault(
        "properties",
        {},
    )[
        "show"
    ] = expr_bool(
        False
    )

    container = visual.setdefault(
        "visualContainerObjects",
        {},
    )

    title = container.setdefault(
        "title",
        [
            {
                "properties": {}
            }
        ],
    )

    if not title:
        title.append(
            {
                "properties": {}
            }
        )

    # The visible title is the separate amber title-band textbox.
    title[0].setdefault(
        "properties",
        {},
    )[
        "show"
    ] = expr_bool(
        False
    )

    background = container.setdefault(
        "background",
        [
            {
                "properties": {}
            }
        ],
    )

    if not background:
        background.append(
            {
                "properties": {}
            }
        )

    background[0].setdefault(
        "properties",
        {},
    ).update(
        {
            "show": expr_bool(
                True
            ),
            "color": solid_color(
                COLORS["white"]
            ),
            "transparency": expr_double(
                0
            ),
        }
    )

    border = container.setdefault(
        "border",
        [
            {
                "properties": {}
            }
        ],
    )

    if not border:
        border.append(
            {
                "properties": {}
            }
        )

    border[0].setdefault(
        "properties",
        {},
    ).update(
        {
            "show": expr_bool(
                True
            ),
            "color": solid_color(
                COLORS["border"]
            ),
            "width": expr_double(
                1
            ),
            "radius": expr_double(
                8
            ),
        }
    )

    write_json(
        band_path,
        band,
    )
    write_json(
        card_path,
        card,
    )

    return [
        {
            "Visual": TARGET_BAND,
            "Action": "Compact title strip",
            "BeforePosition": json.dumps(
                old_band_pos,
                sort_keys=True,
            ),
            "AfterPosition": json.dumps(
                band[
                    "position"
                ],
                sort_keys=True,
            ),
        },
        {
            "Visual": TARGET_CARD,
            "Action": "Give value body enough render height below title strip",
            "BeforePosition": json.dumps(
                old_card_pos,
                sort_keys=True,
            ),
            "AfterPosition": json.dumps(
                card[
                    "position"
                ],
                sort_keys=True,
            ),
        },
    ]


def validate_target_geometry(
    page_dir: Path,
) -> dict:
    card = read_json(
        visual_path(
            page_dir,
            TARGET_CARD,
        )
    )
    band = read_json(
        visual_path(
            page_dir,
            TARGET_BAND,
        )
    )
    anchor = read_json(
        visual_path(
            page_dir,
            NEXT_ROW_ANCHOR,
        )
    )

    band_bottom = (
        float(
            band[
                "position"
            ][
                "y"
            ]
        )
        + float(
            band[
                "position"
            ][
                "height"
            ]
        )
    )

    card_y = float(
        card[
            "position"
        ][
            "y"
        ]
    )
    card_height = float(
        card[
            "position"
        ][
            "height"
        ]
    )
    card_bottom = (
        card_y
        + card_height
    )
    next_row_y = float(
        anchor[
            "position"
        ][
            "y"
        ]
    )

    passed = (
        card_y
        >= band_bottom
        + BODY_GAP
        - 0.01
        and card_height
        >= 38
        and card_bottom
        <= next_row_y
        - BOTTOM_CLEARANCE
        + 0.01
    )

    return {
        "CheckID": "DQFIX-001",
        "TestName": "DQ title/body geometry",
        "Expected": (
            "body below title; body >=38px; no overlap with KPI row"
        ),
        "Actual": (
            f"bandBottom={band_bottom:.2f}; "
            f"cardY={card_y:.2f}; "
            f"cardHeight={card_height:.2f}; "
            f"cardBottom={card_bottom:.2f}; "
            f"nextRowY={next_row_y:.2f}"
        ),
        "Status": (
            "PASS"
            if passed
            else "FAIL"
        ),
    }


def validate_non_target_geometry(
    before: dict[str, dict],
    after: dict[str, dict],
) -> tuple[dict, list[dict]]:
    allowed = {
        TARGET_CARD,
        TARGET_BAND,
    }

    changes = []

    for name in sorted(
        set(
            before
        )
        | set(
            after
        )
    ):
        if name in allowed:
            continue

        if (
            before.get(
                name
            )
            != after.get(
                name
            )
        ):
            changes.append(
                {
                    "VisualName": name,
                    "Before": json.dumps(
                        before.get(
                            name
                        ),
                        sort_keys=True,
                    ),
                    "After": json.dumps(
                        after.get(
                            name
                        ),
                        sort_keys=True,
                    ),
                }
            )

    return (
        {
            "CheckID": "DQFIX-002",
            "TestName": "All other visual geometry preserved",
            "Expected": 0,
            "Actual": len(
                changes
            ),
            "Status": (
                "PASS"
                if not changes
                else "FAIL"
            ),
        },
        changes,
    )


def validate_no_unsupported_page_roots() -> list[dict]:
    offenders = []

    for page_json in PAGES_DIR.glob(
        "*/page.json"
    ):
        payload = read_json(
            page_json
        )
        for prop in sorted(
            UNSUPPORTED_PAGE_ROOT_PROPERTIES
        ):
            if prop in payload:
                offenders.append(
                    {
                        "Page": payload.get(
                            "displayName"
                        ),
                        "Property": prop,
                    }
                )

    return offenders


def validate_no_full_page_visuals() -> list[dict]:
    offenders = []

    for visual_json in PAGES_DIR.rglob(
        "visual.json"
    ):
        payload = read_json(
            visual_json
        )
        pos = payload.get(
            "position",
            {},
        )
        width = float(
            pos.get(
                "width",
                0,
            )
            or 0
        )
        height = float(
            pos.get(
                "height",
                0,
            )
            or 0
        )

        if (
            width
            >= CANVAS_WIDTH * 0.94
            and height
            >= CANVAS_HEIGHT * 0.94
        ):
            offenders.append(
                {
                    "VisualName": payload.get(
                        "name"
                    ),
                    "Width": width,
                    "Height": height,
                    "RelativePath": str(
                        visual_json.relative_to(
                            PROJECT_ROOT
                        )
                    ),
                }
            )

    return offenders


def validate_json_tree() -> list[dict]:
    failures = []

    for path in (
        REPORT_DIR
        / "definition"
    ).rglob(
        "*.json"
    ):
        try:
            read_json(
                path
            )
        except Exception as exc:
            failures.append(
                {
                    "RelativePath": str(
                        path.relative_to(
                            PROJECT_ROOT
                        )
                    ),
                    "Error": str(
                        exc
                    ),
                }
            )

    return failures


def style_report(
    path: Path,
) -> None:
    from openpyxl import load_workbook
    from openpyxl.styles import (
        Alignment,
        Font,
        PatternFill,
    )
    from openpyxl.utils import (
        get_column_letter,
    )

    wb = load_workbook(
        path
    )

    header_fill = PatternFill(
        "solid",
        fgColor="1F4E78",
    )
    header_font = Font(
        color="FFFFFF",
        bold=True,
    )

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

        if (
            ws.max_row
            and ws.max_column
        ):
            ws.auto_filter.ref = (
                ws.dimensions
            )

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        for col_idx in range(
            1,
            ws.max_column + 1,
        ):
            max_len = 0

            for row_idx in range(
                1,
                min(
                    ws.max_row,
                    200,
                )
                + 1,
            ):
                value = ws.cell(
                    row=row_idx,
                    column=col_idx,
                ).value
                if value is not None:
                    max_len = max(
                        max_len,
                        len(
                            str(
                                value
                            )
                        ),
                    )

            ws.column_dimensions[
                get_column_letter(
                    col_idx
                )
            ].width = min(
                max(
                    max_len + 2,
                    10,
                ),
                75,
            )

        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True,
                )

    wb.save(
        path
    )


def main() -> int:
    print("=" * 78)
    print(
        "Transparency in Coverage PUF — "
        "01 Executive DQ Card Render Fix"
    )
    print("=" * 78)

    print(
        f"Script version : {SCRIPT_VERSION}"
    )
    print(
        f"Report folder  : {REPORT_DIR}"
    )
    print(
        f"Backup         : {BACKUP_FILE}"
    )
    print(
        f"Review file    : {REPORT_FILE}"
    )

    if power_bi_desktop_running():
        raise RuntimeError(
            "Power BI Desktop is running. "
            "Save and close it completely before applying this PBIR patch."
        )

    page_dir, _ = (
        find_executive_page()
    )

    visual_count = len(
        list(
            (
                page_dir
                / "visuals"
            ).rglob(
                "visual.json"
            )
        )
    )
    bound_count = (
        count_bound_visuals(
            page_dir
        )
    )

    if visual_count != EXPECTED_VISUALS:
        raise RuntimeError(
            f"Expected {EXPECTED_VISUALS} Executive visuals; "
            f"found {visual_count}. No changes made."
        )

    if bound_count != EXPECTED_BOUND_VISUALS:
        raise RuntimeError(
            f"Expected {EXPECTED_BOUND_VISUALS} bound/query visuals; "
            f"found {bound_count}. No changes made."
        )

    positions_before = (
        capture_positions(
            page_dir
        )
    )

    create_backup()

    patch_rows = (
        patch_dq_block(
            page_dir
        )
    )

    positions_after = (
        capture_positions(
            page_dir
        )
    )

    target_check = (
        validate_target_geometry(
            page_dir
        )
    )

    geometry_check, unexpected = (
        validate_non_target_geometry(
            positions_before,
            positions_after,
        )
    )

    unsupported = (
        validate_no_unsupported_page_roots()
    )
    full_page = (
        validate_no_full_page_visuals()
    )
    json_failures = (
        validate_json_tree()
    )

    validation_rows = [
        target_check,
        geometry_check,
        {
            "CheckID": "DQFIX-003",
            "TestName": "Visual count preserved",
            "Expected": EXPECTED_VISUALS,
            "Actual": len(
                list(
                    (
                        page_dir
                        / "visuals"
                    ).rglob(
                        "visual.json"
                    )
                )
            ),
            "Status": "PASS",
        },
        {
            "CheckID": "DQFIX-004",
            "TestName": "Bound/query visual count preserved",
            "Expected": EXPECTED_BOUND_VISUALS,
            "Actual": count_bound_visuals(
                page_dir
            ),
            "Status": (
                "PASS"
                if count_bound_visuals(
                    page_dir
                )
                == EXPECTED_BOUND_VISUALS
                else "FAIL"
            ),
        },
        {
            "CheckID": "DQFIX-005",
            "TestName": "Unsupported page-root properties",
            "Expected": 0,
            "Actual": len(
                unsupported
            ),
            "Status": (
                "PASS"
                if not unsupported
                else "FAIL"
            ),
        },
        {
            "CheckID": "DQFIX-006",
            "TestName": "Full-page selectable visuals",
            "Expected": 0,
            "Actual": len(
                full_page
            ),
            "Status": (
                "PASS"
                if not full_page
                else "FAIL"
            ),
        },
        {
            "CheckID": "DQFIX-007",
            "TestName": "PBIR JSON parse failures",
            "Expected": 0,
            "Actual": len(
                json_failures
            ),
            "Status": (
                "PASS"
                if not json_failures
                else "FAIL"
            ),
        },
    ]

    failures = sum(
        row[
            "Status"
        ] == "FAIL"
        for row
        in validation_rows
    )

    summary_df = pd.DataFrame(
        [
            (
                "PatchStatus",
                "PASS"
                if failures == 0
                else "FAIL",
            ),
            (
                "ScriptVersion",
                SCRIPT_VERSION,
            ),
            (
                "PatchedAtLocal",
                datetime.now()
                .isoformat(
                    timespec="seconds"
                ),
            ),
            (
                "Page",
                "01 Executive Overview",
            ),
            (
                "Target",
                (
                    "Global Open DQ Exceptions title band + value card"
                ),
            ),
            (
                "BandHeight",
                BAND_HEIGHT,
            ),
            (
                "BodyGap",
                BODY_GAP,
            ),
            (
                "ValueFontSize",
                VALUE_FONT_SIZE,
            ),
            (
                "OtherVisualGeometryChanges",
                len(
                    unexpected
                ),
            ),
            (
                "ValidationFailures",
                failures,
            ),
            (
                "BackupFile",
                str(
                    BACKUP_FILE
                ),
            ),
        ],
        columns=[
            "Item",
            "Value",
        ],
    )

    REVIEW_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    with pd.ExcelWriter(
        REPORT_FILE,
        engine="openpyxl",
    ) as writer:
        summary_df.to_excel(
            writer,
            sheet_name="00_Summary",
            index=False,
        )

        pd.DataFrame(
            patch_rows
        ).to_excel(
            writer,
            sheet_name="01_Patches",
            index=False,
        )

        pd.DataFrame(
            validation_rows
        ).to_excel(
            writer,
            sheet_name="02_Validation",
            index=False,
        )

        if unexpected:
            pd.DataFrame(
                unexpected
            ).to_excel(
                writer,
                sheet_name="03_Unexpected_Geometry",
                index=False,
            )
        else:
            pd.DataFrame(
                [
                    {
                        "Info": (
                            "Only the DQ title band and DQ value-card geometry changed."
                        )
                    }
                ]
            ).to_excel(
                writer,
                sheet_name="03_Unexpected_Geometry",
                index=False,
            )

    style_report(
        REPORT_FILE
    )

    print()
    print(
        "01 Executive DQ card render fix completed."
    )
    print(
        f"Patch status                  : "
        f"{'PASS' if failures == 0 else 'FAIL'}"
    )
    print(
        "Target                        : Global Open DQ Exceptions"
    )
    print(
        f"Title band height             : {BAND_HEIGHT:.0f}px"
    )
    print(
        f"Value font size               : {VALUE_FONT_SIZE:.0f}pt"
    )
    print(
        f"Other visual geometry changes : {len(unexpected)}"
    )
    print(
        f"Visuals preserved             : {EXPECTED_VISUALS}"
    )
    print(
        f"Bound/query visuals preserved : {EXPECTED_BOUND_VISUALS}"
    )
    print(
        f"Full-page selectable visuals  : {len(full_page)}"
    )
    print(
        f"Unsupported page properties   : {len(unsupported)}"
    )
    print(
        f"Validation failures           : {failures}"
    )
    print(
        f"Backup                        : {BACKUP_FILE}"
    )
    print(
        f"Review evidence               : {REPORT_FILE}"
    )

    if failures:
        print()
        print(
            "Validation failed. Do not open/save the PBIP until reviewed."
        )
        return 2

    print()
    print(
        "Next: open TransparencyInCoverage.pbip and inspect only the "
        "Global Open DQ Exceptions block. The value 36 should render below "
        "the amber title strip, while every other visual stays exactly in place."
    )

    return 0


if __name__ == "__main__":
    try:
        sys.exit(
            main()
        )
    except Exception as exc:
        print()
        print(
            "EXECUTIVE DQ CARD RENDER FIX FAILED"
        )
        print(
            str(
                exc
            )
        )
        sys.exit(
            1
        )
