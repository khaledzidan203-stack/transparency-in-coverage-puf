from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import importlib.util
import json
import os
import re
import sys
import tempfile
from typing import Any


SCRIPT_VERSION = "1.0.1"

PROJECT_ROOT = Path(__file__).resolve().parents[2]
POWERBI_ROOT = PROJECT_ROOT / "powerbi"
REPORT_DIR = POWERBI_ROOT / "TransparencyInCoverage.Report"
PAGES_DIR = REPORT_DIR / "definition" / "pages"

SOURCE_SCRIPTS = {
    "08 Issuer & State Explorer": (
        PROJECT_ROOT
        / "src"
        / "validation"
        / "38_build_issuer_state_explorer.py"
    ),
    "09 Data Availability": (
        PROJECT_ROOT
        / "src"
        / "validation"
        / "39_build_data_availability.py"
    ),
}

BUILD_FUNCTIONS = {
    "08 Issuer & State Explorer": "build_issuer_state_explorer",
    "09 Data Availability": "build_data_availability_page",
}


# =====================================================================
# Output location
# =====================================================================

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

MARKDOWN_FILE = (
    REVIEW_DIR
    / "40_PAGES_08_09_MANUAL_VISUAL_BASELINE.md"
)

JSON_FILE = (
    REVIEW_DIR
    / "40_PAGES_08_09_MANUAL_VISUAL_BASELINE.json"
)


# =====================================================================
# Basic IO
# =====================================================================

def read_json(path: Path) -> dict:
    return json.loads(
        path.read_text(
            encoding="utf-8-sig",
        )
    )


def canonical_json(value: Any) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(
            ",",
            ":",
        ),
    )


def find_pages() -> dict[
    str,
    tuple[Path, dict]
]:
    pages = {}

    for page_json in PAGES_DIR.glob(
        "*/page.json"
    ):
        payload = read_json(
            page_json
        )

        display_name = payload.get(
            "displayName"
        )

        if display_name:
            pages[
                display_name
            ] = (
                page_json.parent,
                payload,
            )

    return pages


def load_visuals(
    page_dir: Path,
) -> dict[str, dict]:
    visuals = {}

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
            visuals[
                name
            ] = payload

    return visuals


# =====================================================================
# PBIR decoding helpers
# =====================================================================

def decode_literal(
    value: Any,
) -> Any:
    if not isinstance(
        value,
        str,
    ):
        return value

    if (
        len(value) >= 2
        and value.startswith("'")
        and value.endswith("'")
    ):
        return (
            value[1:-1]
            .replace(
                "''",
                "'",
            )
        )

    if value == "true":
        return True

    if value == "false":
        return False

    if value.endswith(
        "D"
    ) or value.endswith(
        "L"
    ):
        raw = value[:-1]

        try:
            if "." in raw:
                return float(
                    raw
                )

            return int(
                raw
            )
        except Exception:
            return value

    return value


def expr_value(
    obj: Any,
) -> Any:
    if not isinstance(
        obj,
        dict,
    ):
        return None

    literal = (
        obj.get(
            "expr",
            {},
        ).get(
            "Literal",
            {},
        )
    )

    if (
        isinstance(
            literal,
            dict,
        )
        and "Value" in literal
    ):
        return decode_literal(
            literal["Value"]
        )

    return None


def color_value(
    obj: Any,
) -> str | None:
    if not isinstance(
        obj,
        dict,
    ):
        return None

    value = (
        obj.get(
            "solid",
            {},
        ).get(
            "color"
        )
    )

    decoded = expr_value(
        value
    )

    if isinstance(
        decoded,
        str,
    ):
        return decoded

    return None


def textbox_text(
    payload: dict,
) -> str:
    visual = payload.get(
        "visual",
        {},
    )

    if visual.get(
        "visualType"
    ) != "textbox":
        return ""

    paragraphs = (
        visual.get(
            "objects",
            {},
        )
        .get(
            "general",
            [{}],
        )[0]
        .get(
            "properties",
            {},
        )
        .get(
            "paragraphs",
            [],
        )
    )

    output = []

    for para in paragraphs:
        runs = para.get(
            "textRuns",
            [],
        )

        text = "".join(
            str(
                run.get(
                    "value",
                    "",
                )
            )
            for run in runs
        )

        if text.strip():
            output.append(
                text.strip()
            )

    return " | ".join(
        output
    )


def collect_query_refs(
    obj: Any,
) -> list[str]:
    refs: set[str] = set()

    def walk(
        node: Any,
    ) -> None:
        if isinstance(
            node,
            dict,
        ):
            query_ref = node.get(
                "queryRef"
            )

            if isinstance(
                query_ref,
                str,
            ):
                refs.add(
                    query_ref
                )

            for value in (
                node.values()
            ):
                walk(
                    value
                )

        elif isinstance(
            node,
            list,
        ):
            for value in node:
                walk(
                    value
                )

    walk(
        obj
    )

    return sorted(
        refs
    )


def extract_container_style(
    payload: dict,
) -> dict:
    visual = payload.get(
        "visual",
        {},
    )

    container = visual.get(
        "visualContainerObjects",
        {},
    )

    result: dict[str, Any] = {
        "background_show": None,
        "background_color": None,
        "border_show": None,
        "border_color": None,
        "border_width": None,
        "border_radius": None,
        "title_show": None,
        "title_text": None,
    }

    background = (
        container.get(
            "background",
            [],
        )
    )

    if background:
        props = background[0].get(
            "properties",
            {},
        )

        result[
            "background_show"
        ] = expr_value(
            props.get(
                "show"
            )
        )

        result[
            "background_color"
        ] = color_value(
            props.get(
                "color"
            )
        )

    border = container.get(
        "border",
        [],
    )

    if border:
        props = border[0].get(
            "properties",
            {},
        )

        result[
            "border_show"
        ] = expr_value(
            props.get(
                "show"
            )
        )

        result[
            "border_color"
        ] = color_value(
            props.get(
                "color"
            )
        )

        result[
            "border_width"
        ] = expr_value(
            props.get(
                "width"
            )
        )

        result[
            "border_radius"
        ] = expr_value(
            props.get(
                "radius"
            )
        )

    title = container.get(
        "title",
        [],
    )

    if title:
        props = title[0].get(
            "properties",
            {},
        )

        result[
            "title_show"
        ] = expr_value(
            props.get(
                "show"
            )
        )

        result[
            "title_text"
        ] = expr_value(
            props.get(
                "text"
            )
        )

    return result


def visual_snapshot(
    payload: dict,
) -> dict:
    position = payload.get(
        "position",
        {},
    )

    visual = payload.get(
        "visual",
        {},
    )

    query = visual.get(
        "query"
    )

    return {
        "name": payload.get(
            "name"
        ),
        "type": visual.get(
            "visualType"
        ),
        "x": position.get(
            "x"
        ),
        "y": position.get(
            "y"
        ),
        "width": position.get(
            "width"
        ),
        "height": position.get(
            "height"
        ),
        "z": position.get(
            "z"
        ),
        "tabOrder": position.get(
            "tabOrder"
        ),
        "query_bound": bool(
            query
        ),
        "query_refs": (
            collect_query_refs(
                query
            )
            if query
            else []
        ),
        "text": textbox_text(
            payload
        ),
        "container_style": (
            extract_container_style(
                payload
            )
        ),
    }


# =====================================================================
# Builder reconstruction
# =====================================================================

def import_module_from_path(
    path: Path,
    module_name: str,
):
    spec = (
        importlib.util
        .spec_from_file_location(
            module_name,
            path,
        )
    )

    if (
        spec is None
        or spec.loader is None
    ):
        raise RuntimeError(
            f"Could not import builder: {path}"
        )

    module = (
        importlib.util
        .module_from_spec(
            spec
        )
    )

    spec.loader.exec_module(
        module
    )

    return module


def reconstruct_builder_visuals(
    *,
    page_name: str,
    builder_path: Path,
    build_function_name: str,
    index_page_internal_name: str,
) -> dict[str, dict]:
    module = import_module_from_path(
        builder_path,
        "manual_baseline_"
        + re.sub(
            r"[^A-Za-z0-9]+",
            "_",
            page_name,
        ),
    )

    build_function = getattr(
        module,
        build_function_name,
        None,
    )

    if build_function is None:
        raise RuntimeError(
            f"Function '{build_function_name}' not found in "
            f"{builder_path.name}."
        )

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_page = Path(
            temp_dir
        ) / "page"

        temp_page.mkdir(
            parents=True,
            exist_ok=True,
        )

        build_function(
            temp_page,
            index_page_internal_name,
        )

        return load_visuals(
            temp_page
        )


# =====================================================================
# Deep diff
# =====================================================================

IGNORED_STYLE_PATH_PREFIXES = {
    "$schema",
    "name",
    "position",
    "visual.query",
    "visual.drillFilterOtherVisuals",
}


def flatten_json(
    obj: Any,
    prefix: str = "",
) -> dict[str, Any]:
    result: dict[
        str,
        Any,
    ] = {}

    if isinstance(
        obj,
        dict,
    ):
        for key in sorted(
            obj
        ):
            path = (
                f"{prefix}.{key}"
                if prefix
                else key
            )

            result.update(
                flatten_json(
                    obj[
                        key
                    ],
                    path,
                )
            )

    elif isinstance(
        obj,
        list,
    ):
        for idx, value in enumerate(
            obj
        ):
            path = (
                f"{prefix}[{idx}]"
            )

            result.update(
                flatten_json(
                    value,
                    path,
                )
            )

    else:
        result[
            prefix
        ] = obj

    return result


def path_is_style(
    path: str,
) -> bool:
    if path.startswith(
        "visual.objects"
    ):
        return True

    if path.startswith(
        "visual.visualContainerObjects"
    ):
        return True

    return False


def diff_payloads(
    current: dict,
    baseline: dict,
) -> list[dict]:
    left = flatten_json(
        baseline
    )
    right = flatten_json(
        current
    )

    diffs = []

    for path in sorted(
        set(
            left
        )
        | set(
            right
        )
    ):
        before = left.get(
            path,
            "<MISSING>",
        )

        after = right.get(
            path,
            "<MISSING>",
        )

        if before == after:
            continue

        if path.startswith(
            "$schema"
        ):
            continue

        diffs.append(
            {
                "path": path,
                "builder": before,
                "current": after,
                "category": (
                    "geometry"
                    if path.startswith(
                        "position."
                    )
                    else (
                        "query"
                        if path.startswith(
                            "visual.query"
                        )
                        else (
                            "filter"
                            if path.startswith(
                                "filterConfig"
                            )
                            else (
                                "style"
                                if path_is_style(
                                    path
                                )
                                else "other"
                            )
                        )
                    )
                ),
            }
        )

    return diffs


# =====================================================================
# Manual-change analysis
# =====================================================================

def compare_page(
    *,
    page_name: str,
    current_visuals: dict[str, dict],
    builder_visuals: dict[str, dict],
) -> dict:
    current_names = set(
        current_visuals
    )
    builder_names = set(
        builder_visuals
    )

    added = sorted(
        current_names
        - builder_names
    )

    removed = sorted(
        builder_names
        - current_names
    )

    shared = sorted(
        current_names
        & builder_names
    )

    changed_visuals = []

    for name in shared:
        current = current_visuals[
            name
        ]
        baseline = builder_visuals[
            name
        ]

        diffs = diff_payloads(
            current,
            baseline,
        )

        if not diffs:
            continue

        categories = Counter(
            row[
                "category"
            ]
            for row in diffs
        )

        border_diffs = [
            row
            for row in diffs
            if (
                ".border" in row[
                    "path"
                ].lower()
                or "radius" in row[
                    "path"
                ].lower()
            )
        ]

        changed_visuals.append(
            {
                "name": name,
                "current_snapshot": (
                    visual_snapshot(
                        current
                    )
                ),
                "builder_snapshot": (
                    visual_snapshot(
                        baseline
                    )
                ),
                "change_counts": dict(
                    categories
                ),
                "border_change_count": len(
                    border_diffs
                ),
                "diffs": diffs,
            }
        )

    return {
        "page": page_name,
        "added_visuals": added,
        "removed_visuals": removed,
        "changed_visuals": changed_visuals,
        "changed_visual_count": len(
            changed_visuals
        ),
    }


def pair_title_bands(
    visuals: dict[str, dict],
) -> list[dict]:
    pairs = []

    for name, payload in (
        visuals.items()
    ):
        if "_band_" not in name:
            continue

        body_name = name.replace(
            "_band_",
            "_",
            1,
        )

        if body_name not in visuals:
            continue

        band = visual_snapshot(
            payload
        )

        body = visual_snapshot(
            visuals[
                body_name
            ]
        )

        try:
            band_bottom = (
                float(
                    band[
                        "y"
                    ]
                )
                + float(
                    band[
                        "height"
                    ]
                )
            )

            gap = (
                float(
                    body[
                        "y"
                    ]
                )
                - band_bottom
            )
        except Exception:
            gap = None

        pairs.append(
            {
                "band": name,
                "body": body_name,
                "band_geometry": (
                    f"x={band['x']}, y={band['y']}, "
                    f"w={band['width']}, h={band['height']}"
                ),
                "body_geometry": (
                    f"x={body['x']}, y={body['y']}, "
                    f"w={body['width']}, h={body['height']}"
                ),
                "vertical_gap": gap,
                "same_x": (
                    band[
                        "x"
                    ]
                    == body[
                        "x"
                    ]
                ),
                "same_width": (
                    band[
                        "width"
                    ]
                    == body[
                        "width"
                    ]
                ),
            }
        )

    return sorted(
        pairs,
        key=lambda row: (
            row[
                "band"
            ]
        ),
    )


# =====================================================================
# Markdown
# =====================================================================

def md_escape(
    value: Any,
) -> str:
    text = str(
        value
    )

    return (
        text.replace(
            "|",
            r"\|",
        )
        .replace(
            "\n",
            " ",
        )
    )


def geometry_text(
    snapshot: dict,
) -> str:
    return (
        f"({snapshot['x']}, {snapshot['y']}) "
        f"{snapshot['width']}×{snapshot['height']}"
    )


def format_diff_value(
    value: Any,
) -> str:
    if value == "<MISSING>":
        return value

    text = str(
        decode_literal(
            value
        )
    )

    if len(
        text
    ) > 180:
        text = (
            text[:177]
            + "..."
        )

    return text


def build_markdown(
    *,
    page_results: list[dict],
    current_pages: dict[
        str,
        tuple[Path, dict]
    ],
    current_visual_maps: dict[
        str,
        dict[str, dict]
    ],
) -> str:
    lines: list[str] = []

    lines.append(
        "# Manual Visual Baseline — Pages 08 & 09"
    )
    lines.append(
        ""
    )
    lines.append(
        f"- Generated: {datetime.now().isoformat(timespec='seconds')}"
    )
    lines.append(
        f"- Script version: {SCRIPT_VERSION}"
    )
    lines.append(
        "- Purpose: capture the **current saved PBIR state after manual edits** "
        "and compare it with the original page-builder scripts."
    )
    lines.append(
        "- Source of truth for future page engineering: **Current saved PBIR wins** "
        "when it differs from the builder."
    )
    lines.append(
        ""
    )

    total_changed = sum(
        item[
            "changed_visual_count"
        ]
        for item in page_results
    )

    lines.append(
        "## Executive Summary"
    )
    lines.append(
        ""
    )
    lines.append(
        f"- Pages documented: {len(page_results)}"
    )
    lines.append(
        f"- Visuals with detected manual differences: {total_changed}"
    )
    lines.append(
        "- Query/binding differences are called out separately from visual/layout changes."
    )
    lines.append(
        "- Border/radius changes are explicitly identified because these have been "
        "a recurring manual refinement."
    )
    lines.append(
        ""
    )

    for result in page_results:
        page_name = result[
            "page"
        ]

        page_dir, page_payload = (
            current_pages[
                page_name
            ]
        )

        visuals = (
            current_visual_maps[
                page_name
            ]
        )

        snapshots = [
            visual_snapshot(
                payload
            )
            for payload in visuals.values()
        ]

        snapshots.sort(
            key=lambda row: (
                float(
                    row[
                        "y"
                    ]
                    or 0
                ),
                float(
                    row[
                        "x"
                    ]
                    or 0
                ),
                row[
                    "name"
                ],
            )
        )

        lines.append(
            f"## {page_name}"
        )
        lines.append(
            ""
        )
        lines.append(
            f"- Page internal name: `{page_payload.get('name')}`"
        )
        lines.append(
            f"- Canvas: {page_payload.get('width')} × {page_payload.get('height')}"
        )
        lines.append(
            f"- Current visual count: {len(visuals)}"
        )
        lines.append(
            f"- Manual-difference visual count: {result['changed_visual_count']}"
        )
        lines.append(
            f"- Added visuals vs builder: {len(result['added_visuals'])}"
        )
        lines.append(
            f"- Removed visuals vs builder: {len(result['removed_visuals'])}"
        )
        lines.append(
            ""
        )

        lines.append(
            "### Current Saved Visual Geometry"
        )
        lines.append(
            ""
        )
        lines.append(
            "| Visual | Type | Geometry `(x,y) w×h` | Bound | Current text/title | Background | Border |"
        )
        lines.append(
            "|---|---|---:|:---:|---|---|---|"
        )

        for snap in snapshots:
            style = snap[
                "container_style"
            ]

            border = (
                f"show={style['border_show']}; "
                f"color={style['border_color']}; "
                f"width={style['border_width']}; "
                f"radius={style['border_radius']}"
            )

            lines.append(
                "| "
                + " | ".join(
                    [
                        f"`{md_escape(snap['name'])}`",
                        md_escape(
                            snap[
                                "type"
                            ]
                        ),
                        md_escape(
                            geometry_text(
                                snap
                            )
                        ),
                        (
                            "Yes"
                            if snap[
                                "query_bound"
                            ]
                            else "No"
                        ),
                        md_escape(
                            snap[
                                "text"
                            ]
                            or style[
                                "title_text"
                            ]
                            or ""
                        ),
                        md_escape(
                            style[
                                "background_color"
                            ]
                        ),
                        md_escape(
                            border
                        ),
                    ]
                )
                + " |"
            )

        lines.append(
            ""
        )

        lines.append(
            "### Title-Band / Body Alignment"
        )
        lines.append(
            ""
        )

        pairs = pair_title_bands(
            visuals
        )

        if pairs:
            lines.append(
                "| Band | Body | Band geometry | Body geometry | Vertical gap | Same X | Same width |"
            )
            lines.append(
                "|---|---|---|---|---:|:---:|:---:|"
            )

            for pair in pairs:
                lines.append(
                    "| "
                    + " | ".join(
                        [
                            f"`{pair['band']}`",
                            f"`{pair['body']}`",
                            md_escape(
                                pair[
                                    "band_geometry"
                                ]
                            ),
                            md_escape(
                                pair[
                                    "body_geometry"
                                ]
                            ),
                            md_escape(
                                pair[
                                    "vertical_gap"
                                ]
                            ),
                            (
                                "Yes"
                                if pair[
                                    "same_x"
                                ]
                                else "No"
                            ),
                            (
                                "Yes"
                                if pair[
                                    "same_width"
                                ]
                                else "No"
                            ),
                        ]
                    )
                    + " |"
                )
        else:
            lines.append(
                "_No conventional band/body pairs detected._"
            )

        lines.append(
            ""
        )

        lines.append(
            "### Manual Changes Detected vs Builder"
        )
        lines.append(
            ""
        )

        if (
            not result[
                "changed_visuals"
            ]
            and not result[
                "added_visuals"
            ]
            and not result[
                "removed_visuals"
            ]
        ):
            lines.append(
                "_No difference detected between current PBIR and builder output._"
            )
            lines.append(
                ""
            )

        if result[
            "added_visuals"
        ]:
            lines.append(
                "**Added manually:** "
                + ", ".join(
                    f"`{name}`"
                    for name in result[
                        "added_visuals"
                    ]
                )
            )
            lines.append(
                ""
            )

        if result[
            "removed_visuals"
        ]:
            lines.append(
                "**Removed manually:** "
                + ", ".join(
                    f"`{name}`"
                    for name in result[
                        "removed_visuals"
                    ]
                )
            )
            lines.append(
                ""
            )

        for changed in result[
            "changed_visuals"
        ]:
            name = changed[
                "name"
            ]

            current = changed[
                "current_snapshot"
            ]

            builder = changed[
                "builder_snapshot"
            ]

            lines.append(
                f"#### `{name}`"
            )
            lines.append(
                ""
            )
            lines.append(
                f"- Builder geometry: `{geometry_text(builder)}`"
            )
            lines.append(
                f"- Current geometry: `{geometry_text(current)}`"
            )
            lines.append(
                f"- Change categories: `{changed['change_counts']}`"
            )
            lines.append(
                f"- Border/radius differences: `{changed['border_change_count']}`"
            )

            geometry_diffs = [
                row
                for row in changed[
                    "diffs"
                ]
                if row[
                    "category"
                ] == "geometry"
            ]

            style_diffs = [
                row
                for row in changed[
                    "diffs"
                ]
                if row[
                    "category"
                ] == "style"
            ]

            query_diffs = [
                row
                for row in changed[
                    "diffs"
                ]
                if row[
                    "category"
                ] in {
                    "query",
                    "filter",
                }
            ]

            other_diffs = [
                row
                for row in changed[
                    "diffs"
                ]
                if row[
                    "category"
                ] == "other"
            ]

            if geometry_diffs:
                lines.append(
                    "- Geometry changes:"
                )

                for row in geometry_diffs:
                    lines.append(
                        "  - `"
                        + row[
                            "path"
                        ]
                        + "`: "
                        + f"`{format_diff_value(row['builder'])}` → "
                        + f"`{format_diff_value(row['current'])}`"
                    )

            if style_diffs:
                lines.append(
                    "- Style/container changes:"
                )

                for row in style_diffs[:30]:
                    lines.append(
                        "  - `"
                        + row[
                            "path"
                        ]
                        + "`: "
                        + f"`{format_diff_value(row['builder'])}` → "
                        + f"`{format_diff_value(row['current'])}`"
                    )

                if len(
                    style_diffs
                ) > 30:
                    lines.append(
                        f"  - ... {len(style_diffs) - 30} additional style differences "
                        "are preserved in the JSON companion file."
                    )

            if query_diffs:
                lines.append(
                    "- **Query/filter changes — review carefully:**"
                )

                for row in query_diffs[:20]:
                    lines.append(
                        "  - `"
                        + row[
                            "path"
                        ]
                        + "`: "
                        + f"`{format_diff_value(row['builder'])}` → "
                        + f"`{format_diff_value(row['current'])}`"
                    )

            if other_diffs:
                lines.append(
                    "- Other changes:"
                )

                for row in other_diffs[:15]:
                    lines.append(
                        "  - `"
                        + row[
                            "path"
                        ]
                        + "`: "
                        + f"`{format_diff_value(row['builder'])}` → "
                        + f"`{format_diff_value(row['current'])}`"
                    )

            lines.append(
                ""
            )

        lines.append(
            "### Current Query Bindings"
        )
        lines.append(
            ""
        )
        lines.append(
            "| Visual | Query references |"
        )
        lines.append(
            "|---|---|"
        )

        for snap in snapshots:
            if not snap[
                "query_bound"
            ]:
                continue

            lines.append(
                f"| `{snap['name']}` | "
                + md_escape(
                    ", ".join(
                        snap[
                            "query_refs"
                        ]
                    )
                )
                + " |"
            )

        lines.append(
            ""
        )

    lines.append(
        "## Rule for Future Pages"
    )
    lines.append(
        ""
    )
    lines.append(
        "When a builder-script layout conflicts with the saved PBIR documented here, "
        "**preserve the saved PBIR geometry and styling**. Manual border, radius, spacing, "
        "card-height, slicer, and section-layout adjustments are intentional unless explicitly "
        "changed later."
    )
    lines.append(
        ""
    )
    lines.append(
        "Use the companion JSON file when exact nested PBIR properties are required."
    )
    lines.append(
        ""
    )

    return "\n".join(
        lines
    )


# =====================================================================
# Main
# =====================================================================

def main() -> int:
    print(
        "=" * 78
    )
    print(
        "Transparency in Coverage PUF — "
        "Pages 08/09 Manual Visual Baseline Documentation"
    )
    print(
        "=" * 78
    )
    print(
        f"Script version : {SCRIPT_VERSION}"
    )
    print(
        f"Report folder  : {REPORT_DIR}"
    )
    print(
        f"Markdown       : {MARKDOWN_FILE}"
    )
    print(
        f"JSON           : {JSON_FILE}"
    )
    print()

    if not REPORT_DIR.exists():
        raise FileNotFoundError(
            f"Report folder not found: {REPORT_DIR}"
        )

    for page_name, path in (
        SOURCE_SCRIPTS.items()
    ):
        if not path.exists():
            raise FileNotFoundError(
                f"Builder script for '{page_name}' not found: {path}"
            )

    pages = find_pages()

    required_pages = {
        "00 INDEX",
        "08 Issuer & State Explorer",
        "09 Data Availability",
    }

    missing_pages = (
        required_pages
        - set(
            pages
        )
    )

    if missing_pages:
        raise RuntimeError(
            "Required report pages are missing: "
            + ", ".join(
                sorted(
                    missing_pages
                )
            )
        )

    index_internal_name = (
        pages[
            "00 INDEX"
        ][1].get(
            "name"
        )
    )

    if not index_internal_name:
        raise RuntimeError(
            "Could not resolve INDEX internal page name."
        )

    current_visual_maps: dict[
        str,
        dict[str, dict]
    ] = {}

    builder_visual_maps: dict[
        str,
        dict[str, dict]
    ] = {}

    page_results = []

    for page_name in (
        "08 Issuer & State Explorer",
        "09 Data Availability",
    ):
        page_dir, _ = (
            pages[
                page_name
            ]
        )

        current_visuals = (
            load_visuals(
                page_dir
            )
        )

        current_visual_maps[
            page_name
        ] = current_visuals

        builder_visuals = (
            reconstruct_builder_visuals(
                page_name=page_name,
                builder_path=(
                    SOURCE_SCRIPTS[
                        page_name
                    ]
                ),
                build_function_name=(
                    BUILD_FUNCTIONS[
                        page_name
                    ]
                ),
                index_page_internal_name=(
                    index_internal_name
                ),
            )
        )

        builder_visual_maps[
            page_name
        ] = builder_visuals

        page_results.append(
            compare_page(
                page_name=page_name,
                current_visuals=current_visuals,
                builder_visuals=builder_visuals,
            )
        )

    payload = {
        "generated_at_local": (
            datetime.now()
            .isoformat(
                timespec="seconds"
            )
        ),
        "script_version": (
            SCRIPT_VERSION
        ),
        "principle": (
            "Current saved PBIR is authoritative for manual visual/layout refinements."
        ),
        "pages": {},
        "comparisons": page_results,
    }

    for page_name in (
        "08 Issuer & State Explorer",
        "09 Data Availability",
    ):
        page_dir, page_json = (
            pages[
                page_name
            ]
        )

        current_visuals = (
            current_visual_maps[
                page_name
            ]
        )

        payload[
            "pages"
        ][
            page_name
        ] = {
            "page_json": page_json,
            "visuals": {
                name: {
                    "snapshot": visual_snapshot(
                        visual
                    ),
                    "raw": visual,
                }
                for name, visual in sorted(
                    current_visuals.items()
                )
            },
            "title_body_pairs": pair_title_bands(
                current_visuals
            ),
        }

    markdown = build_markdown(
        page_results=page_results,
        current_pages=pages,
        current_visual_maps=(
            current_visual_maps
        ),
    )

    REVIEW_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    MARKDOWN_FILE.write_text(
        markdown,
        encoding="utf-8",
    )

    JSON_FILE.write_text(
        json.dumps(
            payload,
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )

    changed_count = sum(
        item[
            "changed_visual_count"
        ]
        for item in page_results
    )

    border_change_count = sum(
        changed[
            "border_change_count"
        ]
        for page in page_results
        for changed in page[
            "changed_visuals"
        ]
    )

    query_change_count = sum(
        1
        for page in page_results
        for changed in page[
            "changed_visuals"
        ]
        for row in changed[
            "diffs"
        ]
        if row[
            "category"
        ] in {
            "query",
            "filter",
        }
    )

    print(
        "Documentation completed."
    )
    print(
        f"Pages documented             : 2"
    )
    print(
        f"Visuals with manual changes  : {changed_count}"
    )
    print(
        f"Border/radius differences    : {border_change_count}"
    )
    print(
        f"Query/filter differences     : {query_change_count}"
    )
    print(
        f"Markdown                     : {MARKDOWN_FILE}"
    )
    print(
        f"JSON                         : {JSON_FILE}"
    )
    print()
    print(
        "This script is READ ONLY for PBIR: it does not change report or semantic-model files."
    )
    print(
        "Upload the Markdown file to ChatGPT. Keep the JSON file locally unless exact nested "
        "PBIR properties need to be inspected."
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
            "MANUAL VISUAL BASELINE DOCUMENTATION FAILED"
        )
        print(
            str(
                exc
            )
        )
        sys.exit(
            1
        )
