# Repository inventory and script catalog

Release audit: 2026-09-26. Current saved PBIR is authoritative.

No files were moved or deleted. Historical scripts remain in place to preserve path dependencies (notably script 40 imports builders 38 and 39). The catalog provides the conceptual archive; no historical builder is a normal reproduction step. Scripts that patch DAX or relationships are also historical, not validators. Script 40 reconstructs builder output in a temporary directory and is not part of the safe release checks.

## Script roles

| File | Type | State | Decision | Reason |
|---|---|---|---|---|
| src/analysis/04_run_eda.py | analysis | current | KEEP | Release source |
| src/discovery/01_workbook_inventory.py | discovery | current | KEEP | Release source |
| src/transformation/02_build_canonical_dataset.py | data build | current | KEEP | Release source |
| src/validation/03_validate_kpis.py | read-only KPI audit | current | KEEP | Release source |
| src/validation/05_export_powerbi_model_state.py | read-only model export | current | KEEP | Release source |
| src/validation/06_repair_powerbi_relationships.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/07_validate_powerbi_semantic_model.py | read-only semantic audit | current | KEEP | Release source |
| src/validation/08_install_governed_dax_measures.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/09_runtime_dax_reconciliation.py | read-only runtime DAX audit | current | KEEP | Release source |
| src/validation/10_patch_dax_runtime_failures.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/11_audit_powerbi_report_structure.py | read-only PBIR audit | current | KEEP | Release source |
| src/validation/12_build_report_shell_and_index.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/13_center_index_tile_text.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/14_build_executive_overview.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/16_rebuild_executive_overview_v2.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/18_rebuild_executive_overview_v3.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/19_build_network_comparison.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/20_polish_network_comparison.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/21_fix_network_comparison_rendering.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/22_polish_network_comparison_final.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/23_rebuild_network_comparison_storytelling.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/24_finalize_network_storytelling_style.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/25_rebuild_executive_overview_storytelling.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/26_patch_executive_card_body_alignment.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/27_fix_executive_dq_card.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/28_build_denials_storytelling.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/29_patch_denials_lower_zone.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/30_fix_denials_dq_card.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/31_relayout_denials_bottom_zone.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/32_build_denial_reasons_storytelling.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/33_build_appeals_storytelling.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/34_fix_appeals_validation_and_state_chart.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/35_build_resubmissions_storytelling.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/36_build_enrollment_storytelling.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/37_patch_enrollment_plan_type_visibility.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/38_build_issuer_state_explorer.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/39_build_data_availability.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/40_document_manual_layout_pages_08_09.py | manual-layout evidence audit | current | KEEP | Release source |
| src/validation/41_build_data_quality_methodology.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/42_rebuild_index_storytelling.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/43_patch_navigation_hover_tooltips.py | mutating report/model engineering | historical | KEEP | Retain in place: numbered scripts and baseline audit contain path dependencies; never run as release pipeline |
| src/validation/44_fix_dq_scope_filter_context.py | targeted semantic-model patch (MUTATES DAX) | historical, already applied | KEEP | Wraps two Scope predicates in KEEPFILTERS; local backup/evidence writes; reviewed only, do not rerun or treat as a validator |
| src/validation/release_audit.py | read-only publication audit | current | KEEP | Compiles scripts and validates publication files; writes ignored evidence only |

## File inventory

The table records the pre-edit inventory. KEEP includes documentation improved during finalization. All 41 original scripts are retained, plus the release audit and script 44 (43 Python files total). Local caches, review outputs and the safety checkpoint are excluded from publication.

| File/path | Type | State | Decision | Reason |
|---|---|---|---|---|
| .gitignore | documentation | current | KEEP | Release source |
| CHANGELOG.md | documentation | current | KEEP | Empty placeholder; fill with useful navigation or current status |
| PROJECT_INDEX.md | documentation | current | KEEP | Empty placeholder; fill with useful navigation or current status |
| PROJECT_PLAN.md | documentation | current | KEEP | Empty placeholder; fill with useful navigation or current status |
| README.md | documentation | current | KEEP | Empty placeholder; fill with useful navigation or current status |
| Transparency_in_Coverage_PUF.xlsx | raw source | current | IGNORE | Duplicate root copy kept locally; canonical raw path used by pipeline |
| docs/ANALYTICAL_STORY.md | documentation | current | KEEP | Release source |
| docs/ARCHITECTURE.md | documentation | current | KEEP | Empty placeholder; fill with useful navigation or current status |
| docs/BUSINESS_CONTRACT.md | documentation | current | KEEP | Empty placeholder; fill with useful navigation or current status |
| docs/DATA_CONTRACT.md | documentation | current | KEEP | Release source |
| docs/DATA_DICTIONARY.md | documentation | current | KEEP | Empty placeholder; fill with useful navigation or current status |
| docs/ENVIRONMENT_BASELINE.txt | documentation | current | KEEP | Release source |
| docs/KPI_CONTRACT.md | documentation | current | KEEP | Release source |
| docs/POWER_BI_ARCHITECTURE.md | documentation | current | KEEP | Release source |
| docs/SOURCE_HASH.txt | documentation | current | KEEP | Release source |
| docs/VALIDATION.md | documentation | current | KEEP | Empty placeholder; fill with useful navigation or current status |
| screenshots/00 INDEX.png | screenshot | current | KEEP | Final image; preserve bytes |
| screenshots/01 Executive Overview.png | screenshot | current | KEEP | Final image; preserve bytes |
| screenshots/02 Network Comparison.png | screenshot | current | KEEP | Final image; preserve bytes |
| screenshots/03 Denials.png | screenshot | current | KEEP | Final image; preserve bytes |
| screenshots/04 Denial Reasons.png | screenshot | current | KEEP | Final image; preserve bytes |
| screenshots/05 Appeals.png | screenshot | current | KEEP | Final image; preserve bytes |
| screenshots/06 Resubmissions.png | screenshot | current | KEEP | Final image; preserve bytes |
| screenshots/07 Enrollment.png | screenshot | current | KEEP | Final image; preserve bytes |
| screenshots/08 Issuer & State Explorer.png | screenshot | current | KEEP | Final image; preserve bytes |
| screenshots/09 Data Availability.png | screenshot | current | KEEP | Final image; preserve bytes |
| screenshots/10 Data Quality & Methodology.png | screenshot | current | KEEP | Final image; preserve bytes |
| data/canonical/dim_availability_status.csv | canonical data | current | KEEP | Small governed outputs for immediate reuse |
| data/canonical/dim_issuer.csv | canonical data | current | KEEP | Small governed outputs for immediate reuse |
| data/canonical/dim_metric.csv | canonical data | current | KEEP | Small governed outputs for immediate reuse |
| data/canonical/dim_plan.csv | canonical data | current | KEEP | Small governed outputs for immediate reuse |
| data/canonical/dim_reporting_period.csv | canonical data | current | KEEP | Small governed outputs for immediate reuse |
| data/canonical/dim_state.csv | canonical data | current | KEEP | Small governed outputs for immediate reuse |
| data/canonical/dq_exception_register.csv | canonical data | current | KEEP | Small governed outputs for immediate reuse |
| data/canonical/fact_issuer_metric_availability.csv | canonical data | current | KEEP | Small governed outputs for immediate reuse |
| data/canonical/fact_issuer_transparency.csv | canonical data | current | KEEP | Small governed outputs for immediate reuse |
| data/canonical/fact_plan_metric_availability.csv | canonical data | current | KEEP | Small governed outputs for immediate reuse |
| data/canonical/fact_plan_transparency.csv | canonical data | current | KEEP | Small governed outputs for immediate reuse |
| data/raw/Transparency_in_Coverage_PUF.xlsx | raw source | current | KEEP | Public CMS source; hash-pinned |
| powerbi/TransparencyInCoverage.pbip | PBIP entry | current | KEEP | Saved project entry |
| powerbi/TransparencyInCoverage.Report/ | PBIR definitions | current | KEEP | Saved report, pages, visuals, resources |
| powerbi/TransparencyInCoverage.SemanticModel/ | TMDL model | current | KEEP | Saved model, DAX, M, relationships |
| powerbi/**/.pbi/ | local cache | generated | IGNORE | Local settings, editor state, imported cache |
| src/**/__pycache__/ | bytecode | generated | IGNORE | Regenerated by Python |
| sql/; tests/ | empty directories | placeholder | NOT TRACKED | No implemented SQL layer or test suite existed |
