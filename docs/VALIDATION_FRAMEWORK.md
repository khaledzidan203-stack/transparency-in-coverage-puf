# Validation framework

## Scope and safety

Validation separates source/data integrity, model structure, runtime measures and report structure. A script can write an evidence workbook while remaining read-only with respect to the model and data. The [script catalog](REPOSITORY_INVENTORY.md) identifies mutating engineering scripts retained for provenance. Do not batch-run the numbered directory.

## Data and KPI checks

`03_validate_kpis.py` checks the hash-pinned source, required canonical files, dimension keys, fact grain, availability integrity, eligible populations, ratios of totals and exception context. The latest release run passed 32 of 32 acceptance checks with `PASS_WITH_CONTEXT`: preserved source exceptions are expected context, not repository failures. It retains issuer and plan grains and does not overwrite canonical outputs.

## Semantic model checks

`07_validate_powerbi_semantic_model.py` checks table inventory, relationship endpoints, active/single-direction structure, cardinality, absence of auto-date tables, canonical integrity, CSV loader and parameter presence. Its pre-DAX expectations (11 tables, zero measures) were stale. Only those release expectations were updated to include `_Measures` and 43 unique measures; model source was not changed. The current validator is release-specific, not a general TMDL parser or DAX execution engine.

## Runtime DAX reconciliation

`09_runtime_dax_reconciliation.py` compares live DAX with Python baselines using declared tolerances. The fresh local workbook `.local-review/09_DAX_RUNTIME_RECONCILIATION.xlsx`, generated after the approved fix and inspected read-only during this update, records PASS: 32 core, 10 reason, 18 availability and 3 DQ context checks, zero failures. Runtime reconciliation did not modify project or Power BI source. These checks cover the recorded queries, not every possible visual interaction. A sanitized summary is retained in [validation evidence](VALIDATION_EVIDENCE.md); the temporary workbook remains local.

## PBIR and release checks

`11_audit_powerbi_report_structure.py` parses report JSON and inventories pages/visuals without altering them. `release_audit.py` adds source hashes, page ordering, navigation references, PBIP/model paths, Python compilation, local documentation links and a scan of publication candidates. It does not claim full Microsoft JSON-schema validation or live visual rendering. Report JSON parse success cannot prove that a measure behaves correctly under every visual filter.

## Data quality and immutable design

36 open-review rows, 2 preserved known-source rows and 161 diagnostic observations retain their distinct meanings. No clipping, imputation or cross-grain summing is permitted. Final source hashes are compared with the reviewed public-release checkpoint. The release audit accepts an explicit `--baseline` manifest for an approved update and otherwise compares protected tracked files with Git HEAD, avoiding the superseded private-release checkpoint. The page-10 scope-chart issue is resolved: ISSUER=5, PLAN=33, Total=38. See [release checklist](GITHUB_RELEASE_CHECKLIST.md) for verified results and limits.
