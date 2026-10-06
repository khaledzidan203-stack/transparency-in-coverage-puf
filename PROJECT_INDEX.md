# Project index

Start with the [portfolio README](README.md), then the [implemented architecture](docs/ARCHITECTURE.md).

| Area | Entry point |
|---|---|
| Immutable source | [data/raw](data/raw/) and [provenance policy](data/README.md) |
| Canonical datasets | [11 CSV tables](data/canonical/), [dictionary](docs/DATA_DICTIONARY.md) |
| Discovery | [01_workbook_inventory.py](src/discovery/01_workbook_inventory.py) |
| Transformation | [02_build_canonical_dataset.py](src/transformation/02_build_canonical_dataset.py), overwrites canonical outputs |
| Analysis | [04_run_eda.py](src/analysis/04_run_eda.py), generates local evidence |
| KPI validation | [03_validate_kpis.py](src/validation/03_validate_kpis.py) |
| Semantic validation | [07_validate_powerbi_semantic_model.py](src/validation/07_validate_powerbi_semantic_model.py) |
| Optional live DAX | [09_runtime_dax_reconciliation.py](src/validation/09_runtime_dax_reconciliation.py) |
| PBIR audit | [11_audit_powerbi_report_structure.py](src/validation/11_audit_powerbi_report_structure.py) |
| Release checks | [release_audit.py](src/validation/release_audit.py) |
| Validation safety | [validation directory guide](src/validation/README.md) |
| Semantic model | [TMDL definition](powerbi/TransparencyInCoverage.SemanticModel/definition/) |
| Saved report | [PBIP entry](powerbi/TransparencyInCoverage.pbip), [PBIR definition](powerbi/TransparencyInCoverage.Report/definition/) |
| Page images | [screenshots](screenshots/), [page guide](docs/REPORT_PAGE_GUIDE.md) |
| Contracts | [data](docs/DATA_CONTRACT.md), [KPIs](docs/KPI_CONTRACT.md), [business scope](docs/BUSINESS_CONTRACT.md) |
| Current relationship amendment | [implementation contract amendment](docs/IMPLEMENTATION_CONTRACT_AMENDMENT.md) |
| DAX/KPI semantic review | [DAX ↔ KPI contract audit](docs/DAX_KPI_CONTRACT_AUDIT.md) |
| Interpretation | [analytical story](docs/ANALYTICAL_STORY.md), [DQ governance](docs/DATA_QUALITY_GOVERNANCE.md) |
| Validation evidence | [framework](docs/VALIDATION_FRAMEWORK.md), [summary](docs/VALIDATION_EVIDENCE.md) |
| Release review | [checklist](docs/GITHUB_RELEASE_CHECKLIST.md), [security](docs/SECURITY_AND_REDACTION.md), [limitations](docs/KNOWN_LIMITATIONS.md) |
| Python environment | [requirements](requirements.txt), [recorded baseline](docs/ENVIRONMENT_BASELINE.txt) |
| CI gate | [.github/workflows/portfolio-validation.yml](.github/workflows/portfolio-validation.yml) |
| Approved DQ correction | [44_fix_dq_scope_filter_context.py](src/validation/44_fix_dq_scope_filter_context.py), historical mutating patch; already applied, do not rerun |
| Historical engineering | [complete script catalog](docs/REPOSITORY_INVENTORY.md); retained in place, not a normal pipeline |

No physical archive was created: keeping numbered dependencies in place avoids breaking historical audits. All original validation and engineering scripts remain available. Empty `sql/` and `tests/` directories are not advertised as implemented components.
