# Validation Directory Safety Guide

This directory contains both **current read-only validation tools** and **historical mutating engineering scripts** retained for provenance and path dependencies. Do **not** batch-run the directory by filename order.

## Safe current validation entry points

These are the intended release/review commands:

```powershell
python src/validation/03_validate_kpis.py
python src/validation/07_validate_powerbi_semantic_model.py
python src/validation/11_audit_powerbi_report_structure.py
python src/validation/release_audit.py
```

Optional live runtime check:

```powershell
python src/validation/09_runtime_dax_reconciliation.py
```

The runtime DAX check requires Power BI Desktop with this PBIP open and a compatible DAX Studio CLI installation.

## Historical mutating scripts

Scripts that build, patch, repair or install report/model content are retained as historical engineering evidence and are **not** reproduction steps. Examples include relationship repair, governed-measure installers, report builders, page-layout patches and semantic-model fixes.

The authoritative classification for every script is:

`docs/REPOSITORY_INVENTORY.md`

In particular, `44_fix_dq_scope_filter_context.py` is a historical targeted DAX patch that has already been applied. Do not rerun it as a validator.

## Release safety rule

For routine review or CI, use `release_audit.py`. It performs static/publication checks and does not execute the analytical pipeline or historical PBIR/model builders.

Protected release artifacts are:

- `data/`
- `powerbi/`
- `screenshots/`

Changes to those paths require explicit analytical review and revalidation.