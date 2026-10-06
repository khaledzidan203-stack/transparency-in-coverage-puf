# Project Evidence Map

This map connects the repository's main public claims to the files that define or validate them. It is intended to make the implementation easy to review without treating the README or presentation graphics as evidence by themselves.

| Project claim | Primary evidence |
|---|---|
| Source is CMS Transparency in Coverage PUF, PY2026 / experience 2024 | [`data/README.md`](../data/README.md), [`DATA_CONTRACT.md`](DATA_CONTRACT.md) |
| Scope contains 4,956 plans, 348 issuers and 30 states | [`DATA_CONTRACT.md`](DATA_CONTRACT.md), canonical datasets in [`data/canonical/`](../data/canonical/) |
| Canonical analytical layer contains 11 governed CSV tables | [`data/canonical/`](../data/canonical/), [`DATA_DICTIONARY.md`](DATA_DICTIONARY.md) |
| Issuer and plan grains are separate | [`DATA_CONTRACT.md`](DATA_CONTRACT.md), [`ARCHITECTURE.md`](ARCHITECTURE.md) |
| Availability is modeled separately from numeric facts | [`DATA_CONTRACT.md`](DATA_CONTRACT.md), [`KPI_CONTRACT.md`](KPI_CONTRACT.md), [`ARCHITECTURE.md`](ARCHITECTURE.md) |
| Saved semantic model has 43 measures and 14 active single-direction relationships | [`ARCHITECTURE.md`](ARCHITECTURE.md), [`powerbi/TransparencyInCoverage.SemanticModel/definition/`](../powerbi/TransparencyInCoverage.SemanticModel/definition/) |
| Physical relationship topology differs from older conceptual sketches | [`IMPLEMENTATION_CONTRACT_AMENDMENT.md`](IMPLEMENTATION_CONTRACT_AMENDMENT.md) |
| Current DAX release is preserved while latent contract differences are documented | [`DAX_KPI_CONTRACT_AUDIT.md`](DAX_KPI_CONTRACT_AUDIT.md) |
| DQ register contains 36 OPEN_REVIEW rows and 2 preserved known-source rows | [`DATA_QUALITY_GOVERNANCE.md`](DATA_QUALITY_GOVERNANCE.md), canonical DQ register |
| Source anomalies are preserved rather than silently repaired | [`DATA_QUALITY_GOVERNANCE.md`](DATA_QUALITY_GOVERNANCE.md), [`ANALYTICAL_STORY.md`](ANALYTICAL_STORY.md) |
| Power BI release contains 11 pages and 371 saved visuals | [`REPORT_PAGE_GUIDE.md`](REPORT_PAGE_GUIDE.md), [`screenshots/`](../screenshots/), saved PBIR definitions |
| Runtime DAX reconciliation has 63 checks and zero failures | [`VALIDATION_EVIDENCE.md`](VALIDATION_EVIDENCE.md) |
| Repository release checks include structure, links, security and preservation | [`GITHUB_RELEASE_CHECKLIST.md`](GITHUB_RELEASE_CHECKLIST.md), [`src/validation/release_audit.py`](../src/validation/release_audit.py) |
| GitHub-hosted validation is static/read-only for protected analytical artifacts | [`.github/workflows/portfolio-validation.yml`](../.github/workflows/portfolio-validation.yml) |
| Safe versus mutating validation scripts are explicitly classified | [`src/validation/README.md`](../src/validation/README.md), [`REPOSITORY_INVENTORY.md`](REPOSITORY_INVENTORY.md) |
| Project limitations and interpretation boundaries are explicit | [`KNOWN_LIMITATIONS.md`](KNOWN_LIMITATIONS.md), [`ANALYTICAL_STORY.md`](ANALYTICAL_STORY.md) |

## Evidence precedence

When public-facing documentation, historical design notes and saved implementation differ, use this precedence for the current release:

1. governed source/canonical data and explicit contracts for business meaning and grain;
2. saved TMDL/PBIP/PBIR for implemented model/report structure;
3. implementation amendments and semantic audits for documented differences;
4. validation evidence and release checklist for tested release status;
5. README and presentation assets for navigation and explanation only.

Presentation graphics are intentionally non-authoritative. They must summarize evidence already present in the repository and must not introduce new metrics or claims.
