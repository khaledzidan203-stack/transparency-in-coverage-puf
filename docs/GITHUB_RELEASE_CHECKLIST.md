# GitHub release checklist

Audit date: 2026-09-26. Scope: private repository publication of the saved project. Public promotion remains subject to the separately listed review items.

| Gate | Result | Evidence / boundary |
|---|---|---|
| README structure and local links | PASS | Logical Markdown/HTML structure reviewed; 84 local links resolve; hero and compact gallery reference real images |
| Referenced screenshots exist | PASS | All 11 supplied PNGs present; reviewed without editing |
| PBIP and semantic definitions present | PASS | Entry point, report-model reference, 17 TMDL files |
| KPI and canonical integrity | PASS | 32/32 checks, PASS_WITH_CONTEXT, no failures |
| Semantic model checks | PASS | 29/29 checks after correcting two stale pre-DAX validator expectations; 14 active single-direction relationships, 43 measures |
| PBIR structure | PASS | 11 pages, 371 visuals, 385 report JSON files, zero parse warnings |
| Full Microsoft schema / live rendering validation | NOT RUN | Structural parsing is not full schema or engine validation |
| Python compile | PASS | All 41 original scripts plus the release audit compile; historical scripts not executed |
| Secret and redaction scan | PASS | Publication text and raw workbook XML scanned; no credential or personal-path findings; no email matches |
| Screenshot / workbook metadata | PASS | No PNG text/EXIF metadata chunks; source workbook creator and lastModifiedBy empty; no personal data seen in screenshots |
| Personal temporary paths in public docs | PASS | Removed original local environment paths; source hash uses relative path |
| Required documentation | PASS | README, contracts, index, plan, changelog, architecture, framework, security, page guide and limitations present |
| Ignore rules and staged contents | PASS | 499 intended files staged; git check-ignore confirms local checkpoint, evidence, duplicate root workbook and PBIP cache exclusions |
| Repository size and file limits | PASS | Approximately 10 MB publication files; largest file 1,601,891 bytes; no file near 100 MiB |
| Source data policy | PASS | CMS source and terms linked, hash-pinned raw workbook and 11 small derived CSVs included |
| Duplicates | PASS | Root workbook equals raw SHA-256; excluded, not deleted |
| Validation work retained | PASS | All 41 original scripts retained, including historical engineering; classified in inventory |
| Protected source unchanged | PASS | 436 original data, Power BI and screenshot files match the pre-release SHA-256 checkpoint |
| Temporary review XLSX / ZIP staged | PASS | Zero temporary XLSX or ZIP files in the staged index; only the raw source workbook is included |
| Fresh live runtime DAX reconciliation | NOT RUN | Desktop/live model and DAX Studio CLI prerequisites unavailable; historical workbook records PASS, zero failures |
| Page-10 scope breakdown | FAIL — existing visual issue | Screenshot shows 38 for each scope; canonical register has ISSUER 5 / PLAN 33. Preserved and disclosed, not repaired under this task |
| Software license | NOT ADDED | No existing or unambiguous project license; source-data terms do not license project code |
| Private GitHub push | PASS | Initial main push succeeded; GitHub API and local HEAD both resolved to 76aadf3c081b8704d7892d43d906808905c6ab63; visibility PRIVATE and main default branch verified |

No analytical logic, DAX, Power Query, semantic relationships, report colors or visual geometry changed. No package, CLI, dependency or Git LFS component was installed. No PBIR builder or model patcher was executed.

## Before making public

Review the page-10 discrepancy in Power BI and decide on a separately authorized correction and screenshot refresh. Repeat runtime reconciliation in the configured live environment before claiming fresh runtime coverage. Choose software licensing if reuse is intended. See [known limitations](KNOWN_LIMITATIONS.md).

Repository: [transparency-in-coverage-puf](https://github.com/khaledzidan203-stack/transparency-in-coverage-puf). Ten relevant topics are configured. The follow-up documentation commit records this verified publication result; GitHub history provides the current commit identity.
