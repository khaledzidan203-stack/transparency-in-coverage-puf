# GitHub release checklist

Final public-release audit: 2026-09-26. The earlier page-10 issue is resolved. This update includes the already completed approved correction, refreshed screenshot and fresh runtime evidence.

| Gate | Result | Evidence / boundary |
|---|---|---|
| README and local links | PASS | Updated release status; 84 local documentation/image links resolve |
| Screenshot coverage | PASS | All 11 PNGs present; corrected page-10 image visually inspected: ISSUER=5, PLAN=33, Total=38 |
| PBIP and semantic definitions | PASS | Valid project/report/model paths; saved TMDL definitions retained |
| KPI / canonical integrity | PASS | Fresh read-only run: 32/32 acceptance checks; PASS_WITH_CONTEXT; zero failures |
| Semantic model checks | PASS | Fresh read-only run: 29/29 checks, 14 active single-direction relationships, 43 measures |
| JSON / PBIR structure | PASS | 390 JSON-format publication files parse; report audit: 385 report JSON files, 11 pages, 371 visuals, zero parse warnings; 20 navigation targets resolve |
| Python compilation | PASS | All 43 Python scripts compile, including script 44; no mutating script executed |
| Secret / redaction scan | PASS | Publication text and raw workbook XML scanned; no detected credential, personal-path or email findings |
| Personal paths in public documentation | PASS | No personal absolute paths or temporary review locations exposed; evidence uses repository-relative paths |
| Required documentation | PASS | README, index, plan, changelog, inventory, architecture, governance, page guide, limitations and validation evidence reflect the corrected state |
| Repository size | PASS | Approximately 9.94 MB across 500 publication files; largest file 1,601,891 bytes |
| Source data policy | PASS | Hash-pinned raw workbook and 11 canonical CSVs unchanged; attribution and terms documented |
| Validation work retained | PASS | Original 41 scripts, release audit and script 44 retained; 44 classified as historical targeted semantic-model patch, not a validator |
| Approved change scope | PASS | Exact DAX comparison confirms only two KEEPFILTERS wrappers inside Current Context DQ Exception Count; other DAX, Power Query, relationships, colors and visual geometry unchanged |
| Preservation during final checks | PASS | 433 tracked data, Power BI and screenshot files match the reviewed public-update checkpoint; no report/model writer executed |
| Additional saved report metadata | PASS | Only activePageName changes from INDEX to page 10; page order and visual definitions unchanged |
| Fresh live runtime DAX reconciliation | PASS | Fresh local workbook inspected: 32 core, 10 reason, 18 availability and 3 DQ context checks; all detail rows PASS; total failures 0 |
| Page-10 scope breakdown | PASS | ISSUER 5 / PLAN 33 / Total 38; previous issue RESOLVED |
| Local review / ZIP / cache exclusion | PASS | Ignore rules verified; only intended publication paths selected for staging; no review XLSX, backup ZIP or Power BI cache is part of this update |
| Full Microsoft schema validation | NOT RUN | Local parsing/reference checks and runtime DAX checks do not constitute full Microsoft schema validation |
| Software license | NOT ADDED | Licensing remains undecided; no license invented |

Fresh runtime reconciliation was already completed before this update and its workbook was inspected read-only. The workbook remains ignored at `.local-review/09_DAX_RUNTIME_RECONCILIATION.xlsx`; [validation evidence](VALIDATION_EVIDENCE.md) records its SHA-256. It is not part of the commit.

## Publication sequence

Repository: [transparency-in-coverage-puf](https://github.com/khaledzidan203-stack/transparency-in-coverage-puf). The existing remote main was fetched and matched local HEAD before this update. Publication uses one coherent commit on main, a normal push, remote-SHA verification, then the authorized visibility change to public. GitHub commit history and repository visibility show the final publication state.

## Portfolio use

Page 10 and fresh runtime validation no longer have unresolved release warnings. Preserve the documented analytical interpretation limits and set DataFolderPath when refreshing another clone. Choose a software license if broader code reuse is intended. See [known limitations](KNOWN_LIMITATIONS.md).
