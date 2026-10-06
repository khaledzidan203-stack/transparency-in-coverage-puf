# Changelog

## Evidence-first presentation finalization - 2026-10-06

- Added an evidence-aligned project overview visual that uses only validated project facts.
- Added a concise end-to-end case study and a claim-to-evidence map for easier technical review.
- Reorganized the README so scope, architecture, governance, Power BI delivery, validation and safe reproduction are understandable in sequence.
- Added direct navigation between the README, case study, technical walkthrough, evidence map and validation evidence.
- Kept all presentation assets explicitly non-authoritative; governed contracts, canonical data, saved Power BI source and validation evidence remain the source of truth.
- Removed the earlier candidate visual from active presentation because it did not match the governed project scope closely enough.
- Preserved canonical data, transformation logic, DAX, TMDL, PBIR, Power Query, screenshots and validated analytical results.

## Portfolio hardening update - 2026-10-06

- Added an implementation contract amendment so documentation reflects the saved TMDL relationship topology without mutating the semantic model.
- Added a DAX-to-KPI contract audit documenting semantic-drift observations while preserving the validated DAX source.
- Added `requirements.txt` and a safe validation runbook that separates read-only validation from historical mutating engineering scripts.
- Added GitHub Actions static validation and protected-path checks for canonical data, Power BI source, and report screenshots.
- Reorganized the README around engineering scope, governed analytical behavior, validation evidence, and safe reproduction.
- Added a technical walkthrough and refreshed project navigation/status documentation.
- Preserved canonical data, transformation logic, DAX, TMDL, PBIR, Power Query, screenshots, and validated analytical results.

## Final public-release update - 2026-09-26

- Included the approved KEEPFILTERS correction in Current Context DQ Exception Count; no other DAX or analytical logic changed.
- Resolved page 10: ISSUER=5, PLAN=33, Total=38; included the supplied corrected screenshot.
- Verified fresh runtime evidence: 32 core, 10 reason, 18 availability and 3 DQ checks, zero failures.
- Retained script 44 as historical mutating engineering evidence without executing it.
- Preserved the saved active-page selection (page 10); no visual geometry, colors, Power Query or relationships changed.
- Updated release documentation and made preservation checks use Git HEAD or an explicit reviewed manifest.

## v1.0.0 — Portfolio Release — 2026-09-26

- Prepared the existing canonical analytical pipeline and 11 CSV datasets for source control.
- Preserved the governed semantic model, 43 DAX measures and 11-page Power BI report, including manual refinements.
- Added portfolio README, architecture, page guide, validation framework, security notes, provenance policy and release checklist.
- Retained all original validation and historical engineering scripts; classified their side effects and dependencies.
- Made active review-output paths portable and corrected the semantic validator's stale pre-DAX inventory expectations.
- Preserved data-quality controls and historical runtime DAX reconciliation evidence; reran safe KPI/model/PBIR checks.
- Excluded local caches, temporary evidence, backup archives and a byte-identical root workbook from publication.
- Documented the page-10 scope-chart issue during the initial private release; resolved by the final public-release update above.
- Published to a private GitHub repository on main, with ten relevant topics, without installing dependencies or creating a software license.
