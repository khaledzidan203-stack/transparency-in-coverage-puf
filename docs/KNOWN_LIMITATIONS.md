# Known limitations and public-release review

## Existing page-10 scope chart

The supplied `10 Data Quality & Methodology.png` displays 38 for both ISSUER and PLAN in **Exceptions by Scope**. The canonical register contains 5 ISSUER rows and 33 PLAN rows (38 total). The saved visual `p10_scope_chart` uses `Current Context DQ Exception Count`, whose DAX explicitly applies each scope and can replace the axis scope filter. Static inspection explains the discrepancy; live interaction has not been retested in this release.

No DAX, visual binding, layout or screenshot was changed because the release task explicitly preserves the validated report. Before public portfolio promotion, review this discrepancy in Power BI and decide on a separately authorized report correction and screenshot refresh. Do not use that chart as a validated breakdown meanwhile.

## Reproduction and evidence

- Set the existing `DataFolderPath` parameter to the clone's canonical directory before refresh. Local import cache is excluded.
- Historical DAX reconciliation passed, but no fresh live runtime run was possible during finalization. Repeat it in an existing configured Desktop/DAX Studio environment before claiming current runtime coverage.
- Microsoft full JSON-schema validation and live interaction testing are distinct from the passing local PBIR structural checks.
- Historical builders can overwrite manual refinements; do not rerun them to open or reproduce the saved report.
- A software licensing choice remains open. No LICENSE was invented. The raw data's public provenance and terms are documented in [data policy](../data/README.md).

## Analytical interpretation

CMS data are self-reported and may be revised. Availability and comparable populations vary by metric; source anomalies are preserved. Results do not establish causal effects, insurer quality or full-market totals. Resubmission intensity can exceed 100, reason fields need not exhaust all denied claims, and monthly enrollment relationships are not annual churn.
