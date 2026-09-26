# Known limitations and release status

## Resolved page-10 scope chart

**PASS - ISSUER=5, PLAN=33, Total=38.** The previous scope discrepancy is resolved. The approved correction wraps the two Scope predicates in `KEEPFILTERS` inside `Current Context DQ Exception Count`, preserving the visual's current filter context. The supplied corrected page-10 screenshot was inspected and replaces the earlier image.

Fresh runtime DAX reconciliation: **PASS**, 32 core checks, 10 reason checks, 18 availability checks, 3 DQ context checks and zero failures. The local workbook was inspected read-only for this release; its sanitized summary and SHA-256 are in [validation evidence](VALIDATION_EVIDENCE.md). It remains ignored and uncommitted.

## Reproduction and evidence boundaries

- Set the existing `DataFolderPath` parameter to the clone's canonical directory before refresh. Local import cache is excluded.
- Passing runtime checks cover the recorded queries; they do not certify every possible interaction. Full Microsoft JSON-schema validation is separate from the passing local structural checks.
- Historical builders and semantic-model patches can overwrite saved refinements. Do not rerun them as reproduction steps, including script 44.
- A software licensing choice remains open. No LICENSE was invented. Source-data terms are documented in [data policy](../data/README.md).

## Analytical interpretation

CMS data are self-reported and may be revised. Availability and comparable populations vary by metric; source anomalies are preserved. Results do not establish causal effects, insurer quality or full-market totals. Resubmission intensity can exceed 100, reason fields need not exhaust all denied claims, and monthly enrollment relationships are not annual churn.
