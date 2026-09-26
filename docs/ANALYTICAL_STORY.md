---
document_id: TIC-ANALYTICAL-STORY-001
title: Transparency in Coverage PUF — Governed Analytical Story
version: "1.0"
status: APPROVED_FOR_BI_DESIGN
puf_year: 2026
experience_year: 2024
depends_on:
  - docs/DATA_CONTRACT.md
  - docs/KPI_CONTRACT.md
  - data/canonical/*
  - 04_EDA_REPORT.xlsx
---

# Transparency in Coverage PUF — Governed Analytical Story

## 1. Executive Analytical Position

The project should be presented as a transparency and market-pattern analysis of reported claims, denials, resubmissions, appeals, enrollment, and data availability.

It is not a quality-ranking system.

The central analytical story is:

1. denial behavior differs materially between in-network and out-of-network reported claims;
2. market-wide comparisons require common eligible populations because availability differs by metric;
3. reported denial-reason mix is concentrated in a few categories, but the reason fields must not automatically be interpreted as an exhaustive decomposition of all denied claims;
4. appeals show materially different internal vs external overturn patterns;
5. enrollment/disenrollment analysis is strongly limited by source availability and suppression;
6. claims volume is concentrated among a relatively small number of issuers;
7. source anomalies exist and can materially distort local comparisons if they are not explicitly flagged.

---

# 2. Validated Population

```text
PUF Year                         2026
Experience Year                  2024
States                             30
Issuers                           348
Plans                           4,956

Comparable issuer population
for overall denial analysis       303

Comparable plan population
for overall denial analysis     2,808

Fully comparable plan population
for all 10 denial reasons       1,089
```

All comparisons must preserve these population restrictions.

---

# 3. Claims and Denials

## 3.1 Issuer-level rates

Validated ratio-of-totals:

```text
In-network denial rate          18.79%
Out-of-network denial rate      37.16%
Comparable overall rate         20.44%
```

Interpretation:

The reported out-of-network denial ratio is substantially higher than the in-network ratio for the governed eligible issuer populations.

This is a descriptive source pattern only. It does not establish insurer quality, appropriateness of denials, or causality.

## 3.2 Plan-level comparable rate

```text
Comparable overall plan denial rate = 20.27%
Comparable plans                    = 2,808
```

Plan and issuer rates are calculated independently at their approved grains. Their similarity must not be interpreted as permission to combine or double count the two fact tables.

---

# 4. State-Level Pattern

State-level comparable overall denial rates show substantial variation.

The EDA contains state-level patterns, but these must be presented with DQ context because some reported out-of-network denied counts exceed received counts.

Examples of issuer-level source anomalies driving extreme out-of-network ratios include:

```text
Issuer 73836 — Alaska
Issuer 61779 — Kansas
Issuer 37160 — North Dakota
```

Therefore:

- reported state rates remain visible;
- DQ indicators must appear in state/issuer detail;
- high or low state values must not be labeled as better/worse performance;
- sensitivity views may exclude only the specific affected source exceptions and must be labeled explicitly.

---

# 5. Plan Segment Patterns

Comparable plan-level overall denial rates by selected dimensions:

## Market segment

```text
Individual                      20.30%
SHOP                             6.60%
```

The SHOP comparable population is much smaller, so this difference is descriptive and should not be framed as a causal or quality conclusion.

## Plan offering type

```text
SADP                            21.75%
QHP                             20.25%
```

## Plan type

```text
HMO                             21.23%
EPO                             20.63%
PPO                             17.93%
POS                             16.10%
Indemnity                       22.08%  [very small eligible population]
```

Small-population categories must display eligible entity counts.

## Metal level

```text
Catastrophic                    25.68%
Low                             23.34%
Bronze                          20.93%
Silver                          20.58%
High                            19.73%
Gold                            18.04%
Platinum                        17.91%
```

These are observed reported ratios, not product-quality rankings.

---

# 6. Denial Reasons

The denial-reason analysis uses only the 1,089 plans where all ten reason metrics are numerically available.

Reported reason-count composition:

```text
Other                                      30.99%
Administrative reason                      20.56%
Services excluded                          12.48%
Referral / prior authorization required     9.44%
Out-of-network provider                     8.14%
Member not covered                          6.89%
Not medically necessary — non-BH            6.06%
Benefit limit reached                       4.66%
Not medically necessary — BH only           0.65%
Investigational / experimental / cosmetic   0.14%
```

The largest category is `Other`, followed by `Administrative reason`.

Interpretation boundary:

This is the composition of the ten reported reason-count fields within a fully comparable population. It is not automatically the share of all denied claims unless a separate validation proves that the reason fields are exhaustive and mutually exclusive.

---

# 7. Resubmissions

Validated intensity measures:

```text
Issuer — In Network
37.77 resubmission events per 100 denied claims

Issuer — Out of Network
17.98 resubmission events per 100 denied claims

Plan — In Network
35.78 resubmission events per 100 denied claims

Plan — Out of Network
19.44 resubmission events per 100 denied claims
```

Entities above 100 exist and are not automatically DQ failures because the source field is a count of resubmission events and does not establish a one-resubmission-per-denied-claim limit.

Do not label this KPI simply as a bounded “resubmission rate.”

---

# 8. Appeals

Validated issuer-level aggregate overturn ratios:

```text
Internal appeal overturn rate      39.88%
External appeal overturn rate      17.45%
```

Internal and external appeals have different eligible populations and must remain separate analytical measures.

Known source anomaly:

```text
Issuer_ID                          97725
Internal appeals filed             1,203
Internal appeals overturned       16,518
Published overturn %            1,373.07%
```

The published percentage reconciles mathematically to the published numerator and denominator, so this remains a preserved source exception.

Any appeal visual must disclose the DQ flag for this issuer and must never clip the value silently.

---

# 9. Enrollment and Disenrollment

Numerical availability is limited:

```text
Average Monthly Enrollment
Available plans                  2,344 of 4,956
Availability among applicable    51.13%

Average Monthly Disenrollment
Available plans                  1,340 of 4,956
Availability among applicable    29.23%
```

Reported available-population totals:

```text
Average Monthly Enrollment       876,008
Average Monthly Disenrollment    137,632
```

Comparable disenrollment-to-enrollment ratio:

```text
16.37%
```

Interpretation boundary:

The ratio is a reported average-monthly relationship for the comparable population. It is not annual churn, retention probability, or a full-market estimate.

---

# 10. Data Availability as an Analytical Finding

Availability itself is a major part of the story.

Highest governed availability among applicable entities:

```text
Issuer Claims Received — In Network
99.37%
```

Lowest:

```text
Average Monthly Disenrollment
29.23%
```

Selected availability constraints:

```text
Issuer Internal Appeal %             68.99%
Issuer External Appeal %             68.35%
Plan Avg Monthly Enrollment          51.13%
Plan Avg Monthly Disenrollment       29.23%
```

Therefore every BI page using partially available metrics should display an availability/population companion measure.

---

# 11. Claims Concentration

Within the 303-issuer population where all four claims/denial metrics are available:

```text
Top 5 issuer share of claims        33.26%
Top 10 issuer share                 45.79%
Top 20 issuer share                 62.42%
```

This is a volume-concentration finding only.

It must not be labeled as a performance ranking.

---

# 12. Data Quality Context

Canonical DQ register:

```text
Open-review rows                     36
Known source-exception rows           2
Known source-exception entities       1
```

Open-review structure:

```text
Issuer Denied > Received              3
Plan Denied > Received               33
```

Known-source exception:

```text
Issuer 97725 internal appeals anomaly
```

Default analytical policy:

```text
PRESERVE → CALCULATE → FLAG → DISCLOSE
```

No silent correction, deletion, clipping, or imputation.

---

# 13. BI Storyline

The recommended narrative order is:

```text
1. What is covered?
2. How large is the reported claims population?
3. How do denial patterns differ by network?
4. Where do observed patterns differ across states and plan segments?
5. What reported reasons account for denial-reason counts?
6. What happens after denial — resubmissions and appeals?
7. How complete is the source?
8. Where are source anomalies and DQ limitations?
9. Drill into individual issuers and plans.
```

This order keeps the project business-first and prevents users from interpreting isolated ratios without population and availability context.

---

# 14. Approved Interpretation Vocabulary

Use:

```text
reported
observed
higher/lower reported ratio
higher/lower reported volume
concentration
distribution
difference
outlier
source anomaly
availability limitation
suppression
comparable population
```

Avoid:

```text
best
worst
good insurer
bad insurer
quality score
wrongful denial
fraud
poor care
better plan
worse plan
```

unless supported by separate evidence outside this PUF.

---

# 15. Evidence State

```text
CANONICAL DATASET        APPROVED
KPI VALIDATION           APPROVED
EDA                      APPROVED
ANALYTICAL STORY         APPROVED FOR BI DESIGN
POWER BI BUILD           SAVED 11-PAGE REPORT EXISTS (release 2026-09-26)
```
