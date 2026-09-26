# Security and redaction

This project uses a public-use insurer/plan transparency dataset and does not contain patient-level PHI. This statement is a dataset-scope description, not a legal compliance certification.

Publication candidates are scanned by content for credential patterns, authorization headers, private keys, email addresses, user paths and local review locations. Raw workbook XML and image metadata are inspected separately during release preparation. Generic words such as `token` in parsers and `Authorization` in a metric name are reviewed as benign code/domain terms rather than treated as credentials.

Personal environment paths were removed from documentation. Active scripts derive project paths and write evidence to ignored `.local-review/` or an explicit environment override. Historical scripts use Windows environment/registry discovery and remain cataloged as non-release tooling. The saved model's non-personal absolute `DataFolderPath` parameter is intentionally preserved and documented as local setup; no connection password is embedded there.

Local cache state, backup ZIPs, temporary review workbooks, environment files, private key files and the duplicate root workbook are excluded from Git. Git authorship uses the connected account's GitHub noreply identity. No credentials are stored in repository configuration or source.

The scan is a release hygiene check, not a guarantee against every possible secret format. Exact release outcomes and unresolved visual limitations are in [the checklist](GITHUB_RELEASE_CHECKLIST.md).
