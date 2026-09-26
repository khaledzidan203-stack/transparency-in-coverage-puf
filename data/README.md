# Data provenance and repository policy

Source: Centers for Medicare & Medicaid Services, 2026 Transparency Public Use File, reporting 2024 experience. [Official dataset](https://data.healthcare.gov/dataset/dfc1a61d-6e77-4c62-bee1-44422a42cf06) · [CMS documentation entry](https://www.cms.gov/marketplace/resources/data/public-use-files) · [CMS user agreement](https://www.cms.gov/files/document/transparency-coverage-datadisclaimer-py26.pdf-0).

The [federal catalog record](https://healthdata.gov/dataset/Transparency-in-Coverage-PUF-PY2026-g5wq-rx7b-Arch/8jgf-kugu/explore) identifies public access and a public-domain license label. The small, unmodified source workbook (1,079,834 bytes) is therefore retained at `raw/Transparency_in_Coverage_PUF.xlsx`. CMS requests source attribution and makes users responsible for transformed data. The canonical CSVs are project-derived analytical outputs; they are not presented as unmodified CMS data.

The root workbook is byte-identical to the raw copy and is excluded from Git without local deletion. Source SHA-256: `27379dc76590027b0e6d21d736331f87376f0ac87810d68e9f7cca555eef4b1f`.

If the workbook is absent, use the dataset page's Excel resource, save it as `raw/Transparency_in_Coverage_PUF.xlsx`, and compare its SHA-256 with [SOURCE_HASH](../docs/SOURCE_HASH.txt). The published resource is [PY2026 Excel download](https://data.healthcare.gov/datafile/py2026/transparency_in_coverage_PUF.xlsx). Upstream replacements can change the hash; a mismatch requires a separately reviewed source update, not bypassing the integrity check.

All 11 canonical CSVs are committed (approximately 3.1 MB total) for immediate inspection and Power BI refresh. Regenerate only when intended with `python src/transformation/02_build_canonical_dataset.py`. Review Excel files, local caches and development backups are excluded. No Git LFS is needed. Data-source terms do not establish a software license for this repository.
