# Hansen's disease in work-related migration medical requirements — 199-jurisdiction evidence base (R2)

Data and code for the scoping review "Hansen’s disease in work-related migration medical requirements: an evidence-first scoping review of 199 jurisdictions".

## Contents
- `pgph_rebuild/data/` — 199-jurisdiction analytical frame and schema.
- `pgph_rebuild/03_extracted_evidence/` — batch-level extraction records and search logs (nine regional passes).
- `pgph_rebuild/04_validated_master/` — validated master tables: `JURISDICTION_MASTER.csv`, `EVIDENCE_ROWS.csv` (raw) and `EVIDENCE_ROWS_ADJUDICATED.csv`, `POLICY_PATHWAYS.csv`, `SAFEGUARDS.csv`, `POST_TREATMENT_RISK.csv`, `OBSERVED_DENIALS.csv`, `NATIONALITY_ORIGIN_TRIGGERS.csv`, `SEARCH_LOG.csv`, `SOURCE_REGISTER.csv`.
- `pgph_rebuild/05_analysis/` — documented adjudication files (every manual decision with its reason) and results, including `rebuild_numbers.json`, the single source of all numbers reported in the manuscript.
- `pgph_rebuild/07_audits/` — independent second-pass validation rows (AI agent), targeted legal checks, HD count reconciliation and the provision-level legal-consequence audit.
- `data/raw/who_gho/` — WHO Global Health Observatory indicator NTD_LEPR2 as retrieved; `data/raw/naturalearth/` — Natural Earth 1:50m boundaries (public domain).

## Primary source documents
Archived copies of the 768 locally saved primary sources are not redistributed here, because redistribution rights were not cleared for most of them. `SOURCE_REGISTER.csv` lists, for each source, the URL, archive URL where available, retrieval time (UTC), byte size and SHA-256, so that copies can be retrieved and verified. Short quotations needed to verify each coding decision are included in the evidence tables.

## Reproduce
```
pip install -r requirements.txt
bash run_public.sh
```
This regenerates all counts, the legal-consequence coding, the WHO burden comparison (bootstrap seed fixed) and Figures 1–6, and checks the key outputs against `EXPECTED_OUTPUTS.sha256`. `SHA256SUMS.txt` lists checksums for every file in this package.

## Status
Legal codings were verified against located sources by the authors' AI-assisted workflow and an independent AI second pass; these are not a substitute for review by a qualified lawyer in each jurisdiction. Absence of a provision in these tables does not mean that none exists.
