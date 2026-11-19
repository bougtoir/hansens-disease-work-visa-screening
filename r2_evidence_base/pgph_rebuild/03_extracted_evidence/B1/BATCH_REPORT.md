# Batch B1 report - Sub-Saharan Africa (West/Central), 25 jurisdictions

Coder: Devin B1 batch agent (single coder, not independently verified). Date: 2026-11-19 (UTC).
Historical aggregate counts were not used as targets or evidence.

## Searches and documents
- 89 logged web searches (`SEARCH_LOG.csv`; full hit lists in `01_raw_search_logs/B1/raw/Q-B1-*.md`, machine log `search_log.jsonl`). Every jurisdiction has an English pass and an official/local-language pass (fr / pt / es; en is the official language for GHA, GMB, LBR, NGA, SLE). P3/P4 passes were run for a subset (see `stage`).
- 95 retrieved files saved under `02_primary_sources/B1/` with UTC time, bytes and SHA-256 (`SOURCE_REGISTER.csv`, `01_raw_search_logs/B1/fetch_log.csv`). Tier counts: S1 22, S2 1, S3 36, S4 2, S5 26, S7 8 (HTTP errors).
- Extracted text (pdftotext layout/raw; HTML stripped) and OCR text (tesseract fra: Benin Loi 2025-15, Congo Loi 23-96) are in `01_raw_search_logs/B1/text/`. OCR outputs are derived and may be partial.
- Every `source_wording_original` in `EVIDENCE_ROWS.csv` was checked by `tools/build_tables.py` to occur in the saved source text (whitespace/quote/line-break-hyphen normalised). Snippets were never used for classification.
- Budget: initial passes only; no selective extension was triggered. Per-jurisdiction hours were not timed (`hours_spent` blank).
- Checkpoint commits: the work was committed in a small number of commits at the end rather than every ~5 jurisdictions (deviation from instructions).

## Results (verified vs unresolved)
| Status | Jurisdictions |
|---|---|
| DSS documented_positive, HD confirmed | J-LBR |
| DSS documented_absence | none |
| DSS unresolved, HD cannot_determine | the other 24 |

- Liberia: Aliens and Nationality Law (approved 15 May 1973; Refworld PDF S-LBR-006, concordant NGO copy S-LBR-001) §5.1(f) "Aliens who are afflicted with tuberculosis in any form, or with leprosy, or any dangerous contagious disease" are ineligible for visas and excluded; §5.1(a)-(e) name mental conditions, epilepsy, narcotic drug addiction and chronic alcoholism. §5.3(5) excludes leprosy/TB from discretionary bond admission. No treatment/cure distinction was located, so post-treatment risk = `no_treatment_distinction_named_condition`. No explicit post-treatment exclusion was located. Caveat: in-force status of this text was not checked against a current official Liberian publication (confidence medium).
- Generic medical certificate/exam without named conditions (GENERIC_FITNESS only; unresolved): AGO, BEN, COG, CPV, GHA, MLI, NER, TCD, TGO. Generic public-health entry grounds only: GAB, STP (WHO-defined diseases), CPV/AGO (WHO-defined or contagious diseases).
- No medical requirement located in opened S1-S4 sources (unresolved, not absence): BFA, CAF, CMR, COD, GIN, GMB, GNB, GNQ, MRT, NGA, SEN, SLE. GIN, NER, NGA and SEN have medical-certificate references only in S5 sources.
- Observed denials: none located (`OBSERVED_DENIALS.csv` header only). Only three targeted case searches were run (LBR, NGA, GHA).

## Key adjudications
1. Medical certificate requirements with unspecified content were coded GENERIC_FITNESS. They make the pathway `unresolved`, not `documented_absence`.
2. Côte d'Ivoire's official resident-card checklist (S2) lists no medical document. It does not state that no exam is required, and the work-contract visa was not verified, so CIV stays unresolved.
3. The Central African Republic Labour Code art.168 ("maladies contagieuses ou mentales") concerns apprenticeship masters and was not coded as a migration pathway.
4. Burkina Faso's "carnet de santé conformément à la réglementation sanitaire internationale" is a vaccination/health booklet requirement. It was not coded as screening.
5. The Gambia Immigration Bill 2026 (assembly.gm) is not in force and was not used.
6. Tier policy: ILO NATLEX and UNHCR Refworld PDF reproductions of national statutes = S1 (intergovernmental legal database; flagged in notes). Private, NGO and commercial legal hosts (lex.ao, podus, itamba, citizenshiprightsafrica, droitcongolais, unicongo, sierra-leone.org), consultancies, foreign-government advice pages and UNHCR rimap / UNODC Sherloc summary pages = S5. Wayback copies of official pages = S4.

## Inaccessible or incomplete critical sources
- COD Ordonnance du 5 décembre 1953 - police sanitaire de l'immigration (droitcongolais.info, HTTP 503)
- LBR Liberia Immigration Service Service Delivery Charter (HTTP 500)
- SLE slid.gov.sl resident-permits and immigration-regulations (HTTP 404)
- SEN interieur.gouv.sn carte d'identité d'étranger (HTTP 404)
- MRT procedures.gov.mr category 64 (no response)
- AGO Consulado Porto work visa page (certificate error, then 404)
- COG Fragomen medical-assessment note (404; S5 lead only)
- NGA lawnigeria Immigration Act 2015 (403); official full text of the Act and of the 2017 Regulations was not machine-readable or was not located
- BEN Loi 2025-15 and COG Loi 23-96 are scanned; OCR only, possibly partial
- GHA content of the Ghana Immigration Clinic medical report: official specification not located (secondary leads in Q-B1-0089 not used)

## Reproduction
`python3 01_raw_search_logs/B1/tools/build_tables.py` builds all tables from `search_log.jsonl`, `fetch_log.csv` and the adjudication specs (`tools/evidence_spec.json`, `pathway_spec.json`, `jstatus_spec.json`, `source_titles.json`). Then run `python3 scripts/validate_batch.py B1`, which gives PASS.
