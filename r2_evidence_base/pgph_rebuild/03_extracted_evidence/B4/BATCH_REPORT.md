# B4 South/Central Asia - Batch report

Evidence-first rebuild. Single-reviewer extraction by the B4 agent; **independent verification pending**. Historical aggregate counts were not used. Search snippets were never used as evidence; every evidence row's `source_wording_original` was machine-checked to occur verbatim in the saved local source text (`01_raw_search_logs/B4/tools/build_tables.py`).

## Volume

- Queries run: 62 (DuckDuckGo HTML via local Chrome CDP; SERPs saved in `01_raw_search_logs/B4/serp/`). SEARCH_LOG has 337 rows (top-5 hits per query + failed direct retrievals).
- Source register: 82 rows; 58 saved non-S7 source files in `02_primary_sources/B4/` (SHA-256 + UTC time recorded); failed/challenge captures kept separately in `01_raw_search_logs/B4/failed_fetches/` as S7.
- Evidence rows: 57. Observed denials: 0 (no case-level evidence located; 2 P3 case queries for MDV/KAZ returned no case reports).
- Validator: `python3 scoping-review-hansens-disease/pgph_rebuild/scripts/validate_batch.py B4` -> PASS.

## Jurisdiction summary

| Jurisdiction | Queries | Local languages | Best tier | DSS | HD named | Post-treatment risk |
|---|---|---|---|---|---|---|
| J-AFG | 5 | Pashto;Persian/Dari | S3 | unresolved | cannot_determine | cannot_determine |
| J-BGD | 6 | Bengali | S1 | unresolved | cannot_determine | cannot_determine |
| J-BTN | 3 | Dzongkha | S2 | documented_positive | not_established | not_applicable_hd_not_named |
| J-IND | 6 | Hindi | S1 | documented_positive | not_established | not_applicable_hd_not_named |
| J-IRN | 4 | Persian/Dari | S3 | unresolved | cannot_determine | cannot_determine |
| J-KAZ | 6 | Kazakh;Russian | S1 | documented_positive | confirmed | no_treatment_distinction_named_condition |
| J-KGZ | 5 | Kyrgyz;Russian | S3 | unresolved | cannot_determine | cannot_determine |
| J-LKA | 4 | Sinhala;Tamil | S1 | documented_positive | not_established | not_applicable_hd_not_named |
| J-MDV | 5 | Dhivehi | S2 | documented_positive | confirmed | no_treatment_distinction_named_condition |
| J-NPL | 3 | Nepali | S3 | unresolved | cannot_determine | cannot_determine |
| J-PAK | 3 | Urdu | S5 | unresolved | cannot_determine | cannot_determine |
| J-TJK | 4 | Russian;Tajik | S1 | documented_positive | cannot_determine | cannot_determine |
| J-TKM | 4 | Russian;Turkmen | S1 | documented_positive | cannot_determine | cannot_determine |
| J-UZB | 4 | Russian;Uzbek | S1 | unresolved | cannot_determine | cannot_determine |

## Verified (S1-S4, exact wording saved)

- **J-KAZ** - MoH Order No.664 (30.09.2011, Adilet V1100007274, S1): list of diseases whose presence prohibits entry of foreigners: '1. Наркомания. 2. Психические расстройства (заболевания). 3. Туберкулез. 4. Лепра (болезнь Гансена). ...'. HD confirmed. No treatment-status distinction. Caveat: no repeal marker in the captured Adilet page, but current force/successor order not independently verified (extension reason recorded).
- **J-MDV** - MoH/HPA Pre-Departure Health Assessment Guideline MOH-HPA/G/25/201-0 v1.0 (14 Aug 2025, S2): leprosy screening for applicants from leprosy-endemic countries (s.7.7); Table 7.7 Grade B 'Skin patch with loss of sensation'; Grade B = 'ineligible for employment in Maldives'. Also HIV, HBsAg, HCV, VDRL/RPR, CXR, LF/malaria antigen (endemic origin). Earlier form MOH-QA/F/17/67-4 asks 'Have you ever been diagnosed with leprosy?'. HD confirmed; origin (endemic-country) trigger; sign-based grading without treatment-status distinction.
- **J-BTN** - MoH Form D (S2) for non-immigrant work permit: TB history + CXR, syphilis, malaria, HBV, HCV, HIV, drug panel, mental illness history; examines 'Hypo Pigmented patch' and 'Cutaneous Nerve Enlargement' but does not name leprosy -> HD not established (flagged `explicit_named=ambiguous`).
- **J-LKA** - IHAP (MoH, S2/S3): residence-visa applicants screened for malaria, filariasis, HIV, pulmonary TB. HD not in list.
- **J-IND** - MEA FRRO checklist (S3): HIV test report for ages 15-60 on visas > 1 year. Immigration and Foreigners Act/Order/Rules 2025: generic public-health grounds only. HD not established.
- **J-TJK** - MFA consular (S3): HIV certificate for visas > 3 months; Migration Service (S3): certificate of absence of drug addiction and dangerous infectious diseases on a Government-approved list + HIV certificate. HD cannot be determined (list not obtained).
- **J-TKM** - Law On Migration 2012 Art.11(1)(6) (S1 official copy): visa refusal for HIV, venereal disease, drug addiction or other disease on MoH list. HD cannot be determined (MoH list not obtained).

## Unresolved (not absence)

- **J-AFG** - Only S5 summary of a 2020 (pre-2021) medical-certificate requirement; official MoLSA procedure inaccessible; MFA visa page lists no medical item. Reason: official MoLSA work-permit procedure inaccessible; current-regime rules not located
- **J-BGD** - Visa Policy 2026 (S2) lists no medical certificate for E-visas; Foreigners Order 1951 (S5 only) names mental unsoundness and generic loathsome/infectious disease; official BIDA work-permit text not verified (mirror only). Reason: official text of Foreigners Order 1951 (NATLEX 403) and official BIDA guideline not obtained
- **J-IRN** - Executive Regulation (S5 copies) and work-permit instruction (S5) contain no medical provision in captured text; official Majlis/MFA sources inaccessible. Reason: official sources (rc.majlis.ir SSL failure, MFA PDF redirect loop, Refworld 403) inaccessible
- **J-KGZ** - Only news (S5) of HIV test replaced by form 086/U from Jan 2025; official instrument text (Cabinet Resolution 437) not captured. Reason: operative text of Cabinet Resolution 437 not loaded; no official disease list
- **J-NPL** - Rule 28 (S5 English mirror): visa cancellation for insanity, contagious/communicable disease, failure to produce international health certificate; official Nepali text not saved. Reason: official Nepali text not obtained; S5 only
- **J-PAK** - Foreigners Order 1951 (S5 reproduction) names unsound mind and generic loathsome/infectious disease; official Pakistan Code text not obtained. Reason: official text not obtained; no work-visa medical requirement located
- **J-UZB** - News reports (not saved as evidence) of Law ZRU-1095 (2025) denying work permits for HIV/TB; operative text not captured (lex.uz/nrm.uz JS shells). Reason: operative text of ZRU-1095 / work-permit regulation not obtained

## Key adjudications

- Generic 'loathsome or infectious disease', 'contagious or communicable disease', 'public health', 'medical reasons' wording (BGD, PAK, NPL, LKA Act, IND) coded GENERIC_FITNESS, never HD.
- Mental-condition grounds (unsound mind/insane) coded MENTAL; for BGD/PAK/NPL they rest only on S5 reproductions, so DSS kept unresolved.
- Bhutan skin-sign items are cardinal leprosy signs but the disease is not named: HD not established, row kept as `ambiguous` for reviewer attention.
- Where an official source references an unobtained disease list (TJK Government list, TKM MoH list), HD = cannot_determine rather than not_established.
- Kazakhstan ҚР ДСМ-162/2020 (lists leprosy A30) is a free inpatient-care list, not an entry rule - saved as context only.
- Post-treatment: no explicit post-treatment exclusion found anywhere. KAZ and MDV are `no_treatment_distinction_named_condition`. For MDV the residual-anaesthetic-patch risk is an inference from the wording, not a stated rule.
- Kyrgyzstan form 086/U and Uzbekistan ZRU-1095 (HIV/TB) are known only from news; not used to resolve status.

## Inaccessible / failed sources (S7)

- S-B4-IND-ForeignersOrder1948_kanoon-FAIL | https://indiankanoon.org/doc/5736492/
- S-B4-IND-ForeignersOrder1948_refworld-FAIL | https://www.refworld.org/legal/decreees/natlegbod/1948/en/35875
- S-B4-BGD-ForeignersOrder1951_natlex-FAIL | https://natlex.ilo.org/dyn/natlex2/natlex2/files/download/87977/BGD87977.pdf
- S-B4-BGD-bdlaws_old_alpha_index-FAIL | http://old.bdlaws.minlaw.gov.bd/bangladesh-code-alphabetical-index.html
- S-B4-BGD-ForeignersOrder1951_natlex_pdf-NA | https://natlex.ilo.org/dyn/natlex2/natlex2/files/download/87977/BGD87977.pdf
- S-B4-BTN-ImmigrationAct2007_refworld-FAIL | https://www.refworld.org/sites/default/files/2026-03/immigration-act-2007.pdf
- S-B4-BTN-DOI_foreignworker_entry_process-NA | https://www.doi.gov.bt/?page_id=1862&lang=en
- S-B4-BTN-DOI_workpermit_page-NA | https://www.doi.gov.bt/?page_id=1288&lang=en
- S-B4-AFG-MOLSA_procedures_dr-NA | https://molsa.gov.af/dr/%D8%B7%D8%B1%D8%B2%D8%A7%D9%84%D8%B9%D9%85%D9%84
- S-B4-IRN-Majlis_EntryResidenceLaw-NA | https://rc.majlis.ir/fa/law/show/92268
- S-B4-IRN-MFA_employment_foreign_nationals_pdf-NA | https://mfa.gov.ir/files/ecodepen/investment%20regulation/employment%20of%20foreign%20nationals%20in%20iran.pdf
- S-B4-IRN-shenasname_workpermit_instruction-FAIL | https://shenasname.ir/kar/460
- S-B4-IRN-Regs1973_EntryResidence_refworld-FAIL | https://www.refworld.org/legal/decreees/natlegbod/1973/21151
- S-B4-IRN-ekhtebar_EntryResidence_regulation-FAIL | https://www.ekhtebar.ir/%D8%A2%DB%8C%DB%8C%D9%86%E2%80%8C%D9%86%D8%A7%D9%85%D9%87-%D8%A7%D8%AC%D8%B1%D8%A7%DB%8C%DB%8C-%D9%82%D8%A7%D9%86%D9%88%D9%86-%D9%88%D8%B1%D9%88%D8%AF-%D9%88-%D8%A7%D9%82%D8%A7%D9%85%D8%AA-%D8%A7%D8%AA%D8%A8%D8%A7%D8%B9-%D8%A8%DB%8C%DA%AF%D8%A7%D9%86%D9%87/
- S-B4-IRN-eplonline_workpermit_instruction_amend-NA | 
- S-B4-IRN-novinlaw_EntryResidence_law_regs-NA | 
- S-B4-KAZ-lawgovkz_doc61445_pdf-FAIL | http://law.gov.kz/api/documents/61445/rus/download/pdf
- S-B4-KAZ-kncdiz_disease_list_ru_pdf-FAIL | https://kncdiz.kz/files/00012169.pdf
- S-B4-KAZ-kncdiz_disease_list_kk_pdf-FAIL | https://kncdiz.kz/files/00012176.pdf
- S-B4-KGZ-minjust_CabMin_7-30592_ru | https://cbd.minjust.gov.kg/7-30592/edition/13806/ru
- S-B4-KAZ-MigrationLaw477_ru_refworld-FAIL | https://www.refworld.org/sites/default/files/2025-12/zakon_respubliki_kazakhstan_ot_22_iyulya_2011_goda_no_477-iv_o_migracii_naseleniya.pdf
- S-B4-TJK-portalihuquqi_workpermit_rules-FAIL | http://portali-huquqi.tj/publicadliya/view_qonunhoview.php?asosi_id=10613
- S-B4-UZB-lex_ORQ998_2024 | https://lex.uz/uz/docs/-7220850
- S-B4-UZB-nrm_ZRU1095_2025 | https://nrm.uz/contentf?doc=793774_zakon_respubliki

## Files
- Raw: `01_raw_search_logs/B4/` (search_log.jsonl, fetch_log.jsonl, serp/, failed_fetches/, WORKING_NOTES.md, tools/)
- Sources: `02_primary_sources/B4/`
- Tables: `03_extracted_evidence/B4/*.csv`
