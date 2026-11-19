# B7 batch report — Europe (South/East), 26 jurisdictions

Evidence-first rebuild. Classification uses only saved, re-read sources (02_primary_sources/B7). Search snippets were never used for classification; failed retrievals are 'unresolved'. Historical aggregate counts were not used.

## Counts
- Queries run: 80 (P1_english 26, P2_local 44, P3_applicability 8, P4_ambiguity 2) — engine: Devin web_search; full hit lists in 01_raw_search_logs/B7/raw/
- Sources saved and registered: 88 (S1 25, S2 10, S3 29, S4 5, S5 12, S6 1, S7 6) with SHA-256 and UTC retrieval time
- Evidence rows: 63 (all quoted wording machine-verified against saved text by tools/build_tables.py, except where noted)
- Pathways coded: 27
- Observed denials: 0 (no case-level refusal evidence located)

## Endpoint results (main work pathway per jurisdiction)
- DSS documented_positive: J-HRV, J-HUN, J-MDA, J-MLT, J-RUS
- DSS documented_absence: J-ESP, J-ITA, J-SMR, J-SRB
- DSS unresolved: J-ALB, J-AND, J-BGR, J-BIH, J-BLR, J-CZE, J-GRC, J-MKD, J-MNE, J-POL, J-PRT, J-ROU, J-SVK, J-SVN, J-UKR, J-VAT, J-XKX
- HD_NAMED confirmed: J-RUS
- HD_NAMED not_established: J-ESP, J-HRV, J-HUN, J-ITA, J-MLT, J-SMR, J-SRB
- HD_NAMED cannot_determine: J-ALB, J-AND, J-BGR, J-BIH, J-BLR, J-CZE, J-GRC, J-MDA, J-MKD, J-MNE, J-POL, J-PRT, J-ROU, J-SVK, J-SVN, J-UKR, J-VAT, J-XKX
- POST_TREATMENT_RISK no_treatment_distinction_named_condition: J-RUS
- POST_TREATMENT_RISK not_applicable_hd_not_named: J-ESP, J-HRV, J-HUN, J-ITA, J-MLT, J-SMR, J-SRB
- POST_TREATMENT_RISK cannot_determine: J-ALB, J-AND, J-BGR, J-BIH, J-BLR, J-CZE, J-GRC, J-MDA, J-MKD, J-MNE, J-POL, J-PRT, J-ROU, J-SVK, J-SVN, J-UKR, J-VAT, J-XKX

## Per-jurisdiction summary

| Jurisdiction | Queries | Sources | Best tier | DSS | HD | Unresolved reason / note |
|---|---|---|---|---|---|---|
| J-ALB | 4 | 5 | S1 | unresolved | cannot_determine | generic grounds only; implementing list not located |
| J-AND | 4 | 4 | S5 | unresolved | cannot_determine | delegated disease list (regulation 2010) not retrievable |
| J-BGR | 4 | 6 | S1 | unresolved | cannot_determine | Ministry of Health criteria not retrieved |
| J-BIH | 4 | 2 | S1 | unresolved | cannot_determine | certificate content generic; no list |
| J-BLR | 4 | 1 | S3 | unresolved | cannot_determine | MoH disease list PDF timed out |
| J-CZE | 4 | 6 | S3 | unresolved | cannot_determine | current MoH measure not established |
| J-ESP | 5 | 2 | S2 | documented_absence | not_established | hours not tracked (time budget ~10-20 min/jurisdiction applied) |
| J-GRC | 4 | 5 | S1 | unresolved | cannot_determine | current exam content unspecified; historical TB tests |
| J-HRV | 3 | 7 | S1 | documented_positive | not_established | extended: imported leprosy case notice + new 2026 decision |
| J-HUN | 4 | 5 | S1 | documented_positive | not_established | extended: historical (114/2007, Act II 2007) vs current (Act XC 2023) conflict |
| J-ITA | 4 | 3 | S1 | documented_absence | not_established | hours not tracked (time budget ~10-20 min/jurisdiction applied) |
| J-MDA | 3 | 3 | S1 | documented_positive | cannot_determine | hours not tracked (time budget ~10-20 min/jurisdiction applied) |
| J-MKD | 2 | 1 | S5 | unresolved | cannot_determine | generic definition only |
| J-MLT | 2 | 6 | S1 | documented_positive | not_established | extended: Cap. 36 leprosy listing applicability check |
| J-MNE | 2 | 0 | S7 | unresolved | cannot_determine | official sources timed out |
| J-POL | 3 | 3 | S1 | unresolved | cannot_determine | delegated MoH disease-list regulation not retrievable |
| J-PRT | 2 | 3 | S1 | unresolved | cannot_determine | generic grounds; exam content unspecified |
| J-ROU | 2 | 1 | S3 | unresolved | cannot_determine | OUG 194/2002 & visa pages 403 |
| J-RUS | 4 | 9 | S1 | documented_positive | confirmed | extended: HD-specific evidence (official scanned PDF read page-by-page) |
| J-SMR | 2 | 4 | S1 | documented_absence | not_established | hours not tracked (time budget ~10-20 min/jurisdiction applied) |
| J-SRB | 3 | 2 | S1 | documented_absence | not_established | hours not tracked (time budget ~10-20 min/jurisdiction applied) |
| J-SVK | 3 | 3 | S1 | unresolved | cannot_determine | medical report content (decree 38/2011) not retrieved |
| J-SVN | 2 | 2 | S3 | unresolved | cannot_determine | ZTuj-2 text not rendered |
| J-UKR | 2 | 2 | S3 | unresolved | cannot_determine | no medical rule located; law text timed out |
| J-VAT | 2 | 2 | S1 | unresolved | cannot_determine | no labour-migration pathway identified |
| J-XKX | 2 | 1 | S1 | unresolved | cannot_determine | generic grounds; sub-legal acts not retrieved |

## Key adjudications
- **Russia (HD confirmed).** MoH Order 1079n of 19.11.2021 (in force 01.03.2022–01.03.2028), appendix 'Перечень инфекционных заболеваний, представляющих опасность для окружающих', item 2 'А30 лепра (болезнь Гансена)'. Official publication PDF (publication.pravo.gov.ru) is a scan without a text layer; page 9 was rendered (01_raw_search_logs/B7/derived/RUS_1079n_p-09.png) and read visually; identical wording verified in the Garant consolidated appendix text. Garant shows original item 5 (COVID-19) repealed from 01.03.2022 by Order 94n — not coded. 115-FZ art. 7(1)(13): TRP refused/annulled for listed diseases. No stage/treatment distinction stated → risk_status no_treatment_distinction_named_condition. Predecessor Order 384n (2015) list (pre-rebuild snapshot re-read) also named leprosy — coded as recovered_historical.
- **Croatia.** New MoH Decision NN 79/2026: visa-regime TCNs extending temporary residence must exclude active TB, S. typhi/paratyphi carriage, HIV, HBV, HCV; positives referred for treatment (no refusal consequence stated). Leprosy not listed. HZJZ notice of one imported leprosy case is epidemiological, not a migration rule and not a denial case.
- **Hungary.** 32/2007 EüM decree (in force) annex lists TB, HIV, lues, typhoid/paratyphoid carriage, hepatitis B; Act XC of 2023 §226(1)(c) refusal/withdrawal only if the person also does not submit to mandatory care (coded mentioned_conditional for active/on-treatment). Decree 114/2007 and Act II of 2007 time states are historical.
- **Malta.** IDCU 2026 rules: chest X-ray by TB-risk country of birth/residence; HBsAg for certain jobs; Cap. 217 art. 5(2)(b) mental disorder. Leprosy appears only in Cap. 36 art. 10(1) (domestic notifiable-disease list for arts. 5–9), not a migration criterion → HD not_established.
- **Moldova.** MFA page requires a certificate of absence of HIV/AIDS plus a 'standard medical certificate' whose content (MoH order on migrant examination) could not be retrieved → DSS positive, HD cannot_determine.
- **Czech Republic.** Decree 274/2004 (TB, syphilis) repealed 18.12.2015; current 'doklad o bezinfekčnosti' only if an MoH measure exists or on suspicion → unresolved.
- **Greece.** Law 5038/2023 art. 8 generic (WHO/EU diseases; exams non-systematic; post-entry illness protected). 2010/2012 circulars under repealed law 3386/2005 required chest X-ray + Mantoux — historical/current conflict unresolved.
- **Generic-only or delegated lists (unresolved):** Albania, Andorra, Bulgaria (MoH criteria), Bosnia and Herzegovina, North Macedonia, Portugal, Slovakia (§32(10) medical report), Poland (Act art. 36 delegates a refusal-of-entry disease list to MoH regulation Dz.U. 2015 poz. 1501 — not retrievable), Kosovo. Generic wording coded GENERIC_FITNESS only.
- **documented_absence (medium confidence):** Spain (fixed IHR-2005 certificate wording names no disease), Italy and San Marino (official document checklists contain no medical certificate), Serbia (MUP list: health insurance only).
- **Insufficient access (unresolved):** Belarus (MoH list PDF timed out), Montenegro (gov.me/wapi timed out), Romania (IGI/OUG 403), Slovenia (ZTuj-2 not rendered), Ukraine (law text timed out), Vatican (no labour-migration pathway identified).

## Time-budget extensions
- J-CZE: extended: repealed 274/2004 list vs current conditional measure
- J-GRC: extended: 2010/2012 circulars (old law) vs Law 5038/2023
- J-HRV: extended: imported leprosy case notice + new 2026 decision
- J-HUN: extended: historical (114/2007, Act II 2007) vs current (Act XC 2023) conflict
- J-MLT: extended: Cap. 36 leprosy listing applicability check
- J-RUS: extended: HD-specific evidence (official scanned PDF read page-by-page)

## Documents opened (saved)
See SOURCE_REGISTER.csv (local_path, URL, retrieved_utc, bytes, sha256). Pre-rebuild snapshots copied and re-read: MT_Cap36_pdf, RU_MoH384n_garant (original retrieval times kept).
HTTP 4xx/5xx error pages were not retained as sources (recorded in 01_raw_search_logs/B7/fetch_ledger.jsonl).

## Inaccessible sources (attempted, not saved)

| ISO | short | HTTP | error | URL |
|---|---|---|---|---|
| AND | LQI_art43 | 429 | HTTP error page not retained | https://regest.ad/ca/laws/llei_qualificada-llei-qualificada-d-immigracio-de-14-5-2002/articles/43 |
| BIH | PB7_EN | 404 | HTTP error page not retained | https://sps.gov.ba/dokumenti/dokumentacija/en/PB-7.pdf |
| BIH | LawAliens_EN | 404 | HTTP error page not retained | https://sps.gov.ba/dokumenti/zakoni/Law%20on%20Aliens.pdf |
| BIH | PB7_BS | 404 | HTTP error page not retained | https://sps.gov.ba/dokumenti/dokumentacija/PB-7.pdf |
| BIH | PB4_1 | 404 | HTTP error page not retained | https://sps.gov.ba/dokumenti/dokumentacija/PB-4.1.pdf |
| AND | RevisioMedicaInfo | 000 | curl: (28) Operation timed out after 90004 milliseconds with | https://www.govern.ad/documents/d/guest/t2fg002c-1 |
| AND | FAQ | 000 | curl: (28) Operation timed out after 90003 milliseconds with | https://www.govern.ad/ca/tematiques/immigracio/preguntes-frequents |
| BLR | epasluga_RVP | 404 | HTTP error page not retained | https://e-pasluga.by/en_US/services/citizenship-passport-id-card/200126issuance-of-a-temporary-residence-permi |
| ALB | Udhezim114_2024_natlex | 403 |  | https://natlex.ilo.org/dyn/natlex2/natlex2/files/download/118132/ALB-118132.pdf |
| BLR | MZ_Post2012_75 | 000 | curl: (28) Operation timed out after 90002 milliseconds with | https://minzdrav.gov.by/upload/lcfiles/%D0%BF%D0%BE%D1%81%D1%82%D0%B0%D0%BD%D0%BE%D0%B2%D0%BB%D0%B5%D0%BD%D0%B |
| AND | ConsentProves |  | TimeoutError('APIRequestContext.get: Timeout 90000ms exceede | https://www.govern.ad/documents/d/guest/t2g015 |
| AND | InfoRevisioMedica |  | TimeoutError('APIRequestContext.get: Timeout 90000ms exceede | https://www.govern.ad/documents/d/guest/t2fg002c |
| AND | AreaMedica |  | TimeoutError('Page.goto: Timeout 90000ms exceeded.\nCall log | https://www.govern.ad/ca/ministeris-i-secretaries-d-estat/ministeri-de-justicia-i-interior/departament-d-imigr |
| BIH | Pravilnik_BOS | 404 |  | https://sps.gov.ba/dokumenti/pravilnici/Pravilnik%20o%20ulasku%20i%20boravku%20stranaca%20BOS.pdf |
| ESP | DGGM_criterio5_2025_certmed | 403 |  | https://www.inclusion.gob.es/documents/d/migraciones/criterio-de-gestion-dggm-5-2025_rloex_certificado-medico- |
| GRC | MFA_KYA_eisodos2024 | 403 |  | https://www.mfa.gr/images/docs/ethnikes_theoriseis/KYA_eidika_dikaiologhtika_eisodoy_2024_B_1807-1.pdf |
| GRC | MFA_codification_EN | 403 |  | https://www.mfa.gr/media/sites/9/2026/01/codification_of_legislation_en.pdf |
| BIH | PB6 | 404 |  | https://sps.gov.ba/dokumenti/dokumentacija/PB-6.pdf |
| ESP | RD1155_2024_consol | 000 | curl: (28) Connection timed out after 90000 milliseconds | https://www.boe.es/buscar/pdf/2024/BOE-A-2024-24099-consolidado.pdf |
| BLR | MVD_PermRes |  | TimeoutError('Page.goto: Timeout 90000ms exceeded.\nCall log | https://www.mvd.gov.by/ru/page/departament-po-grazhdanstvu-i-migraci/migraciya/vydacha-razresheniya-na-postoya |
| MKD | LawForeigners2023_MK_refworld | 403 |  | https://www.refworld.org/sites/default/files/2024-05/5_law_on_foreigners_2023_consolidated_mkd.pdf |
| MKD | MTSP_dozvola_rabota_prestoj | 000 | curl: (6) Could not resolve host: cms.mtsp.gov.mk | https://cms.mtsp.gov.mk/dozvola-za-rabota-i-prestoj-na-stranec.nspx |
| POL | MOS_czasowy_praca_EN | 000 | curl: (60) SSL certificate problem: unable to get local issu | https://mos.cudzoziemcy.gov.pl/en/informacje/czasowy-praca/wymogi_EN |
| ESP | SEM_hoja55 | 000 | curl: (60) SSL certificate problem: unable to get local issu | https://inclusion.gob.es/web/migraciones/w/55.-modificacion-desde-situaciones-de-autorizaciones-de-residencia- |
| MDA | USMF_Ordin147_migranti | 000 | curl: (28) Connection timed out after 90005 milliseconds | https://usmf.md/sites/default/files/2020-01/147%20%28Ordin%20cu%20privire%20la%20organizarea%20examin%C4%83rii |
| MNE | GovME_temporary_residence | 000 | curl: (28) Connection timed out after 90000 milliseconds | https://www.gov.me/en/article/temporary-residence |
| MNE | Pravilnik_obrasci_wapi | 000 | curl: (28) Connection timed out after 90000 milliseconds | https://wapi.gov.me/download/d0cfc11c-d7b1-48c1-8e95-19e5a1fb117c?version=1.0 |
| MNE | ZakonStrancima_wapi | 000 | curl: (28) Connection timed out after 90000 milliseconds | https://wapi.gov.me/download/d361f791-fc73-42f5-826b-e6852c3f069a?version=1.0 |
| MNE | Biznis_dozvola_rad_boravak | 000 | curl: (28) Connection timed out after 90000 milliseconds | https://biznis.gov.me/vodici-i-informacije/boravak-za-strance-na-osnovu-zaposlenja-sezonskog-rada-ili-prekogra |
| ROU | MAE_NY_longstay_employment | 503 |  | https://newyork.mae.ro/en/node/486 |
| ROU | IGI_viza_angajare | 403 |  | https://igi.mai.gov.ro/viza-de-lunga-sedere-in-scop-de-angajare-detasare/ |
| ROU | OUG194_2002_IGI | 403 |  | https://igi.mai.gov.ro/wp-content/uploads/2026/04/ORDONANTA-DE-URGENTA-nr.-194-din-12-decembrie-2002.pdf |
| SRB | LawForeigners_EN_natlex | 403 |  | https://natlex.ilo.org/dyn/natlex2/natlex2/files/download/109269/SRB-109269%20(EN).pdf |
| RUS | MoH384n_minzdrav_attach | 404 |  | https://static-0.minzdrav.gov.ru/system/attachments/attaches/000/028/389/original/%D0%9F%D1%80%D0%B8%D0%BA%D0% |
| RUS | MoH1079n_gp180 | 000 | curl: (92) HTTP/2 stream 0 was not closed cleanly: INTERNAL_ | https://gp180.mos.ru/global_ruffe_tech/docs/1079nmz.pdf |
| SVK | EPI_509_2011 | 403 |  | https://www.epi.sk/zz/2011-509 |
| SRB | ZakonStrancima_MUP_lat | 000 | curl: (28) Operation timed out after 90002 milliseconds with | https://www.paragraf.rs/propisi_download/zakon_o_strancima.pdf |
| SRB | Pravilnik_privremeni_boravak_paragraf | 000 | curl: (28) Connection timed out after 90000 milliseconds | https://www.paragraf.rs/propisi/pravilnik-blizim-uslovima-odobrenje-privremenog-boravka.html |
| XKX | MPB_lejeqendrim | 000 | curl: (28) Operation timed out after 90001 milliseconds with | https://mpb.rks-gov.net/f/45/Leje-qendrim-i-perkohshem-per-te-huajt |
| UKR | Law3773_17_EN_rada | 000 | curl: (28) Connection timed out after 90001 milliseconds | https://zakon.rada.gov.ua/laws/show/en/3773-17 |
| ESP | RD1155_2024_eli |  | TimeoutError('Page.goto: Timeout 90000ms exceeded.\nCall log | https://www.boe.es/eli/es/rd/2024/11/19/1155/ |
| POL | DzU2010_484_MZ_wykaz | 000 | curl: (28) Connection timed out after 90001 milliseconds | https://www.dziennikustaw.gov.pl/D2010075048401.pdf |
| POL | DzU2015_1501_MZ_choroby | 000 | curl: (28) Connection timed out after 90001 milliseconds | https://www.dziennikustaw.gov.pl/D2015000150101.pdf |
| POL | Prawo_DzU2015_1501 | 404 |  | https://www.prawo.pl/akty/dz-u-2015-1501,18226880.html |

## Files
- 01_raw_search_logs/B7: raw/ (per-query hit lists), search_staging.jsonl, fetch_ledger.jsonl, text/ (text extractions), derived/ (renders, raw pdftotext), tools/ (fetch, coding, build scripts), WORKING_NOTES.md
- 02_primary_sources/B7: unmodified source files
- 03_extracted_evidence/B7: ten schema tables + this report

Validator: `python3 scoping-review-hansens-disease/pgph_rebuild/scripts/validate_batch.py B7`.
