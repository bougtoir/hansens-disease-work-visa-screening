# B9 batch report (South/North America + Oceania, 28 jurisdictions)

Generated from the B9 tables by the agent; all classifications trace to rows in EVIDENCE_ROWS.csv / POLICY_PATHWAYS.csv. Historical aggregate counts were not used.

## Searches run
- 66 web searches (P1_english and P2_local for every jurisdiction; raw result pages in `01_raw_search_logs/B9/web_search_raw/`, ledger `QUERY_LEDGER.csv`).
- P2 for English-official Pacific/Guyana jurisdictions used official legal databases/portals in English; vernacular-language searches (Samoan, Tongan, Tok Pisin, etc.) were not run (recorded per jurisdiction in JURISDICTION_STATUS.notes). Spanish, Portuguese, French and Dutch queries were written and read directly by the agent (no machine translation).
- No search-budget extension beyond the initial allocation was used.

## Documents opened
- 102 fetch attempts recorded in SOURCE_REGISTER.csv; 71 saved and read (raw files in `02_primary_sources/B9/`, SHA-256 + UTC in register); 31 inaccessible/unusable (tier S7).
- Tool incident: one parallel fetch run shared a temp file; the one file whose content could not be verified was discarded (register row marked); the PIC 4007 file from that run was checked and kept. Fixed with per-process temp files.

## Verified vs unresolved
- DSS documented_positive: J-AUS, J-CAN, J-GUY, J-NZL, J-USA, J-VUT, J-WSM
- DSS documented_absence: J-ARG, J-COL, J-PER
- DSS unresolved: J-BOL, J-BRA, J-CHL, J-ECU, J-FJI, J-FSM, J-KIR, J-MHL, J-NRU, J-PLW, J-PNG, J-PRY, J-SLB, J-SUR, J-TON, J-TUV, J-URY, J-VEN
- HD confirmed: J-CAN, J-GUY, J-USA, J-WSM
- HD not_established: J-ARG, J-AUS, J-COL, J-NZL, J-PER, J-VUT
- HD cannot_determine: J-BOL, J-BRA, J-CHL, J-ECU, J-FJI, J-FSM, J-KIR, J-MHL, J-NRU, J-PLW, J-PNG, J-PRY, J-SLB, J-SUR, J-TON, J-TUV, J-URY, J-VEN
- Evidence rows: 46; pathways: 31; observed denials: 0 (no case-level evidence located).

## Key adjudications
- USA: 42 CFR 34.2(b) lists "Hansen's disease, infectious"; CDC Technical Instructions: untreated MB/PB HD = Class A, after 7 days of treatment may be reclassified Class B -> treatment distinction (conditional). Waiver (INA 212(g)) and 42 CFR 34.8 reexamination recorded. Temporary work (NIV) exam is not routine (22 CFR 41.108).
- Canada: IRCC Panel Members' Guide physical-exam row "Screen for skin cancer, leprosy, ..." -> HD named (screening item); no HD-specific consequence or treatment rule located -> post-treatment cannot_determine. Syphilis screening mandatory 15+.
- Samoa: MPMC100 Temporary Residence application Q16 asks about "communicable illness or disease (eg HIV, T.B., Leprosy, STD)" -> HD named in a self-declaration; no consequence stated -> post-treatment cannot_determine.
- Guyana: Immigration Act Cap. 14:02 prohibited immigrant "is suffering from any variety of leprosy ..." -> HD named; no treatment-state wording -> no_treatment_distinction_named_condition (not an explicit post-treatment exclusion; 'suffering from' does not establish exclusion of cured persons). Text is the Laws of Guyana revised edition hosted on sela.org (official mola.gov.gy copy not retrieved).
- Australia (PIC 4005 'free from tuberculosis'), New Zealand (INZ1007 compulsory syphilis/HBV/HCV/HIV tests), Vanuatu (Medical Report HIV/HBV/VDRL) -> DSS positive; HD not named in the sources read -> not_established (not a claim of absence in all instruments).
- Generic clauses ("contagious or infectious disease", "serious communicable disease", "social disease", Carné de Salud fitness) were kept as GENERIC_FITNESS and did not establish DSS (PNG, FJI, FSM, PLW, URY, WSM Act).
- Chest X-ray fields were not coded as TB screening. Behavioural grounds (narcotic addiction, chronic alcoholism) were recorded but not used for DSS status.
- Documented absence (ARG, COL, PER) is based on the full official requirement list for the work pathway plus the governing statute's impediment article naming no condition; confidence medium.

## Inaccessible sources (critical)
- J-USA CDC_TI_civil_hansens: https://www.cdc.gov/immigrant-refugee-health/hcp/civil-surgeons/hansens-disease.html (fetch failed: HTTP 403 | fetch mode curl, content_type text/html)
- J-USA CDC_TI_panel_hansens: https://www.cdc.gov/immigrant-refugee-health/hcp/panel-physicians/hansens-disease.html (fetch failed: HTTP 403 | fetch mode curl, content_type text/html)
- J-USA USCIS_PM_v8pB_ch6: https://www.uscis.gov/policy-manual/volume-8-part-b-chapter-6 (fetch failed: HTTP 403 | fetch mode curl, content_type text/html)
- J-USA USCIS_PM_v8pB_ch3: https://www.uscis.gov/policy-manual/volume-8-part-b-chapter-3 (fetch failed: HTTP 403 | fetch mode curl, content_type text/html)
- J-CAN IME_tuberculosis: https://www.canada.ca/content/dam/ircc/migration/ircc/english/department/partner/pp/pdf/imei_tuberculosis.pdf (fetch failed: HTTP 000 | fetch mode curl, content_type)
- J-CAN IME_syphilis: https://www.canada.ca/content/dam/ircc/migration/ircc/english/department/partner/pp/pdf/imei_syphilis.pdf (fetch failed: HTTP 000 | fetch mode curl, content_type)
- J-AUS PanelMemberInstructions: https://immi.homeaffairs.gov.au/support-subsite/files/panel-member-instructions.pdf (fetch failed: HTTP 200 | fetch mode curl, content_type application/pdf)
- J-PNG ICA_visa_info_guide: https://ica.gov.pg/uploads/media/post_file_5227574-visa-information-guide-online.pdf (fetch failed: HTTP 200 file discarded by agent: saved during temp-file collision between parallel fetches (tool bug), co)
- J-PNG ICA_medical_exam_form: https://ica.gov.pg/uploads/media/post_file_2390902-medical-examination-form.pdf (URL of ICA medical examination form returned an HTML page, not the form -> form content inaccessible | fetch mode curl, )
- J-NZL OIA_MedAssessorGuidelines: https://fyi.org.nz/request/29782/response/118106/attach/4/DOIA%20REQ%200008226%20Appendix%201.pdf (fetch failed: HTTP 200 | fetch mode curl, content_type application/pdf; charset=utf-8)
- J-PNG PacLII_MigrationAct1978: https://www.paclii.org/pg/legis/consol_act/ma1978118/ (fetch failed: HTTP 403 | fetch mode browser, content_type text/html(rendered DOM))
- J-PNG PacLII_MigrationReg1979: https://www.paclii.org/pg/legis/consol_act/mr1979209/ (fetch failed: HTTP 403 | fetch mode browser, content_type text/html(rendered DOM))
- J-PNG ICA_medical_exam_form: https://ica.gov.pg/uploads/media/post_file_2390902-medical-examination-form.pdf (URL of ICA medical examination form returned an HTML page, not the form -> form content inaccessible | fetch mode curl, )
- J-PNG WB_PacLII_MigrationAct1978: http://archive.org/wayback/available?url=paclii.org/pg/legis/consol_act/ma1978118/ (Wayback availability-API JSON (locator only, not a legal text) | fetch mode curl, content_type application/json)
- J-KIR natlex_113881: https://natlex.ilo.org/dyn/natlex2/natlex2/files/download/113881/KIR113881.pdf (fetch failed: HTTP 403 | fetch mode curl, content_type text/html; charset=UTF-8)
- J-MHL natlex_74131: https://natlex.ilo.org/dyn/natlex2/natlex2/files/download/74131/MHL74131.pdf (fetch failed: HTTP 403 | fetch mode curl, content_type text/html; charset=UTF-8)
- J-FSM RevImmigReg: https://fsmembassy.fm/wp-content/uploads/2022/05/RevFSMImmigrationReg.pdf (scanned PDF without text layer; not read | fetch mode curl, content_type application/pdf)
- J-SLB natlex_71918: https://natlex.ilo.org/dyn/natlex2/natlex2/files/download/71918/SLB71918.pdf (fetch failed: HTTP 403 | fetch mode curl, content_type text/html; charset=UTF-8)
- J-PLW DOL_Rules2019: https://ministryofhrctd.pw/wp-content/uploads/2022/03/Division-of-Labor-Rules-Regulation-2019.pdf (fetch failed: HTTP 403 | fetch mode curl, content_type text/html)
- J-MHL rmiparl_1970_0049: https://rmiparliament.org/cms/images/LEGISLATION/PRINCIPAL/1970/1970-0049/1970-0049_1.pdf (fetch failed: HTTP 502 | fetch mode curl, content_type text/html)
- J-TON ImmigRegs: https://ago.gov.to/cms/images/LEGISLATION/SUBORDINATE/2502/2502-0001/ImmigrationRegulations_3.pdf (fetch failed: HTTP 000 | fetch mode curl, content_type)
- J-TUV leg_1968_0010: https://www.tuvalu-legislation.tv/cms/images/LEGISLATION/PRINCIPAL/1968/1968-0010/1968-0010_2.pdf (fetch failed: HTTP 502 | fetch mode curl, content_type text/html)
- J-TUV leg_2014_0003: https://www.tuvalu-legislation.tv/cms/images/LEGISLATION/SUBORDINATE/2014/2014-0003/2014-0003_2.pdf (fetch failed: HTTP 502 | fetch mode curl, content_type text/html)
- J-BRA Lei13445: https://www.planalto.gov.br/ccivil_03/_ato2015-2018/2017/lei/l13445.htm (fetch failed: HTTP 429 | fetch mode curl, content_type text/html)
- J-BRA Decreto9199: https://www.planalto.gov.br/ccivil_03/_ato2015-2018/2017/decreto/d9199.htm (fetch failed: HTTP 429 | fetch mode curl, content_type text/html)
- J-VEN Ley37944_oas: http://www.oas.org/dil/Migrants/Venezuela/Ley%20N%C2%B0%2037.944%20de%20migraci%C3%B3n%20y%20extranjer%C3%ADa%20del%201%20de%20julio%20de%202004.pdf (fetch failed: HTTP 503 | fetch mode curl, content_type text/html; charset=utf-8)
- J-VEN SAIME_visa_TL: https://info.saime.gob.ve/extranjeria/visa/visa_tl_tf (fetch failed: HTTP 000 | fetch mode curl, content_type)
- J-BOL DS1923_Reglamento: https://migracion.gob.bo/sites/default/files/2021-06/REGLAMENTACI%C3%93N%20DE%20LA%20LEY%20N%C2%B0370%20DE%20MIGRACI%C3%93N%20DE%2012%20DE%20MARZO%20DE%202014.pdf (fetch failed: HTTP 000 | fetch mode curl, content_type)
- J-BRA Lei13445_b: https://www.planalto.gov.br/ccivil_03/_ato2015-2018/2017/lei/l13445.htm (browser capture returned Planalto portal home page (not the law text) -> law text inaccessible | fetch mode browser, con)
- J-BRA Decreto9199_b: https://www.planalto.gov.br/ccivil_03/_ato2015-2018/2017/decreto/d9199.htm (browser capture returned Planalto portal home page (not the decree text) -> inaccessible | fetch mode browser, content_t)
- J-VEN SAIME_visa_TL_b: https://info.saime.gob.ve/extranjeria/visa/visa_tl_tf (fetch failed: HTTP none Error('Page.goto: net::ERR_NAME_NOT_RESOLVED at https://info.saime.gob.ve/extranjeria/visa/visa_)

## Process notes
- Checkpoint commits were made by region after collection rather than strictly every five jurisdictions.
- hours_spent was not tracked per jurisdiction (left blank).
