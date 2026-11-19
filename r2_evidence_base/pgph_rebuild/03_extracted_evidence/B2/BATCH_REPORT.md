# Batch B2: Sub-Saharan Africa (East/South), batch report

Branch: `devin/1795078343-rebuild-B2`. Collection date: 2026-11-19 (UTC). Agent: Devin B2 batch agent.
All counts below come from the batch CSVs in this directory. Historical aggregates were not used as targets.

## 1. Coverage
- Jurisdictions investigated: 23 of 23. Each one has P1 English and P2 local/official-language searches. P3 applicability searches were run for BWA, KEN, MWI, MUS and NAM.
- Searches run: 96 (`SEARCH_LOG.csv`, Q-B2-0001 to Q-B2-0096). Raw result pages are in `01_raw_search_logs/B2/raw/`.
  - Language split: 57 in English, 39 in other languages: fr, pt, am, ti, ar, sw, so, rw, mg, ny, tn, st, ss, af, sn, bem, crs.
  - Stage split: P1 50, P2 41, P3 5.
  - Translation method: non-English queries were either composed by the agent in the language or machine-translated by the agent.
  - Q-B2-0081 and Q-B2-0084 were English queries mislabelled bem/sn in the raw headers. The language label is corrected in `SEARCH_LOG`.
- Saved files: 115 in `02_primary_sources/B2/`, each registered with SHA-256, byte count and UTC retrieval time.
  - 90 are HTTP-200 documents. The rest are saved access-error pages (tier S7).
  - 44 sources were relied on for adjudication.
  - Tier split of all registered files: S1 26, S2 3, S3 47, S4 4, S5 12, S7 23. S3/S5 for non-relied files is a host-based heuristic, stated in the notes.
- Evidence rows: 80. Of these, 38 name a condition explicitly, 5 are ambiguous (test named but not the disease, e.g. VDRL, chest X-ray) and 37 are generic or administrative.
- Policy pathways: 29.
- Search snippets were never used for classification. Every classification rests on full text that was opened and saved.

## 2. Endpoint results (jurisdiction level, work pathways P1/P2 only)
| Endpoint | Jurisdictions |
|---|---|
| DSS documented_positive (9) | BDI, ETH, MUS, MWI, NAM, SSD, TZA, ZAF, ZWE |
| DSS documented_absence (0) | none |
| DSS unresolved (14) | BWA, COM, DJI, ERI, KEN, LSO, MDG, MOZ, RWA, SOM, SWZ, SYC, UGA, ZMB |
| HD named: confirmed (3) | MUS, NAM, ZWE |
| HD named: not_established (2) | ETH, TZA |
| HD named: cannot_determine (18) | all others |
| Explicit post-treatment exclusion (0) | none |
| Any post-treatment provision (HD) | MUS (EDB Annex 9.1 note 4: retest after treatment); NAM (form asks lifetime history "ever suffered"; no consequence stated) |
| Treatment distinction (exempt/conditional) | PW-MUS-P2 |
| No treatment distinction, named condition | PW-MUS-P1, PW-NAM-P1, PW-ZWE-P1 |
| Observed denials (case-level) | none located (`OBSERVED_DENIALS.csv` contains the header only) |

SYC and ZMB have a documented_positive pathway only for residence permits that are not specific to workers (PW-SYC-P3, PW-ZMB-P3, MENTAL). Their jurisdiction DSS stays unresolved.

### Explicitly named conditions in current, applicable S1–S4 evidence (jurisdictions; any pathway)
| Condition | Jurisdictions |
|---|---|
| HD | MUS, NAM, ZWE |
| TB, named | NAM, ZAF, ZWE |
| TB, chest X-ray only (ambiguous) | MUS |
| HIV | MUS, NAM |
| SYPH, named | NAM, ZWE |
| SYPH, VDRL/TPHA only (ambiguous) | MUS |
| HBV | MUS |
| FILARIASIS | MUS |
| MENTAL | BDI, MUS, MWI, NAM, SSD, SYC (P3), TZA, ZAF, ZMB (P3), ZWE |
| SUBST_DEP | ETH, MUS |
| ALCOHOL | MUS |
| GENERIC_FITNESS / administrative only | all 23 have at least one generic or administrative row |

## 3. Key adjudications
- **MUS: HD confirmed.** Two current official instructions name leprosy:
  - Ministry of Health and Wellness, "Requirements for Medical Clearance for Migrant Workers" (OHU, Nov 2025, for 2026). Nationals of 15 listed countries (Brazil, India, Indonesia, Bangladesh, DRC, Ethiopia, Madagascar, Mozambique, Myanmar, Nepal, Nigeria, Philippines, Somalia, Sri Lanka, Tanzania) must be "certified by a doctor as free of leprosy based on examination", before arrival and again on arrival. This is a nationality trigger.
  - EDB Occupation/Residence Permit guidelines, Annex 9/9.1. The leprosy test is "restricted to Indian nationals only". Note 4 says that applicants who tested positive for leprosy and then "undergone the appropriate treatment" must undergo new tests. This is the only explicit post-treatment rule in B2. It is conditional, not an exclusion, and the outcome after the retest is not stated.
  - The 2016 Ministry of Labour version (India only) is recorded as historical.
  - Immigration Act s.8(1)(b) is generic. s.8(2) lets the Minister authorise admission (waiver).
- **NAM: HD confirmed, medium confidence.**
  - Current Immigration Regulations (GN 134/1994, LAC annotated consolidation to 2026), Annexure M "Application for an Employment Permit", asks whether the applicant has "ever suffered, or is at present suffering from ... (v) leprosy". The same question appears in Annexures B, F, G and R.
  - The operative prohibited-immigrant list (reg. 13) listed leprosy in its original 1994 text. GN 133/2010 replaced it with a PHEIC-based generic definition, so leprosy is named only as a disclosure item, with no stated consequence. This is coded as no treatment distinction (lifetime history).
- **ZWE: HD confirmed.** Immigration Regulations 1998 (SI 195/1998, as amended by SI 126/2005), reg. 44(g), lists leprosy among the prescribed diseases that make a person a prohibited person unless they hold a permit with reg. 39 conditions. TB and syphilis are also listed. Reg. 7 requires a radiologist certificate showing no active pulmonary TB.
  - Numbering caveat: the regulations refer to s.14(1)(e), but the official Act PDF has the prescribed-disease ground at s.14(1)(d).
  - Amendments after 2005 were not checked.
- **ZAF: TB named; HD cannot_determine.**
  - The radiological report is defined as showing no active pulmonary TB, and Form DHA-1738 asks about TB.
  - s.29(1)(a) and reg. 26(1) delegate the disease list to the International Health Regulations Act regulations and Department of Health determinations. These were not obtained.
  - Ministerial Directive GN 7714 (17 July 2026) waives the medical report for temporary visas and permanent residence. The radiological report is not mentioned in the waiver.
- **ETH, TZA: HD not_established.**
  - ETH: visa denial grounds (Proclamation 354/2003, Art. 5) were read in full. They name drug addiction and an undefined "dangerous contagious disease", with no delegated list.
  - TZA: Immigration Act 1995 s.10 names mental disorder and uses an individually certified contagious-disease ground.
- **Prescribed or delegated lists not obtained, so HD cannot_determine:** BWA, MWI, SSD, SYC, ZMB, BDI (OCR source), KEN (Regulations LN 64/2012 not retrieved).
- **Generic or administrative only, so DSS unresolved (never absence):** COM, DJI, ERI, LSO, MDG, MOZ, RWA, SOM, SWZ, UGA. A "medical certificate/report" requirement whose content is not specified is not a complete requirement and was not coded as documented_absence.
- **Historical sources recorded but not used for current status:**
  - Botswana Immigration Act 1966 (Refworld file mislabelled "2011").
  - Zambia Cap. 123, which prescribed TB and smallpox; repealed by Act 18/2010.
  - Namibia original reg. 13.
  - Mauritius 2016 work-permit guidelines.
- **Safeguards:** a safeguard is marked explicit_available only where text states it. Missing statements are no_statement_located. Appeal/review provisions were not searched for specifically.

## 4. Budget extensions (addendum conditions; also recorded in JURISDICTION_STATUS.notes)
- MUS, NAM, ZWE: plausible Hansen-specific evidence.
- NAM, ZAF: conflict between historical and current rules.

All other jurisdictions stayed within the initial budget and are left unresolved where no evidence was found. Time was not tracked per jurisdiction because the work was parallelised, so `hours_spent` is blank.

## 5. Inaccessible or not obtained (critical for unresolved items)
- Botswana: gov.bw permit pages (HTTP 500). Prescribed-disease regulations not located.
- Djibouti: egouv.dj (DNS failure).
- Eswatini:
  - gov.sz Immigration Act 1982 PDF and medical.pdf: timeouts.
  - eswatinilii Immigration Act 1982: HTTP 403, including through the browser.
- Kenya:
  - immigration.go.ke prohibited-immigrants page: HTTP 500.
  - Kenya Citizenship and Immigration Regulations 2012 (LN 64/2012): kenyalaw returned HTTP 202 with no content. The NatLex "2012" file contains the Act, not the regulations.
- Namibia:
  - mha.gov.na / mhaiss.gov.na: HTTP 468.
  - Immigration Control Act 7 of 1993 itself (LAC): timeout. The s.39(2)(e) wording was taken only as cross-referenced in the regulations.
- Seychelles: egov.sc Gainful Occupation Permit application form (DNS failure).
- Somalia: molsa.gov.so (HTTP 403); immigration.gov.so law PDF (HTTP 404).
- Ethiopia: work-permit directive page (justice.gov.et) returned a 90-byte placeholder.
- Madagascar: Loi 62-006 not retrieved.
- Zambia: zambialii 2010 Act (HTTP 403). Used the NatLex copy via the browser instead. Prescribed-disease regulations under the 2010 Act not located.
- Delegated lists not obtained: Malawi, South Sudan, Seychelles and Zambia prescribed-disease lists; South Africa IHR Act regulations.

## 6. Data integrity and verification
- `validate_batch.py B2`: PASS. This checks headers, coverage, SHA-256 for every local file, enumerations and source references.
- SHA-256 and byte counts were also checked against the files when the register was built.
- Automated quote verification: all 80 `source_wording_original` strings were checked against text extracted from the saved files.
  - 71 match exactly after whitespace normalisation.
  - 7 match token by token in order (multi-column PDF layout, page headers, form boxes).
  - 2 (E-BDI-003, E-BDI-004) come from OCR of a scanned gazette with OCR artefacts corrected. The corrections are listed in `classification_note`, and the passages were checked manually against the OCR output.
- English translations of French and Portuguese passages are the agent's own translations.
- **Independent verification is NOT done.** All rows are `provenance_class=new_collected`. A second, independent reviewer pass is still required by the Gate 2/3 authorization before these rows can be promoted to `independently_verified`.
- Checkpointing: tables were built in one pass after collection rather than every ~5 jurisdictions. Raw logs and sources were persisted to the repository directories throughout.
