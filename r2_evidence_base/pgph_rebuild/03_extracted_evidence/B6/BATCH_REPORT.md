# B6 Europe (North/West) - batch report

Evidence-first rebuild, batch B6 (19 jurisdictions). Single extraction pass by one agent; independent verification has not been performed (provenance_class = new_collected). Historical labels/counts were not used.

## Method

- Searches: 81 queries (DuckDuckGo via Chrome CDP), each logged with up to 8 result rows in SEARCH_LOG.csv (648 rows). Stages: P1 English and P2 official/local language for every jurisdiction, plus a P2 Hansen-term query (lepra / lèpre / spedalskhed / spedalskhet / holdsveiki / leepra / raupsai) per jurisdiction. No HD-term query returned an immigration-related official source.
- Documents fetched and saved: 78 files in 02_primary_sources/B6/ (3 early challenge pages moved to 01_raw_search_logs/B6/failed_fetches/). SHA-256 and UTC retrieval time in SOURCE_REGISTER.csv. Tier counts: S1=13, S2=4, S3=19, S4=4, S5=5, S7=33. S7 rows are failed captures (403, bot/captcha challenge, empty JS shell) and were not used as evidence.
- Search-snippet text in SEARCH_LOG.relevant_excerpt is a lead only; every classification below rests on full text read from a saved file.
- Internet Archive (Wayback) copies were used as S4 where the live official site blocked automated access (AUT RIS NAG s11; FIN Migri; CHE EDA; MCO service page).
- Time budget: the initial 10-20 min per jurisdiction budget was not extended for any jurisdiction (no HD-specific lead found). JURISDICTION_STATUS.hours_spent is the elapsed span of logged activity and is approximate because work was interleaved.
- Checkpoint commits: all jurisdictions were collected before the first commit, so checkpoint commits were made by stage (raw logs and sources first, then extracted tables) rather than every ~5 jurisdictions.

## Results by jurisdiction

| Jurisdiction | DSS | HD_NAMED | Best tier | Named conditions (S1-S4) | Basis / reason unresolved |
|---|---|---|---|---|---|
| J-AUT | unresolved | cannot_determine | S3 | - | Official (S4) NAG s11(7) requires a health certificate only where FPG s23 applies; the named categories (incl. TB) were read only on S5 reproductions; designating ordinance not located |
| J-BEL | documented_positive | not_established | S1 | OTHER_NAMED, TB | Annex point 2 names active/progressive respiratory TB (S1); certificate model (S2) |
| J-CHE | unresolved | cannot_determine | S3 | - | No medical requirement text located; no source states that no medical exam is required; fedlex not opened |
| J-DEU | unresolved | cannot_determine | S1 | - | AufenthG s5 general conditions contain no health condition; no source states no exam is required; IfSG s36(4) concerns admission to institutions, not work migration |
| J-DNK | unresolved | cannot_determine | S7 | - | nyidanmark.dk captures failed (JS); retsinformation not opened |
| J-EST | unresolved | cannot_determine | S1 | - | Only generic "health" requirement located; no named condition and no complete medical requirement |
| J-FIN | unresolved | cannot_determine | S4 | - | Archived Migri requirement list contains no medical requirement but does not state that no exam is required; THL screening guidance capture failed |
| J-FRA | unresolved | cannot_determine | S3 | - | S3 page lists exam components without naming a disease; disease-specific criteria (TB, mental disorders) read only on S5 reproduction of arrete 11/01/2006 |
| J-GBR | documented_positive | not_established | S1 | OTHER_NAMED, TB | Appendix TB (S1) names active pulmonary TB |
| J-IRL | documented_positive | not_established | S1 | MENTAL, OTHER_NAMED, SUBST_DEP, SYPH, TB | First Schedule (S1) names respiratory TB, syphilis, drug addiction, profound mental disturbance |
| J-ISL | documented_positive | not_established | S2 | HBV, HCV, HIV, OTHER_NAMED, SYPH, TB | Directorate of Health procedures (S2) require TB, HBV, HCV, HIV, syphilis testing |
| J-LIE | unresolved | cannot_determine | S1 | - | ZAV keyword search found no health requirement; AuG capture failed |
| J-LTU | unresolved | cannot_determine | S1 | - | Public-health refusal ground and any implementing disease list not located in captured text |
| J-LUX | documented_positive | cannot_determine | S3 | TB | Guichet.lu (S3) names tuberculosis screening |
| J-LVA | documented_positive | not_established | S1 | TB | MK noteikumi 825 p.2 (S1) names active-phase TB |
| J-MCO | unresolved | cannot_determine | S1 | - | OS 3.153 carte de sejour provisions located without medical requirement; official service page blocked; employer medical (visite d embauche) not checked |
| J-NLD | documented_positive | not_established | S2 | TB | IND (S3) and form 7603 (S2) name TB test |
| J-NOR | documented_positive | not_established | S1 | TB | Forskrift om tuberkulosekontroll (S1) names TB |
| J-SWE | unresolved | cannot_determine | S1 | - | No work-permit medical requirement located; no source states that no exam is required |

Counts: DSS documented_positive=8, unresolved=11; HD_NAMED cannot_determine=12, not_established=7. HD_NAMED confirmed = 0. Observed denials recorded = 0 (no case-level evidence located; none inferred).

## Key adjudications

- BEL: Annex to Loi 15/12/1980 (Justel, S1) and the official model certificate (S2) name active/progressive respiratory TB; IHR quarantinable diseases are a category by reference; item 3 ("autres maladies infectieuses ou parasitaires contagieuses ...") coded GENERIC_FITNESS, not HD. Single-permit medical certificate requirement in Art. 61/25-2 with a ministerial waiver for duly justified impossibility.
- IRL: Immigration Act 2004 First Schedule (S1, as-enacted text) names respiratory TB, syphilis, drug addiction (SUBST_DEP) and profound mental disturbance (MENTAL); item 4 generic. Applies to permission to land, coded P3_other. Current consolidated status not checked.
- GBR: Appendix TB (S1) - active pulmonary TB certificate for >6-month entry clearance after >=6 months in a TB6 country; TB5 waiver. SUI 19 guidance (S2) gives SARS/Covid-19 and Ebola as examples; no leprosy.
- LVA: MK noteikumi Nr. 825 (S1) - the only refusal disease is active-phase TB.
- ISL: Directorate of Health procedures (S2) - TB, HBV, HCV, HIV, syphilis, enteric infections; origin-region trigger. Tests only; no diagnosis-based refusal stated.
- NLD, NOR, LUX: TB test/screening obligations (S1-S3). LUX HD_NAMED = cannot_determine because the content and legal basis of the general medical examination were not opened.
- AUT: NAG s11(7) (S4 archived RIS) requires a health certificate only where FPG s23 would apply. The disease categories (incl. notifiable TB) and the Vorarlberg TB screening rule were read only on unofficial portals (S5), and the designating ordinance was not located, so DSS = unresolved and HD = cannot_determine. S5 rows are kept as leads.
- FRA: the official OFII page (S3) lists a chest X-ray "depending on your country of origin" and "screening for infectious diseases" without naming a disease (chest X-ray not coded as TB). The arrete of 11/01/2006 (TB, mental disorders) was read only on an unofficial reproduction (S5); Legifrance was blocked. DSS = unresolved.
- EST: only a generic "tervis"/"health" requirement (GENERIC_FITNESS); unresolved.
- CHE, DEU, DNK, FIN, LIE, LTU, MCO, SWE: no medical requirement located, and no S1-S4 source states that no medical exam is required. These are unresolved, not documented_absence.

## Post-treatment and clinical state

HD is not named in any located S1-S4 source, so CLINICAL_STATUS_RULES (HD columns) are not_addressed everywhere. POST_TREATMENT_RISK: not_applicable_hd_not_named for J-BEL, J-GBR, J-IRL, J-ISL, J-LVA, J-NLD, J-NOR; cannot_determine for J-AUT, J-FRA, J-LUX, J-EST, J-CHE, J-DEU, J-DNK, J-FIN, J-LIE, J-LTU, J-MCO, J-SWE. No explicit post-treatment exclusion was found. TB-specific clinical wording (e.g., BEL "active ou à tendance évolutive", LVA "aktīvā fāzē", GBR "active pulmonary") is recorded in EVIDENCE_ROWS.clinical_state_rule.

## Inaccessible or blocked sources

- S-BEL-001 J-BEL: https://dofi.ibz.be/sites/default/files/2021-04/Certificat_m%C3%A9dical.pdf (http=200)
- S-BEL-002 J-BEL: https://dofi.ibz.be/fr/themes/faq/long-stay/certificat-medical (http=200)
- S-BEL-003 J-BEL: https://www.ejustice.just.fgov.be/eli/loi/1980/12/15/1980121550/justel (http=200)
- S-BEL-005 J-BEL: https://dofi.ibz.be/sites/default/files/2021-04/Certificat_m%C3%A9dical.pdf (http=chrome)
- S-BEL-006 J-BEL: https://dofi.ibz.be/fr/themes/faq/long-stay/certificat-medical (http=chrome)
- S-AUT-004 J-AUT: https://www.ris.bka.gv.at/GeltendeFassung.wxe?Abfrage=Bundesnormen&Gesetzesnummer=20004243 (http=chrome)
- S-FRA-002 J-FRA: https://www.legifrance.gouv.fr/loda/id/LEGITEXT000006053178/ (http=chrome)
- S-FRA-003 J-FRA: https://etrangers-en-france.interieur.gouv.fr/vivre-en-france/la-visite-medicale-obligatoire-a-l-ofii (http=chrome)
- S-FIN-001 J-FIN: https://migri.fi/en/residence-permit-for-an-employed-person (http=chrome)
- S-ISL-003 J-ISL: https://www.althingi.is/altext/145/s/1467.html (http=403)
- S-LIE-001 J-LIE: https://www.gesetze.li/konso/2008.311 (http=200)
- S-LIE-002 J-LIE: https://www.gesetze.li/konso/2008.350 (http=200)
- S-CHE-001 J-CHE: https://www.eda.admin.ch/countries/usa/en/home/visa/entry-ch/work-permits.html (http=403)
- S-MCO-001 J-MCO: https://monservicepublic.gouv.mc/en/themes/nationality-and-residency/residency/new-entrants/how-to-apply-for-a-residence-permit (http=403)
- S-MCO-002 J-MCO: https://monservicepublic.gouv.mc/thematiques/nationalite-et-residence/residence/nouveaux-arrivants/demander-une-carte-de-sejour (http=403)
- S-FIN-002 J-FIN: https://migri.fi/en/residence-permit-for-an-employed-person (http=chrome)
- S-FRA-004 J-FRA: https://www.legifrance.gouv.fr/loda/id/JORFTEXT000000241240 (http=chrome)
- S-LIE-003 J-LIE: https://www.gesetze.li/konso/2008350000 (http=chrome)
- S-LIE-004 J-LIE: https://www.llv.li/serviceportal2/amtsstellen/auslaender-und-passamt/checkliste-verlaengerung-des-aufenthalts-b-c-d-.pdf (http=403)
- S-EST-002 J-EST: https://www.riigiteataja.ee/akt/108072025009 (http=200)
- S-FIN-003 J-FIN: https://thl.fi/ohjeita-ammattilaisille/maahanmuuttaneiden-terveyspalvelut-ja-tartuntatautien-ehkaisy/maahanmuuttaneet-ja-tartuntatautien-seulonta (http=403)
- S-DNK-001 J-DNK: https://www.nyidanmark.dk/en-GB/Applying/Work (http=chrome)
- S-FIN-004 J-FIN: https://thl.fi/ohjeita-ammattilaisille/maahanmuuttaneiden-terveyspalvelut-ja-tartuntatautien-ehkaisy/maahanmuuttaneet-ja-tartuntatautien-seulonta (http=chrome)
- S-FRA-005 J-FRA: https://www.legifrance.gouv.fr/loda/id/JORFTEXT000000241240 (http=403)
- S-DNK-002 J-DNK: https://www.nyidanmark.dk/en-GB/Applying/Work (http=chrome)
- S-AUT-008 J-AUT: https://www.ris.bka.gv.at/Dokumente/Bundesnormen/NOR40067827/NOR40067827.pdf (http=503)
- S-FRA-006 J-FRA: https://web.archive.org/web/2025/https://www.legifrance.gouv.fr/loda/id/JORFTEXT000000241240 (http=404)
- S-DNK-003 J-DNK: https://web.archive.org/web/2025/https://www.nyidanmark.dk/en-GB/Applying/Work (http=200)
- S-AUT-010 J-AUT: https://web.archive.org/web/2025/https://www.ris.bka.gv.at/NormDokument.wxe?Abfrage=Bundesnormen&Gesetzesnummer=20004241&Paragraf=23 (http=404)
- S-FRA-008 J-FRA: https://web.archive.org/web/2025/https://etrangers-en-france.interieur.gouv.fr/vivre-en-france/la-visite-medicale-obligatoire-a-l-ofii (http=404)
- S-AUT-011 J-AUT: https://web.archive.org/web/2024id_/https://www.ris.bka.gv.at/NormDokument.wxe?Abfrage=Bundesnormen&Gesetzesnummer=20004241&Paragraf=23 (http=404)
- S-FRA-009 J-FRA: https://web.archive.org/web/2024id_/https://www.legifrance.gouv.fr/loda/id/JORFTEXT000000241240/ (http=404)
- S-FRA-010 J-FRA: https://web.archive.org/web/2023id_/https://www.legifrance.gouv.fr/jorf/id/JORFTEXT000000241240 (http=404)

## Files

- 01_raw_search_logs/B6/: queries.jsonl (each query with its results), fetches.jsonl (fetch ledger), saved search-result HTML pages, failed_fetches/, tools/ (scripts and query lists used).
- 02_primary_sources/B6/: raw snapshots, unmodified.
- 03_extracted_evidence/B6/: nine tables plus JURISDICTION_STATUS.csv, this report, and B6_bundle.tar.gz.
