source("DZA_work_residence","Work and residence in Algeria","Algerian Ministry of Foreign Affairs","official portal page",notes="Available page does not expose a complete permit checklist.")
source("DZA_gazette_1982","Journal officiel de la République algérienne, No. 56 (1982)","Algerian Official Gazette","gazette issue",notes="Image-only issue examined as a law lead. Not used to infer the current administrative requirements.")
source("DZA_work_rendered","Guide de procédures","Algerian Ministry of Labour, Employment and Social Security","partially rendered official work instructions",
       notes="Retrieved with web_get_contents. Output was truncated in the provisional permit section. Cannot be used to establish completeness of permit criteria.")
source("DZA_resident_rendered","Délivrance de la carte de résident","Algerian Ministry of the Interior","official rendered residence guide",
       notes="Retrieved with web_get_contents. Applies general medical documentation alongside provisions for employed applicants.")
path("DZA","P1","Employment by a foreign national",U,C,"DZA_work_residence DZA_work_rendered","not_stated","not_stated","unknown","Available official portal, partial work instructions and older legislation did not establish operative medical criteria.")
path("DZA","P2","Resident card including salaried foreign workers",U,C,"DZA_resident_rendered","yes","not_stated","exam_required",
     "Official instructions name general medicine, phtisiology and serology certificate categories. The actual disease criteria and serologic targets were not verified.")
e("DZA","P2","DZA_resident_rendered","GENERIC_FITNESS","Trois (3) certificats médicaux du demandeur (médecine générale - phtisiologie - sérologie)",
  "Que doit contenir le dossier ?, resident-card requirements","exam_required",en="Three medical certificates from the applicant (general medicine, phtisiology, serology)",
  kind="administrative_test",note="Phtisiology is a specialty associated with tuberculosis. The specific condition and the rule applied to that condition are not stated.")
settle("DZA",U,C,"Official residence guide identifies three categories of medical certificates.","Current work-permit health form and underlying clinical and serological criteria could not be verified.")

source("EGY_decree_279_gazette","Minister of Labour Decision No. 279/2025 on licences to employ foreigners","Egyptian Ministry of Labour / Al-Waqa’i al-Misriyya","gazette facsimile",date="2025-12-25",
       notes="Official-gazette issue 290 supplement B, 24 December 2025, pp. 31–43, reproduced by LawHub. See separately issued ministry procedural forms.")
source("EGY_alexandria_rendered","القوى العاملة — بحث طلبات تراخيص العمل للأجانب","Alexandria Governorate / Directorate of Manpower","official rendered service instructions",
       notes="Undated local procedure, retrieved via web_get_contents. Contains an AIDS-free certificate checklist and exemptions; continuing effect under the nationwide 2025 instrument was not verified.")
source("EGY_ministry_rendered","Work permits for foreigners","Egyptian Ministry of Labour","official legacy-site rendered work instructions",
       notes="Retrieved via web_get_contents. Instructions appear to describe an earlier version of the procedure, including older fees and prior decision references.")
path("EGY","P1","Current permit under Decision No. 279/2025",U,C,"EGY_decree_279_gazette EGY_alexandria_rendered EGY_ministry_rendered","not_stated","not_stated","unknown",
     "Contemporary implementation checklist could not be reconciled with undated local guidance naming AIDS.")
e("EGY","P1","EGY_alexandria_rendered","HIV","شهادة تفيد خلو الأجنبى من مرض نقص المناعة ( الإيدز )",
  "بحث طلبات تراخيص العمل للأجانب, checklist item before listed exemptions","exam_required",en="A certificate that the foreigner is free from immunodeficiency disease (AIDS).",
  kind="administrative_test",confidence="medium",note="Verified wording of undated official local guidance. Currentness after Decision 279/2025 unresolved; this row does not establish contemporary DSS.")
settle("EGY",U,C,"The undated Alexandria procedure names AIDS; it is recorded as a source observation.","Continuing effect of that requirement under Decision 279/2025 could not be verified in contemporary administrative instructions.",
       additional="Additional legal-version check targeted the new labour decision and Ministry of Labour service material.")

source("GEO_aliens_law","Ordinance No. 520 Approving Procedures for Reviewing and Deciding the Grant of a Georgian Residence Permit","Government of Georgia / Legislative Herald (Matsne)","government ordinance",date="2014-09-01",
       notes="English display has amendment links through 2026 and may not reproduce the latest consolidated wording.")
source("GEO_ordinance_current","Ordinance No. 70 on granting the right to work to labour immigrants and self-employed foreigners","Government of Georgia / Legislative Herald (Matsne)","government ordinance",
       notes="Georgian work-authorisation text, reviewed alongside the residence process.")
source("GEO_residence","Migration – Residence Permits","Public Service Development Agency of Georgia","official residence guidance")
path("GEO","P1","Right to work for labour immigrants",A,N,"GEO_ordinance_current","no","not_stated","unknown",
     "The work-authorisation ordinance sets out the application documents without a medical certificate or named medical condition.",
     "Finding limited to application for the right to work; separate residence route may request a health certificate.")
path("GEO","P2","Work residence permit",U,C,"GEO_aliens_law GEO_residence","conditional","not_stated","exam_required",
     "Article 4(1) allows a certificate when an outbreak abroad may pose a risk. Applicable disease designations or separately published operational criteria were not established.")
e("GEO","P2","GEO_aliens_law","GENERIC_FITNESS",
  "The Agency may request a health certificate in addition to the documents required by these Procedures in cases where there is a spread of disease in another country, and where the nature, gravity and duration of the disease may pose a risk to the population of Georgia.",
  "Annex, Article 4(1), General requirements for documents to be submitted","exam_required",kind="administrative_test",safeguard=yes,trigger="Outbreak in another country may pose risk")
saf("GEO","P2","GEO_aliens_law",discretion=yes,notes="Article 4(1) grants discretion to request a certificate under the outbreak condition.")
settle("GEO",U,C,"Separate work and residence adjudications; the work application checklist documents absence.",
       "Work residence rules include conditional outbreak-related certification without a verified condition schedule.",
       "Extended version check covers the work-authorisation system introduced in 2026.")

source("IRQ_consular_cph","شروط منح سمة الدخول — قانون الإقامة رقم (76) لسنة 2017","Iraqi Ministry of Foreign Affairs / Embassy in Copenhagen","official consular instruction",
       notes="Reports a notification from the Higher Committee for Health and National Safety. Statutory quotation names AIDS; wording elsewhere includes returning Iraqis, while the quoted passage refers to the foreign entrant.")
source("IRQ_statute_rendered","Foreigners' Residence Law No. 76 of 2017","Iraqi Ministry of Justice / Iraqi Gazette","partially rendered official English translation",language="en",
       notes="Ministry translation labelled 2021, with Article 3 intact; web_get_contents output is truncated later in Article 7.")
source("IRQ_rules_facsimile","Instructions No. 1 of 2026, Recruitment, Employment and Licensing of Foreign Workers","Iraqi Ministry of Labour and Social Affairs; Al-Taswia transcription","unofficial transcription",
       notes="Secondary copy of Iraqi Gazette No. 4863 dated 30 March 2026. Used for applicability context only.")
source("IRQ_krd_health_rendered","پشکنینی شاندکراوانی بیانی","Kurdistan Regional Government service portal / Ministry of Health","official rendered medical service description",
       notes="Rendered full visible Kurdish service text, retrieved with web_get_contents.")
source("IRQ_krd_work_rendered","پێدانی مۆڵەتی کارکردنی کرێکاری بیانی","Kurdistan Regional Government service portal / Ministry of Labour and Social Affairs","official rendered work-permit description")
path("IRQ","P1","Foreign worker permit under Instructions No. 1 of 2026",U,C,"IRQ_rules_facsimile","yes","not_stated","exam_required",
     "The transcription describes generic medical and communicable-disease proof. The official work permit schedule was not retrieved.")
e("IRQ","P1","IRQ_rules_facsimile","GENERIC_FITNESS","تقديم ما يثبت خضوعه للفحوصات الطبية وخلوه من الأمراض المعدية والسارية من جهة رسمية مختصة وتمتعه باللياقة الصحية والبدنية المطلوبة وبحسب نوع العمل.",
  "Instruction No. 1 of 2026, Article 2(5), private transcription","exam_required",
  en="Provide evidence from a competent official authority of medical examination, freedom from infectious and communicable diseases, and the required health and physical fitness relevant to the type of work.",kind="administrative_test",confidence="medium")
path("IRQ","P2","Entry/residence eligibility under Foreigners' Residence Law applicable to foreign workers",P,N,"IRQ_statute_rendered IRQ_consular_cph IRQ_rules_facsimile","yes","not_stated","exam_required",
     "Iraqi embassy reproduces the statutory requirement to certify freedom from acquired immunodeficiency syndrome. The 2026 worker-rule transcription directs foreign workers to comply with this residence law.",
     "The embassy statement supports a named general entry condition applicable to foreign entrants. Contemporary administration and the primary text of worker-specific rules were not verified.")
e("IRQ","P2","IRQ_consular_cph","HIV","ان يقدم ما يُثبت خلوه من الامراض السارية والمعدية والعوز المناعي المكتسب",
  "Consular notice quoting Foreigners' Residence Law 76/2017 arts. 8(5) and 3(3)","exam_required",
  en="To submit proof of freedom from communicable and infectious diseases and acquired immunodeficiency syndrome.",kind="administrative_test",confidence="medium")
e("IRQ","P2","IRQ_statute_rendered","HIV","He/ she shall prove free from communicable and contagious diseases and acquired immunodeficiency in accordance with the law.",
  "Gazette translation p. 5, Article 3(3)","exam_required",kind="administrative_test",note="Corroborates consular quotation from the statutory text.")
krd="IRQ_krd_health_rendered"
path("IRQ","P2-KRI","Kurdistan Region examination for foreigners arriving for work and stay",P,N,f"{krd} IRQ_krd_work_rendered","yes","destination","exam_required",
     "Regional government names HIV, hepatitis B and C, malaria and schistosomiasis. Selection of tests depends on the applicant's country of origin.",
     "Kurdistan Region only. The health service gives separate timing linked to residence card issuance or renewal.")
for code,phrase,en,other in (
    ("HIV","ڤایرۆسی ئایدز","HIV",""),
    ("HBV","هەوکردنی جگەری ڤایرۆسی جۆری B و C","Viral hepatitis B and C",""),
    ("HCV","هەوکردنی جگەری ڤایرۆسی جۆری B و C","Viral hepatitis B and C",""),
    ("MALARIA","مەلاریا","Malaria",""),
    ("OTHER_NAMED","بەلهاریزیا","Schistosomiasis","Schistosomiasis")):
    e("IRQ","P2-KRI",krd,code,phrase,"Service page, description, named communicable disease tests","exam_required",
      en=en,other=other,scope="destination",kind="administrative_test",trigger="Country of origin affects test selection")
trig("IRQ","P2-KRI",krd,"origin_country","Country the foreign entrant comes from",
     "The service explicitly says the types of tests vary by country of origin; a complete country-to-test mapping was not included.")
settle("IRQ",P,N,"Named federal entry/residence condition and separate Kurdistan Region screening. Federal worker permit panel remains unresolved.",
       additional="Additional applicability checks pursued the 2026 worker instruction and 2017 residence act.")

source("ISR_emb_work_fr","Demande de visa de travail en Israël","Embassy of Israel in France","official B/1 visa instructions")
source("ISR_emb_work_aze","Application for work visa in Israel","Embassy of Israel in Azerbaijan","official B/1 visa instructions")
source("ISR_bhr_ar_rendered","استمارة طلب تأشيرة عمل في إسرائيل","Embassy of Israel in Bahrain","official rendered B/1 visa instructions",
       notes="Rendered Arabic consular page. Requires proof of testing in a mission-approved facility; does not specify tests.")
source("ISR_b1_checklist","B1 就労ビザ (労働ビザ)","Embassy of Israel in Japan","official B/1 checklist",language="ja",
       notes="Official PDF gives medical certificate requirement with no specified examinations if medical fitness certified, and an exemption for qualifying Japanese professionals staying ≤45 days.")
isrf="Un bilan de santé comprenant les analyses suivantes: tuberculose, hépatite, HIV et un bilan sanguin général."
isra="The list of the medical tests: AIDS, hepatitis, tuberculosis and general blood test"
path("ISR","P1","B/1 work visa, France and Azerbaijan consular procedures",P,N,"ISR_emb_work_fr ISR_emb_work_aze","yes","origin_or_residence","exam_required",
     "Local consular procedures specify TB, hepatitis (type unspecified), and HIV/AIDS testing.","Scope is the published consular procedures; requirements can vary by application post.")
for key, original, en in (("ISR_emb_work_fr",isrf,"A health assessment including the following tests: tuberculosis, hepatitis, HIV, and general blood tests."),
    ("ISR_emb_work_aze",isra,"")):
    for c,other in (("TB",""),("HIV",""),("OTHER_NAMED","Hepatitis (type unspecified)")):
        e("ISR","P1",key,c,original,"Visa-specific additional information, named medical tests","exam_required",
          en=en,kind="administrative_test",scope="origin_or_residence",other=other)
path("ISR","P1-JPN","B/1 work visa, Japan consular procedure",A,N,"ISR_b1_checklist","conditional","origin_or_residence","exam_required",
     "Issued checklist states that no tests are specified when a doctor certifies fitness. It expressly exempts Japanese nationality professionals staying ≤45 days.")
e("ISR","P1-JPN","ISR_b1_checklist","GENERIC_FITNESS","指定はございませんが、「海外就労可能」の記載が必要となります",
  "PDF p. 1, medical certificate item","exam_required",en="If the regular physician certifies health, no particular examination items are specified, but the certificate must state that overseas employment is possible.",
  kind="administrative_test",scope="origin_or_residence")
trig("ISR","P1-JPN","ISR_b1_checklist","nationality","Japanese nationals in specialist occupations for stays of up to 45 days",
     "These workers are exempt from submitting a health certificate, PDF p. 1.")
path("ISR","P1-BHR","B/1 work visa, Bahrain consular procedure",U,C,"ISR_bhr_ar_rendered","yes","origin_or_residence","exam_required",
     "Consular page requires medical testing at approved sites and leaves the panel unspecified.")
e("ISR","P1-BHR","ISR_bhr_ar_rendered","GENERIC_FITNESS","تأكيد الفحوصات الطبية التي أجريت في المختبرات أو المستشفيات المعترف بها من قبل البعثة",
  "Required documents, medical examinations item","exam_required",en="Confirmation of medical tests done at laboratories or hospitals recognised by the mission.",
  kind="administrative_test",scope="origin_or_residence")
settle("ISR",P,N,"B/1 requirements vary by consular post; France and Azerbaijan document named tests, while the Japan checklist documents a generic certificate and a limited exemption.")
