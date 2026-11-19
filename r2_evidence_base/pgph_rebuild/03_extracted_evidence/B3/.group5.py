source("SYR_embassy_residence","Residency in Syria","Embassy of Syria in Australia","official consular checklist",
       notes="Undated official-hosted PDF. One-year checklist includes Department of Labour and Social Affairs approval and HIV test. Contemporary domestic implementation could not be independently confirmed.")
source("SYR_law2","Law No. 2 of 2014 on Entry, Exit and Residence of Foreigners","Syrian People's Assembly","official legal text")
path("SYR","P2","One-year residence requiring labour authority approval",P,N,"SYR_embassy_residence SYR_law2","yes","not_stated","exam_required",
     "Published official consular checklist for one-year residence explicitly requires an HIV test. It also requires labour approval.",
     "Undated checklist. Current in-country implementation and shorter/work-specific streams were not verified.")
e("SYR","P2","SYR_embassy_residence","HIV","- HIV test.","PDF p. 1, section 2, one-year residence checklist","exam_required",kind="administrative_test",confidence="medium")
trig("SYR","P2","SYR_embassy_residence","visa_class","One-year residence",notes="The same source has a distinct one-to-three-month checklist without this item.")
settle("SYR",P,N,"One-year labour-approved residence has a named test on the published consular form. Current domestic application remains uncertain.",
       additional="Follow-up addressed whether the consular checklist remains operative after the governmental transition.")

source("TUN_sicad_residence","Octroi de la carte de séjour pour étrangers","Tunisian public services portal (SICAD) / Ministry of Interior","official application checklist")
source("TUN_jort_foreigners","Law No. 68-7 of 8 March 1968 on the status of foreigners, Journal officiel No. 11","Tunisia; JORT.tn republication","legal republication",
       notes="Third-party legal database. Law read for framework only.")
source("TUN_work_rendered","Emploi des étrangers","Tunisian Ministry of Employment and Vocational Training","official rendered procedures index",
       notes="Confirms work contract endorsement or certificate of exemption, with category-specific document lists linked. Rendered copy includes index but not the linked checklists.")
source("TUN_contract_rendered","Contrat de travail pour travailleur étranger","Tunisian Ministry of Employment and Vocational Training","rendered official contract form",
       notes="Form does not enumerate all separate supporting documents.")
path("TUN","P1","Authorisation and endorsement of a foreign worker's employment contract",U,C,"TUN_work_rendered TUN_contract_rendered TUN_jort_foreigners","not_stated","not_stated","unknown",
     "Ministry index links occupation- and class-specific supporting-document lists. Complete medical requirements for endorsement were not verified.")
path("TUN","P2","Worker residence card (salarié)","documented_absence",N,"TUN_sicad_residence","no","not_stated","unknown",
     "The government procedure's complete list for a paid worker includes employment, housing, travel and application documents and no medical requirement.",
     "Scope is the worker residence card checklist. It requires a previously endorsed work contract or exemption.")
settle("TUN",U,C,"Worker residence-card checklist documents absence at pathway level.","Complete current medical requirements for the prior employment-contract endorsement could not be determined.")

source("TUR_work_documents","Yabancıların Çalışma Hayatı – Belgeler","Türkiye Ministry of Labour and Social Security","official work permit checklists")
source("TUR_work_faq","Çalışma İzni","Türkiye Presidency of Migration Management","official FAQ")
source("TUR_statute_6458","Yabancılar ve Uluslararası Koruma Kanunu, Law No. 6458","Türkiye Legislation Information System","consolidated statute")
path("TUR","P1","Standard work permit application",A,N,"TUR_work_documents TUR_work_faq","no","not_stated","unknown",
     "Ministry's required-documents checklist for initial application contains no health-exam item; health-sector professional approval is credentialing.",
     "Scope is the work permit procedure itself.")
path("TUR","P2","Related general entry/visa eligibility for foreign workers",U,C,"TUR_statute_6458 TUR_work_faq","not_stated","not_stated","entry_denial;visa_refusal_inadmissibility",
     "Law 6458 links the ordinary work permit to residence and bars visa issuance to carriers of diseases deemed a threat to public health. The applicable designated schedule was not verified.")
e("TUR","P2","TUR_statute_6458","GENERIC_FITNESS","Kamu sağlığına tehdit olarak nitelendirilen hastalıklardan birini taşıyanlar",
  "Article 15(1)(ç); see Article 7(1)(c) cross-reference and Article 27 on work permission as residence","visa_refusal_inadmissibility",
  en="Persons carrying one of the diseases classified as a threat to public health.")
settle("TUR",U,C,"Work-permit application checklist documents absence at pathway level.","Separate visa/entry rules contain a generic designated-disease restriction; a binding list applicable to this restriction was not verified.",
       additional="Extended search for the designations under Law 6458 and a potential HD connection.")

source("YEM_entry_law","Law No. 47 of 1991 on Entry and Residence of Foreigners","Yemen National Information Centre","official legal text")
source("YEM_labour_law","Labour Law No. 5 of 1995","Yemen National Information Centre","official legal text")
source("YEM_work_exec","Executive Regulation of the Labour Law","Yemen National Information Centre","official legal text")
path("YEM","P1","Work authorisation for non-Yemeni workers",U,C,"YEM_labour_law YEM_work_exec YEM_entry_law","not_stated","not_stated","unknown",
     "Labour Law Article 20(2) requires complete health fitness but does not specify a disease, test or procedure. Current implementing instructions could not be verified.")
e("YEM","P1","YEM_labour_law","GENERIC_FITNESS","استيفاء شروط شغل المهنة وان يكون المصرح له بالعمل في لياقة صحية تامة.",
  "Labour Law, Article 20(2)","unknown",en="Meet the occupational requirements and ensure the person authorised to work is in full health fitness.")
settle("YEM",U,C,"Generic health fitness is statutory.","No current administering authority's requirements or complete medical schedule verified; authority and geographic applicability remain uncertain.")
