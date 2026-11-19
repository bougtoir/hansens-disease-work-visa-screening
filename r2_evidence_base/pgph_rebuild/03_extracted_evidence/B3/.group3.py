source("JOR_health_certificate","استصدار الشهادة الصحية للوافد لغايات إكمال إجراءات الإقامة والعمل","Jordanian Ministry of Health","official medical procedure",date="2022-12-26")
source("JOR_work_principles","The Principles for Issuing Work Permits for Non-Jordanian Workers with Specialized Skills for the Year 2024","Jordanian Ministry of Labour","official administrative instructions")
path("JOR","P1","Non-Jordanian worker health certification for residence and employment",P,N,"JOR_health_certificate JOR_work_principles","yes","destination","exam_required",
     "The health certificate procedure explicitly describes hepatitis B/C testing and TB investigations after a suspicious chest X-ray.",
     "The permit principles concern specialised workers. The Ministry of Health workflow is for residence and employment.")
for code in ("HBV","HCV"):
    e("JOR","P1","JOR_health_certificate",code,"التهاب الكبد الوبائي","PDF pp. 8–9, step 10, hepatitis B or C; Roman B/C and PCR also appear on the printed lines",
      "exam_required",en="Viral hepatitis (B or C, as stated in the step)",kind="administrative_test",scope="destination")
e("JOR","P1","JOR_health_certificate","TB","مرض التدرن",
  "PDF p. 9, step 13, tests if chest imaging raises suspicion","exam_required",en="Required tests for diagnosing tuberculosis",
  kind="administrative_test",scope="destination",trigger="Suspicious chest X-ray")
saf("JOR","P1","JOR_health_certificate",individual=yes,notes="PDF p. 9, step 13: physician requests diagnostic tests when chest imaging suggests TB. HD-specific assessment is unverified.")
settle("JOR",P,N,"Named tests and conditional follow-up in the Ministry of Health certificate instructions.")

source("KWT_decision139_report","Kuwait Adopts the Updated Wafid Programme Regulations","GLA & Company","secondary legal analysis",
       notes="Lead for national adoption and authoritative-text retrieval. Does not establish fitness criteria.")
kwt_case=source("ARE_denial_lead1","The deportation of two Ethiopian migrant workers with leprosy","Lambert and Walker, Leprosy Review 86(1), 108–111 (2015)","peer-reviewed clinical case report",
       notes="LSHTM repository copy. Observed removal evidence from Kuwait. A fresh download was saved with its original search-lead filename.",jurisdiction="KWT")
wafid("KWT",notes="The separately reported national adoption in Decision 139/2026 was not retrieved from the official gazette.")
path("KWT","P2","Destination examination for private-sector work residence",U,C,"KWT_decision139_report","not_stated","destination","unknown",
     "Contemporary national health requirements in the official worker portal and Ministerial Decision 139/2026 were inaccessible. The register records a legal lead for national adoption of Wafid.")
row("OBSERVED_DENIALS",case_id="OD-KWT-001",jurisdiction_id="J-KWT",date="2013",
    person_status="Ethiopian migrant domestic worker",clinical_status="Borderline tuberculoid HD, bacterial index zero. Initial drugs were dispensed before removal, and MDT was completed in Ethiopia later.",
    decision="removal_deportation",source_type="S5 peer-reviewed clinical case report",source_id=kwt_case,
    notes="Lambert and Walker, Case 2, journal pp. 109–110. On-treatment case; dispense timing documented, treatment completion before expulsion and microbiological proof of non-infectiousness unverified. Exact removal date not given; April 2013 biopsy anchors the chronology.")
row("OBSERVED_DENIALS",case_id="OD-KWT-002",jurisdiction_id="J-KWT",
    person_status="Ethiopian migrant cleaner",clinical_status="Borderline lepromatous HD with type 1 reaction confirmed after return. Received treatment during hospitalisation in Kuwait and completed MDT later in Ethiopia.",
    decision="removal_deportation",source_type="S5 peer-reviewed clinical case report",source_id=kwt_case,
    notes="Lambert and Walker, Case 1, journal pp. 108–109. Police transferred the worker from hospital to a detention centre and then onto a flight to Ethiopia. Date of removal and proof of non-infectiousness were not reported.")
settle("KWT",P,H,"Shared Wafid coverage and two separately recorded removals described in the same clinical report.",
       additional="Further searches targeted national adoption instrument, Article 18 checklist, domestic screening rules, and treated-case evidence.")

source("LBN_general_security","Documents required to recruit a worker in domestic services","Lebanese General Directorate of General Security","official work and residence instructions")
source("LBN_res_fr_rendered","Les documents requis pour l'obtention d'un travailleur dans les services domestique","Lebanese General Directorate of General Security","official rendered work and residence instructions")
source("LBN_residency_ar","تسوية أوضاع العمال والعاملات العرب والأجانب","Lebanese General Directorate of General Security","official regularisation notice")
path("LBN","P1","Work permit for foreign domestic workers",U,C,"LBN_general_security LBN_res_fr_rendered","yes","not_stated","exam_required",
     "General Security names laboratory tests in section 3, domestic-worker work permit. A test panel or clinical schedule was not located.")
e("LBN","P1","LBN_general_security","GENERIC_FITNESS","Medical laboratory tests to which the foreigner is submitted",
  "Domestic service instructions, section 3, work permit documents","exam_required",kind="administrative_test")
trig("LBN","P1","LBN_res_fr_rendered","occupation","Foreign domestic service workers")
path("LBN","P2","Residence card after obtaining a domestic worker permit",A,N,"LBN_general_security LBN_res_fr_rendered","no","not_stated","unknown",
     "Section 4 of the complete domestic-service residence checklist calls for the work permit and other documents without another medical examination.",
     "The prior work permit laboratory-test requirement remains unresolved.")
settle("LBN",U,C,"General Security lists laboratory tests for domestic worker work permits.","Applicable laboratory panel and health certificate were not retrieved.")

source("LBY_ncdc_certificate","آلية استخراج الشهادة الصحية الإلكترونية","Libyan National Centre for Disease Control","official health-certificate instructions")
source("LBY_decree98","Decision No. 98 of 2007 on rules and mechanisms governing foreign labour","General People's Committee; Law Society of Libya transcription","private legal transcription",
       notes="Cited to contextualise workers' international health certificate; primary gazette version/current status unverified.")
path("LBY","P1","Foreign-worker health certification",U,C,"LBY_ncdc_certificate LBY_decree98","yes","destination","exam_required",
     "Disease Control Centre instructions name imaging and a triple laboratory test without identifying its components. Private transcription of the employment rule is also generic.")
e("LBY","P1","LBY_ncdc_certificate","GENERIC_FITNESS","إتمام الفحص الطبي ( أشعة الصدر – التحليل الثلاثي في المختبر المرجعي ).",
  "Instructions for electronic health certificate, medical examination step","exam_required",
  en="Complete the medical examination (chest X-ray and triple test at the reference laboratory).",kind="administrative_test",scope="destination")
settle("LBY",U,C,"Verified generic health certificate workflow.","Official analytes and consequences for foreign-worker certification could not be verified.",
       "Follow-up investigated whether the named triple test could be resolved from an official primary source.")

source("MAR_residence_decree","Decree No. 2.09.607 implementing Law No. 02.03 on entry and residence","Moroccan Ministry of Justice (Adala)","official decree")
source("MAR_immigration_law","Law No. 02.03 on the Entry and Residence of Foreigners in the Kingdom of Morocco and Illegal Emigration","Moroccan Ministry of Justice (Adala)","official law")
source("MAR_us_embassy","Residency Requirements for Foreigners in Morocco","U.S. Embassy in Morocco","external consular guidance",tier="S5",
       notes="Third-country official guidance. Operational lead only.")
source("MAR_work_rendered","Procédure d’emploi des salariés étrangers au Maroc","Moroccan Ministry of Labour and Professional Integration","partial rendered administrative instructions",
       notes="Marked updated December 2018. Only introductory and legal-framework sections exposed by web_get_contents; current full document criteria could not be checked.")
path("MAR","P1","Approval of the foreign worker employment contract",U,C,"MAR_work_rendered","not_stated","not_stated","unknown",
     "Ministry procedure verifies the labour authorisation mechanism. Full contemporary document requirements could not be checked.")
path("MAR","P2","Work-related registration and residence under Law 02.03",U,C,"MAR_residence_decree MAR_immigration_law MAR_us_embassy","not_stated","not_stated","unknown",
     "Verified the legal framework, but could not establish a complete current worker-specific medical requirement or disease schedule.")
settle("MAR",U,C,"Moroccan law and implementing decree read.","Worker-specific health certificate format and administration remain unverified.")
