source("OMN_national_manual","Medical Fitness Examination Guideline for Expatriates 2022","Oman Ministry of Health","official exam guideline",language="ar;en")
source("OMN_manual_home","Medical Fitness Examination Guideline","Oman Ministry of Health","official guidance index")
source("OMN_work_visa","Work Visa","Royal Oman Police","official visa instructions")
source("OMN_work_card","Get Medical Fitness Certificate for Work Card","Gov.om / Oman Ministry of Health","official service instructions")
wafid("OMN","OMN_work_visa","The Royal Oman Police separately lists the nationalities requiring a pre-entry medical certificate.")
trig("OMN","P1","OMN_work_visa","nationality","India, Pakistan, Philippines, Bangladesh, Indonesia, Sri Lanka, Egypt, Sudan, Ethiopia, Syria and Nepal",
     "Royal Oman Police work-visa certificate requirement, independently of the regional programme's examination location list.")
omn="OMN_national_manual"
path("OMN","P2","Destination medical fitness, including work-card certification",P,N,f"{omn} OMN_manual_home OMN_work_card","yes","destination","exam_required;unfit_certificate",
     "Oman Ministry of Health has destination protocols for HIV, syphilis, HBV, HCV and suspected TB, with differentiated treatment and occupation handling.")
for page,code,phrase in [(5,"HIV","HIV test 1&2 for New applicants & Renewals"),(6,"SYPH","Syphilis test for New applicants & Renewals"),
    (7,"HBV","Hepatitis (B) for New applicants & Renewals"),(8,"HCV","Hepatitis (C) for New applicants & Renewals"),
    (9,"TB","Suspected cases of TB for new applicants"),(10,"TB","Suspected cases of TB for Renewal applicants")]:
    e("OMN","P2",omn,code,phrase,f"PDF p. {page} (English flowchart heading)","exam_required",en=phrase,kind="administrative_test",scope="destination")
for page,code,note in [(5,"HIV","Positive confirmatory branch for HIV."),(7,"HBV","Category-specific positive determination and private-driver exception shown in the flowchart."),
    (8,"HCV","Category-specific positive determination and private-driver exception shown in the flowchart."),
    (9,"TB","Confirmed positive new-applicant branch.")]:
    e("OMN","P2",omn,code,"All categories Unfit" if code=="TB" else "Unfit",f"PDF p. {page}, positive-result flowchart branch","unfit_certificate",
      en="All categories Unfit" if code=="TB" else "Unfit",scope="destination",note=note)
saf("OMN","P2",omn,individual=yes,retest=yes,notes="PDF pp. 6–10: conditional syphilis follow-up and retest, private-driver hepatitis treatment commitment, and differentiated TB renewal. These provisions concern non-HD conditions.")
for code,page in [("HBV",7),("HCV",8)]:
    e("OMN","P2",omn,code,"Private driver is fit for residence after sponsor commitment for treatment & follow up",
      f"PDF p. {page}, private-driver exception","unknown",en="Private driver is fit for residence after sponsor commitment for treatment & follow up",
      scope="destination",note="Conditional fitness exception to unfitness for the named condition.",trigger="Private driver")
trig("OMN","P2",omn,"occupation","Private driver",notes="Conditional hepatitis B/C fitness after sponsor commitment, PDF pp. 7–8.")
settle("OMN",P,H,"Wafid names HD for a covered pre-departure pathway. Destination examination protocol separately names other diseases.")

source("PSE_labour_law","Palestinian Labour Law No. 7 of 2000","Palestinian Ministry of Health","statute snapshot",notes="Work permits for non-Palestinians are addressed separately from occupational medical requirements.")
path("PSE","P1","Work permit for non-Palestinian workers",U,C,"PSE_labour_law","not_stated","not_stated","unknown",
     "Occupational examination rules do not establish a medical criterion for the immigration/work permit.")
settle("PSE",U,C,"Occupational and regulatory materials located.","Full text of applicable permit decision and any work-permit-specific health criteria not verified. University law-database URLs returned general labour-law indexes.")

source("QAT_moph_rendered","Medical fitness Examination for residency and employment purposes","Qatar Ministry of Public Health","official rendered service description",
       notes="Rendered full visible service text, retrieved with web_get_contents; raw direct HTTP page also inspected. The medical panel is not listed.")
source("QAT_sharek","Medical Examination for Issuing Residency Permits or Extending Visas","Qatari government (Sharek) / Ministry of Public Health","official procedure")
source("QAT_phcc","القومسيون الطبي","Qatar Primary Health Care Corporation","official service description")
source("QAT_france_diplomatie","Qatar – Entrée / séjour, Examens médicaux","French Ministry for Europe and Foreign Affairs","third-country consular guidance",
       notes="Names HIV/AIDS, HBV, HCV, TB and syphilis. Source does not establish the Qatari destination panel at S1–S4 tier.")
wafid("QAT",notes="The GCC pre-departure pathway is distinct from the destination Medical Commission assessment.")
path("QAT","P2","Medical Commission screening for residency and employment",U,C,"QAT_moph_rendered QAT_sharek QAT_phcc QAT_france_diplomatie","yes","destination","exam_required",
     "Qatari health agencies verify the examination. The full named panel remains unresolved from Qatari sources; third-country consular guidance provides a lead.")
e("QAT","P2","QAT_moph_rendered","GENERIC_FITNESS","The service is provided as per medical, health conditions required to be met by the residence and work applicant in accordance with the State laws.",
  "Medical fitness Examination for residency and employment purposes, service description","exam_required",scope="destination",kind="administrative_test")
trig("QAT","P2","QAT_phcc","occupation","Newly arriving domestic workers, such as housekeepers and family drivers",
     "PHCC designates certain health centres where the required residence tests can be completed; this is a service-location difference.")
settle("QAT",P,H,"Shared Wafid coverage. Destination examination requirement confirmed, operative domestic panel unresolved.",
       additional="Extended search compared Qatari ministry/clinic text with foreign consular guidance for the domestic panel.")

source("SAU_expats_ministry","Guide to Services Provided for Expatriates","Saudi Ministry of Human Resources and Social Development","official administrative guide")
source("SAU_moh_verification","E-Verification of Expats’ Medical Examination Results Service","Saudi Ministry of Health","official service instructions")
wafid("SAU",notes="Saudi domestic residence delivery is reviewed separately.")
path("SAU","P2","Domestic examination for residence issuance",U,C,"SAU_expats_ministry SAU_moh_verification","yes","destination","exam_required",
     "Government guide requires a passed examination in approved centres. Domestic clinical panel and consequences were not established.")
e("SAU","P2","SAU_expats_ministry","GENERIC_FITNESS","Passing the medical examination for a family member or an expatriate in approved",
  "PDF p. 56, Terms of Service, residence issuance","exam_required",en="Passing the medical examination for a family member or an expatriate in approved medical centers.",scope="destination",kind="administrative_test")
settle("SAU",P,H,"HD confirmation confined to the shared Wafid pre-departure mechanism. Domestic administrative guidance establishes an examination.")

source("SDN_emb_work","Work and Humanitarian visa","Embassy of Sudan in Sweden","official consular instructions",
       notes="Consular route as displayed. Coverage of domestic health certification and rules in current administrative practice remains uncertain.")
source("SDN_non_sudanese_law_copy","Law Regulating Employment of Non-Sudanese of 2001","Sudanese legislature; HUDO Centre copy","legal reproduction",
       notes="Lead for work-permit requirements. Official implementing text and applicable medical schedule unresolved.")
path("SDN","P1","Employment visa and work permit for non-Sudanese workers",U,C,"SDN_emb_work SDN_non_sudanese_law_copy","not_stated","not_stated","unknown",
     "Available consular instruction and legal reproduction could not establish whether separate medical requirements apply.")
settle("SDN",U,C,"Consular material and reproduction of non-Sudanese employment law reviewed.","No verified contemporary operative work-permit medical requirements; direct work-authority materials unavailable.",
       "Follow-up sought the enabling statute and competent authority publication because of incomplete portal access.")
