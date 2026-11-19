source("ARE_wafid_reg_2021", "Regulations of Medical Examination of Expatriates Coming to GCC States for Work or Residence, Sixth Version 2021", "Gulf Health Council", "regional regulations", status="historical_version", notes="Approved 2021-10-21. GHC reports subsequent update. Read for version history, concordance and historical complaints process.")
source("ARE_ghc_program", "Examination of arrivals", "Gulf Health Council", "official programme page", notes="Identifies GCC scheme and reports an October 2023 update.")
source("ARE_wafid_faq_archive", "Wafid FAQ, English", "Gulf Health Council", "archived operational FAQ", status="archived_snapshot", notes="Unmodified pre-rebuild rendered HTML, captured 2026-11-05 and re-read. Public links only; shared regional source.")
source("SAU_wafid_faq_ar", "وافد — الأسئلة الشائعة", "Gulf Health Council", "operational FAQ", status="observed_on_official_website", notes="Arabic edition available at collection time. Shared by all six destinations.")

wreg = "ARE_wafid_reg_2021"
wghc = "ARE_ghc_program"
wfaq = "ARE_wafid_faq_archive"
war = "SAU_wafid_faq_ar"
wcore = f"{war} {wfaq} {wghc}"
wloc = "FAQ, medical unfitness cases, Infectious Diseases"
wloc_noninfectious = "FAQ, medical unfitness cases, Non-infectious Diseases"
wafid_tests = (
    ("HIV", "1. فيروس نقص المناعة البشرية (الإيدز)", "1. Human immunodeficiency virus (AIDS)"),
    ("HBV", "2. التهاب الكبد (ب) - (نتيجة إيجابية للمستضد السطحي)", "2. Hepatitis B (positive surface antigen result)"),
    ("HCV", "3. التهاب الكبد (ج) (نتيجة إيجابية للأجسام المضادة)", "3. Hepatitis C (positive antibody result)"),
    ("FILARIASIS", "4. الديدان الخيطية والملاريا", "4. Filarial worms and malaria"),
    ("MALARIA", "4. الديدان الخيطية والملاريا", "4. Filarial worms and malaria"),
    ("HD", "5. الجذام", "5. Leprosy"),
    ("TB", "أ. وجود أدلة سابقة على الإصابة بمرض السل", "a. Previous evidence of tuberculosis"),
)

def wafid(j, national="", notes="", pw="P1", coverage=""):
    keys = " ".join(filter(None,(wcore,national)))
    path(j,pw,"Wafid pre-departure examination for covered applicants",P,H,keys,"yes","origin_or_residence",
         "exam_required;unfit_certificate","Gulf Health Council Wafid operational FAQ explicitly lists conditions evaluated as unfit.",
         "Shared regional mechanism. Applies to covered pathways or examination locations. " + notes)
    for c, original, translation in wafid_tests:
        e(j,pw,war,c,original,wloc,"unfit_certificate",en=translation,scope="origin_or_residence",
          trigger="Wafid-covered examination location",note="Unfitness list. The source does not classify subsequent visa decisions.")
    for n,c,original,label in (
        (1,"OTHER_NAMED","الفشل الكلوي","Kidney failure"),
        (2,"OTHER_NAMED","فشل الكبد أو قصور الكبد","Liver failure or insufficiency"),
        (3,"OTHER_NAMED","فشل القلب","Heart failure"),
        (4,"OTHER_NAMED","ارتفاع ضغط الدم غير المنضبط","Uncontrolled hypertension"),
        (5,"OTHER_NAMED","مرض السكري غير المنضبط","Uncontrolled diabetes mellitus"),
        (6,"OTHER_NAMED","أنواع مختلفة من السرطان","Various cancers"),
        (7,"MENTAL","الاضطرابات النفسية والعصبية","Psychiatric and neurological disorders"),
        (8,"OTHER_NAMED","أي تشوه أو بتر أو إعاقة جسدية تعيق أداء صاحب الطلب","Any deformity, amputation or physical disability that impairs the applicant’s performance")):
        e(j,pw,war,c,f"{n}. {original}",wloc_noninfectious,"unfit_certificate",en=f"{n}. {label}",
          scope="origin_or_residence",other=label if c=="OTHER_NAMED" else "",trigger="Wafid-covered examination location",
          note="Rule is phrased separately from leprosy. No post-HD-sequelae application was established." if n==8 else "")
    e(j,pw,wfaq,"HD","5. Leprosy",wloc,"unfit_certificate",scope="origin_or_residence",
      trigger="Wafid-covered examination location",note="Corroborating 2026-11-05 English archived FAQ wording.")
    post(j,pw,f"{war} {wfaq}","no_treatment_distinction_named_condition",
         "Leprosy is unqualified in the FAQ. Effect on a treated, cured, or non-infectious applicant cannot be determined.",
         "Listed as medically unfit without a stated distinction by treatment or infectiousness.")
    saf(j,pw,f"{war} {wreg}",notes="2021 rules, annex 6 PDF pp. 55–58, allowed complaints about the examination service; formal appeals there were by the centre against its penalty. Current applicant redress against the fitness or visa decision could not be verified.")
    trig(j,pw,f"{war} {wfaq}","programme_coverage", "Source-country/location list in Wafid FAQ",
         "The official programme lists countries and booking selects examination country, city and GCC destination. " + coverage)

source("ARE_2022_moh_rendered","Cabinet Resolution No. (75) of 2022 amending Cabinet Resolution No. (7) of 2008 — legislation brief",
       "UAE Ministry of Health and Prevention","official rendered legislation brief",
       notes="Retrieved as rendered full visible text with web_get_contents. The legislation brief and whole PDF have distinct evidence status; the PDF was unavailable.")
source("ARE_work_guide_rendered","Preparing to work","Official Platform of the UAE Government","official rendered administrative guidance",
       notes="Retrieved with web_get_contents. The page describes residency denial or nonrenewal for leprosy-positive cases and cites earlier amending Resolution 5/2016.")
source("ARE_dha_circular","Dubai Health Authority Circular CIR-2022-00000195", "Dubai Health Authority", "official circular notice",
       date="2022-09-14",notes="Notice links Cabinet Decision 75 of 2022. Attachment request returned portal shell; notice alone does not establish the disease schedule.")
wafid("ARE", notes="The UAE domestic residence programme is separately evidenced.")
are_moh="ARE_2022_moh_rendered"
are_work="ARE_work_guide_rendered"
are_keys=f"{are_moh} {are_work} ARE_dha_circular"
path("ARE","P2","UAE residence medical under Cabinet Decisions 7/2008 and 75/2022",P,H,are_keys,"yes","destination","exam_required;unfit_certificate;visa_refusal_inadmissibility",
     "Federal ministry's official summary of the 2022 amendment names HD and specifies unfitness for residence when positive on first issue or renewal. Official government work guide corroborates residence denial/nonrenewal.",
     "The full amended instrument could not be retrieved; scope and consequences follow the two government descriptions.")
for c,phrase,other in (
    ("HIV","screening for AIDS (acquired immunodeficiency virus)",""),
    ("OTHER_NAMED","screening for viral hepatitis, which is limited to certain groups","Viral hepatitis (type unspecified)"),
    ("TB","an examination for tuberculosis, which is conducted on all new arrivals to the country",""),
    ("OTHER_NAMED","a pregnancy examination that is limited to domestic workers","Pregnancy")):
    e("ARE","P2",are_moh,c,phrase,"Legislation brief, amended Articles 2 and 3","exam_required",kind="administrative_test",scope="destination",other=other)
e("ARE","P2",are_moh,"HD","an examination of leprosy, of which positive cases are considered unfit for residence purposes, whether for new arrivals or upon renewal",
  "Legislation brief, amended Articles 2 and 3","unfit_certificate",scope="destination")
e("ARE","P2",are_work,"HD","Leprosy test: conducted for new applicants and upon renewal. Residency is denied or not renewed for positive cases.",
  "Health screening, item 4","visa_refusal_inadmissibility",scope="destination")
for c,phrase,num in (("HBV","Hepatitis B test",2),("SYPH","Syphilis test",5)):
    e("ARE","P2",are_work,c,phrase,f"Health screening, item {num}","exam_required",scope="destination",kind="administrative_test",
      note="Official operational guide citing earlier amendment. Subsequent official legislative brief gives a summary, not the full schedule.")
trig("ARE","P2",are_moh,"occupation","Domestic workers",notes="Pregnancy examination only. The HD rule is described for new and renewal applicants.")
post("ARE","P2",f"{are_moh} {are_work}","no_treatment_distinction_named_condition",
     "Official descriptions mention a positive result but do not define whether a previously treated, cured, or non-infectious person can be considered positive by the programme.",
     "Positive HD test results trigger residence unfitness at initial application and renewal.")
saf("ARE","P2",are_moh,notes="Official brief mentions health-fitness exemption criteria generally. Its HD-specific applicability could not be determined from the brief.")
settle("ARE",P,H,"HD confirmed in the shared Wafid mechanism and separately in the national residence medical programme.",
       additional="Additional searches pursued national HD provision, current regional technical appendix and case reports because they could change treatment and consequence interpretation.")

source("ARM_arlis_231843","Law of the Republic of Armenia on Foreigners (official English text)","National Assembly of Armenia / ARLIS","statute",
       notes="Article 8(1)(d) refers to the Government list; Article 20 provides judicial review of residence refusal.")
source("ARM_arlis_95625","Government Decision No. 49-N of 25 January 2008 approving the infectious diseases that bar entry of foreigners or stateless persons", "Government of Armenia / ARLIS", "government decision",
       date="2015-02-28",status="official_portal_indicates_in_force",notes="Date of current incorporated version shown by portal. The decision implements Article 8(1)(d).")
source("ARM_arlis_amend_current","HO-11-N amending the Law on Foreigners, adopted 20 January 2026","National Assembly of Armenia / ARLIS","amending statute",
       date="2026-11-01",notes="Article 12 replaces Article 19 and retains the disease cross-reference in amended 19(1)(4).")
source("ARM_work_residence","Acquiring Employment-based Residence Status","Migration and Citizenship Service of Armenia","official procedure")
arm_src="ARM_arlis_231843 ARM_arlis_95625 ARM_arlis_amend_current ARM_work_residence"
path("ARM","P2","Employment-based residence status subject to the Law on Foreigners",P,N,arm_src,"not_stated","not_stated","visa_refusal_inadmissibility",
     "Amended Law on Foreigners Article 19(1)(4) incorporates Article 8(1)(d); Decision 49-N lists five conditions.",
     "This is a named eligibility restriction. Universal medical examination not established.")
saf("ARM","P2","ARM_arlis_231843 ARM_arlis_amend_current",review=yes,notes="Article 20(1) provides judicial review of residence refusal. The amended reapplication rule is an administrative rule and does not establish a disease waiver.")
for c, word, trans, num in (
    ("OTHER_NAMED","Ժանտախտ (թոքային ձև)","Plague (pneumonic form)",1),
    ("OTHER_NAMED","Խոլերա","Cholera",2),
    ("TB","Շնչառական օրգանների ակտիվ տուբերկուլոզ (հարուցչի արտազատմամբ` բոլոր ձևերը)","Active respiratory tuberculosis (all forms with excretion of the pathogen)",3),
    ("MALARIA","Արևադարձային մալարիա","Tropical malaria",4),
    ("OTHER_NAMED","Ատիպիկ թոքաբորբ","Atypical pneumonia",5)):
    e("ARM","P2","ARM_arlis_95625",c,word,f"Decision 49-N, annex item {num}, via amended Law on Foreigners Article 19(1)(4)","visa_refusal_inadmissibility",
       en=trans,other=trans if c=="OTHER_NAMED" else "",note="Eligibility condition by statutory cross-reference; examination procedure not specified.")
settle("ARM",P,N,"Employment-based residence has a named eligibility rule by cross-reference.",
       additional="Extended applicability check because the Law on Foreigners changed on 2026-11-01 and the implementing disease list is separate.")

source("AZE_work_asan_www","Issuance of work permits to foreigners and stateless persons to engage in paid employment", "ASAN Service / State Migration Service", "official application requirements",
       status="observed_on_official_website")
source("AZE_migration_74_www","Azərbaycan Respublikasının ərazisində müvəqqəti yaşamaq üçün icazənin verilməsi", "State Migration Service of Azerbaijan", "official application requirements",
       status="observed_on_official_website")
aze_work="a medical certificate of a foreigner and stateless person not being a carrier of infectious diseases (human immunodeficiency virus infection, hepatitis B and C viruses) approved by the Cabinet of Ministers of the Republic of Azerbaijan."
aze_res="müvafiq icra hakimiyyəti orqanı tərəfindən təsdiq edilən təhlükəli yoluxucu xəstəliklərin siyahısında nəzərdə tutulmuş xəstəlik virusu daşıyıcısı olmaması haqqında arayış (insanın immunçatışmazlığı virusu infeksiyası, B və C hepatitləri təhlükəli yoluxucu xəstəlikləri);"
for pw,key,name,word,loc in (
    ("P1","AZE_work_asan_www","Paid-employment work permit",aze_work,"Required documents, medical certificate item"),
    ("P2","AZE_migration_74_www","Temporary residence certificate requirement",aze_res,"Temporary residence, required documents, dangerous infectious diseases certificate")):
    path("AZE",pw,name,P,N,key,"yes","not_stated","exam_required","Official permit checklist explicitly requires certification concerning HIV, HBV and HCV.",
         "Certificate condition; exam location and any discrete visa-refusal decision were not established.")
    for c in ("HIV","HBV","HCV"):
        e("AZE",pw,key,c,word,loc,"exam_required",en=aze_work if pw=="P2" else "",kind="administrative_test")
settle("AZE",P,N,"Named certificate rules on two official agency pages. The full enabling Cabinet instrument was not retrieved.")

source("BHR_lmra_mobile","Medical Examination for Expatriate Employees","Bahrain Labour Market Regulatory Authority","official work-permit procedure")
source("BHR_lmra_health","New Work Permit","Bahrain Labour Market Regulatory Authority","official work-permit procedure")
source("BHR_moh_program","Expatriate Health Check Up","Bahrain Ministry of Health","official medical examination instructions")
source("BHR_lmra_guidelines","Expatriate Medical Checkup Guidelines","Bahrain Ministry of Health","official medical examination guidelines")
source("BHR_lmra_requirements","Medical checkup requirement for permission work visa in kingdom of Bahrain","Bahrain Labour Market Regulatory Authority","official alternate examination form")
source("BHR_regulations_2022","Minister of Health Decision No. 21 of 2022 amending Decision No. 30 of 2017 on expatriate medical examinations",
       "Bahrain Legislation and Legal Opinion Commission / Ministry of Health","ministerial decision",date="2022-08-28")
wafid("BHR","BHR_lmra_mobile BHR_lmra_health","LMRA explicitly directs applicants abroad to Wafid centres and allows a separate form where no accredited centre exists.",
      coverage="Bahrain LMRA has an official alternate non-Wafid form where there is no approved centre.")
bhr_keys="BHR_moh_program BHR_lmra_guidelines BHR_regulations_2022 BHR_lmra_health"
path("BHR","P2","Post-arrival expatriate health check for work and residence",P,N,bhr_keys,"yes","destination","exam_required;unfit_certificate;removal_deportation",
     "National examination rules explicitly name conditions. LMRA says employer must process deportation if the worker failed the required examination.",
     "The documented removal instruction is general and is not applied to HD by inference.")
for c in ("HIV","HBV","HCV"):
    e("BHR","P2","BHR_moh_program",c,"HIV, Hepatitis B and Hepatitis C positive expatriates are UNFIT to work in Bahrain",
       "Medical Examination Requirements, item 4", "unfit_certificate",scope="destination")
e("BHR","P2","BHR_moh_program","TB","Tuberculin skin test (for specific jobs, e.g. health care , kindergartens workers).",
   "Medical Examination Requirements, item 3","exam_required",kind="administrative_test",scope="destination",trigger="specific occupations")
e("BHR","P2","BHR_moh_program","SYPH","Blood investigations (HIV, HBsAg, HCV, VDRL/RPR) for all expatriates.",
   "Medical Examination Requirements, item 4; cross-check BHR Ministry of Health Expatriate Medical Checkup Guidelines, syphilis treatment section","exam_required",kind="administrative_test",scope="destination",
   note="Syphilis target corroborated by the same ministry's guideline. The test listing does not state a syphilis exclusion.")
e("BHR","P2","BHR_lmra_health","GENERIC_FITNESS","The employer must process the deportation of the expatriate employee immediately, if the employee had failed to pass the required medical examination.",
   "Service Conditions","removal_deportation",scope="destination")
saf("BHR","P2","BHR_moh_program",individual=yes,notes="Fitness Procedure: some expatriates may require further assessment or more investigations. No post-MDT or applicant appeal evidence identified in this provision.")
trig("BHR","P2","BHR_moh_program","occupation","Specific jobs, e.g. health care and kindergartens workers (tuberculin skin test).")
settle("BHR",P,H,"Pre-departure Wafid and destination checks are separate.",
       additional="Extended applicability and review search for the 2022 amendment and distinction between Wafid and domestic checks.")

source("CYP_single_permit_documents","Remunerated employment (single permit – GEN), forms", "Republic of Cyprus Migration Department", "official forms index")
source("CYP_single_permit_entry_doc","Supporting documents – entry permit, registration and temporary residence and work permit (general employment)", "Republic of Cyprus Migration Department", "official checklists",
       status="observed_on_official_website", notes="Original legacy Word document. Two checklists distinguish country of origin from post-arrival certification.")
source("CYP_single_permit_renew_doc","Supporting documents – renewal of the residence and work permit (general employment)", "Republic of Cyprus Migration Department", "official checklist",
       status="observed_on_official_website", notes="Complete document list has no medical examination or test. Health insurance or GeSY paperwork appears separately.")
cyp_key="CYP_single_permit_entry_doc"
cyp_words=[
    ("P1","General employment single permit: entry, first issue","origin","Entry checklist item 3",
    "Original results of medical blood tests to ensure that the third-country national does not suffer from hepatitis B and C, HIV, syphilis, and chest X-ray for tuberculosis, valid for up to 4 months from the issuing date, from the country of origin, duly certified."),
    ("P2","General employment single permit: residence/work registration after arrival","destination","After arrival checklist item 3",
    "Original results of medical blood tests demonstrating that the third-country national does not suffer from hepatitis B and C, HIV, syphilis, and chest X-ray for tuberculosis, duly certified by a Doctor in the Republic.")]
for pw, pname, scope, loc, phrase in cyp_words:
    path("CYP",pw,pname,P,N,cyp_key,"yes",scope,"exam_required","Official employment checklist requires medical tests for several named conditions.",
         "Renewal is a later stage with a separately published checklist and no medical test item.")
    trig("CYP",pw,cyp_key,"visa_class","Third-country national in general employment single permit (GEN).")
    for c in ("HBV","HCV","HIV","SYPH","TB"):
        e("CYP",pw,cyp_key,c,phrase,loc,"exam_required",kind="administrative_test",scope=scope)
path("CYP","P2-RENEW","General employment single permit: renewal",A,N,"CYP_single_permit_renew_doc","no","not_stated","unknown",
     "Complete official general-employment renewal checklist has no medical examination or tests.")
settle("CYP",P,N,"Named tests in both first-issue checklists; no medical item in the general renewal checklist.")
