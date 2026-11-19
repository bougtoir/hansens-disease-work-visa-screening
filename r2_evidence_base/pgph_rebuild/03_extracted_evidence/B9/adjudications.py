# B9 adjudications (executed by build_tables.py; E() verifies each quote verbatim against the saved raw file text)
# ---------------- USA
E("USA","P2","eCFR42CFR34","42 CFR 34.2(b)","HD","diagnostic","Hansen's disease, infectious","","yes","listed as communicable disease of public health significance (Class A ground via INA 212(a)(1)(A)(i))","34.2 Definitions, (b) Communicable disease of public health significance","S1","high","",scope="panel physician abroad / civil surgeon in US")
E("USA","P2","eCFR42CFR34","42 CFR 34.2(b)","SYPH","diagnostic","Syphilis, infectious","","yes","listed communicable disease of public health significance","34.2(b)","S1","high","")
E("USA","P2","eCFR42CFR34","42 CFR 34.2(b)","TB","diagnostic","Tuberculosis, active","","yes","listed communicable disease of public health significance","34.2(b)","S1","high","")
E("USA","P2","eCFR42CFR34","42 CFR 34.2(b)","STI_OTHER","diagnostic","Gonorrhea","","yes","listed communicable disease of public health significance","34.2(b)","S1","high","",other="gonorrhea")
E("USA","P2","CDC_TI_civil_hansens","CDC Technical Instructions for Civil Surgeons - Hansen's Disease","HD","diagnostic","Patients with multibacillary or paucibacillary Hansen's disease are considered to be Class A","","yes","untreated HD = Class A (inadmissible absent waiver)","CDC civil surgeon TI, Hansen's disease page","S2","high","",clin="untreated -> Class A; >=7 days treatment -> may be Class B")
E("USA","P2","CDC_TI_civil_hansens","CDC Technical Instructions for Civil Surgeons - Hansen's Disease","HD","diagnostic","after 7 days of treatment, they may be re-classified as Class B","","yes","treatment of >=7 days permits reclassification from Class A to Class B","CDC civil surgeon TI, Hansen's disease page","S2","high","Class B is a noted condition, not an inadmissibility ground",clin="on_treatment >=7 days -> Class B")
E("USA","P2","CDC_TI_panel_hansens","CDC Technical Instructions for Panel Physicians - Hansen's Disease","HD","diagnostic","Hansen's disease","","yes","panel physicians abroad evaluate HD; untreated Class A, treated/partially treated Class B","CDC panel physician TI, Hansen's disease page","S2","high","full classification text in saved file",scope="panel physician abroad")
E("USA","P2","USCIS_PM_v8pB_ch6","USCIS Policy Manual Vol 8 Part B Ch 6","HD","diagnostic","Hansen's Disease (Leprosy), infectious","","yes","listed communicable disease of public health significance","USCIS PM Vol.8 Pt.B Ch.6","S2","high","")
E("USA","P1","INA212_uscode","INA 212(a)(1)(A)(i) / 8 USC 1182","GENERIC_FITNESS","diagnostic","to have a communicable disease of public health significance","","no","health-related inadmissibility ground (diseases defined by HHS rule 42 CFR 34.2)","8 USC 1182(a)(1)(A)(i)","S1","high","statute itself names no disease; named list is in 42 CFR 34.2")
E("USA","P1","eCFR22CFR41_108","22 CFR 41.108","GENERIC_FITNESS","administrative_test","medical examination","","no","nonimmigrant (incl. temporary work) visa applicants examined only in listed cases (K visas, medical treatment, or where consular officer has reason to believe exam might disclose ineligibility)","22 CFR 41.108(a)","S1","medium","routine pre-visa exam not required for temporary workers per 41.108; when performed, 42 CFR 34 applies")
# ---------------- CAN
E("CAN","P1","PanelMemberGuide","IRCC Panel Members' Guide - Physical examination table, Skin and lymph nodes","HD","administrative_test","Screen for skin cancer, leprosy, surgical scars, tattoos and piercings.","","yes","physical examination instruction; no consequence or treatment rule stated in this row","Panel Members' Guide, Table: Physical examination, 'Skin and lymph nodes'","S2","high","HD named as screening target; outcome governed generically by IRPA s38",scope="panel physician (IME)")
E("CAN","P1","PanelMemberGuide","IRCC Panel Members' Guide - Syphilis serology","SYPH","administrative_test","Syphilis screening is required for all clients 15 years of age and older undergoing an IME.","","yes","mandatory test at IME","Panel Members' Guide, 'Syphilis serology'","S2","high","",scope="panel physician (IME)")
E("CAN","P1","IRPA_s38","IRPA s.38(1)","GENERIC_FITNESS","diagnostic","danger to public health","","no","health inadmissibility ground","IRPA s.38(1)(a)","S1","high","no disease named in statute")
# ---------------- AUS
E("AUS","P1","PIC4005_extract","Migration Regulations 1994 Sch 4 PIC 4005(1)(a)","TB","diagnostic","is free from tuberculosis","","yes","health criterion: applicant must be free from TB","PIC 4005(1)(a)","S1","high","")
E("AUS","P1","PIC4007_extract","Migration Regulations 1994 Sch 4 PIC 4007","GENERIC_FITNESS","diagnostic","a threat to public health in Australia or a danger to the Australian community","","no","health criterion (public health/community danger); waiver under subclause (2) for cost limb","PIC 4007(1)","S1","high","")
# ---------------- NZL
E("NZL","P1","INZ1007","INZ 1007 General Medical Certificate, Section E2","GENERIC_FITNESS","administrative_test","Standard (compulsory) blood tests for all applicants 15 years of age and over","","no","compulsory blood test panel (items below)","INZ1007 Section E2","S2","high","header of compulsory test list",scope="panel physician")
E("NZL","P1","INZ1007","INZ 1007 Section E2","SYPH","administrative_test","Syphilis test (VDRL or RPR)","","yes","compulsory test (15+)","INZ1007 Section E2","S2","high","")
E("NZL","P1","INZ1007","INZ 1007 Section E2","HBV","administrative_test","Hepatitis B surface antigen (Hep B sAg)","","yes","compulsory test (15+)","INZ1007 Section E2","S2","high","")
E("NZL","P1","INZ1007","INZ 1007 Section E2","HCV","administrative_test","Hepatitis C antibody test","","yes","compulsory test (15+)","INZ1007 Section E2","S2","high","")
E("NZL","P1","INZ1007","INZ 1007 Section B5","HIV","administrative_test","Have you ever had an abnormal or reactive HIV blood test?","","yes","self-declaration question; HIV also in compulsory test list","INZ1007 Section B5","S2","high","")
E("NZL","P2","OM_A4_10","INZ Operational Manual A4.10.1","TB","diagnostic","MDR-TB or XDR-TB","","yes","listed as condition indicating applicant may not have acceptable standard of health (residence) unless cleared by NZ specialist","A4.10.1","S2","high","",clin="history/treatment of MDR/XDR-TB listed unless specialist clearance")
E("NZL","P1","OM_A4_15","INZ Operational Manual A4.15","GENERIC_FITNESS","diagnostic","danger to public health","","no","temporary entry acceptable standard of health","A4.15","S2","high","")
# ---------------- PNG
E("PNG","P1","WB_MigrationAct1978","Migration Act 1978, 'Power to refuse entry' (1)(b)(ii)","GENERIC_FITNESS","diagnostic","from a disease which would make his presence in the country a danger to the community","","no","refusal of entry on officer's opinion","Migration Act 1978 'POWER TO REFUSE ENTRY' (1)(b)(ii)","S4","medium","archived PacLII text")
E("PNG","P1","WB_MigrationAct1978","Migration Act 1978, 'Power to refuse entry' (1)(b)(i)","MENTAL","diagnostic","from a mental illness","","ambiguous","refusal of entry on officer's opinion","'POWER TO REFUSE ENTRY' (1)(b)(i)","S4","medium","category, not a specific diagnosis")
# ---------------- FJI
E("FJI","P1","ImmigrationAct2003_sherloc","Immigration Act 2003 s13(2)(d)(ii)","GENERIC_FITNESS","diagnostic","suffering from a contagious or infectious disease, or from a mental disease or disorder, that makes the person's presence in the Fiji Islands dangerous to the community","","no","prohibited immigrant","s13(2)(d)(ii)","S1","high","no disease named")
E("FJI","P1","permit_to_work_form","Fiji Immigration Department permit application form, Note 9","GENERIC_FITNESS","administrative_test","Medical reports, which are less than 3 months old, must be submitted by the applicant","","no","medical report required; content of Department's medical report form not retrieved","Form notes, item 9","S2","high","")
# ---------------- FSM
E("FSM","P1","Code_T50_Ch01","FSM Code T50 §107(1)(i)","GENERIC_FITNESS","diagnostic","carrying a serious communicable disease","","no","ground to deny entry/permit or deport","50 FSMC 107(1)(i)","S1","high","no disease named")
E("FSM","P1","Code_T50_Ch01","FSM Code T50 §107(1)(h)","SUBST_DEP","behavioral","addiction to the use of narcotic drugs","","yes","ground to deny entry/permit or deport","50 FSMC 107(1)(h)","S1","high","behavioral condition; not counted as disease-specific screening")
# ---------------- KIR
E("KIR","P3","MFA_MedicalForm","Kiribati MFA medical form, B1","TB","administrative_test","Have you ever been diagnosed with Tuberculosis (TB)?","","yes","history question on medical form","Medical form B1","S2","medium","applicability of this form to work visas not established from sources read")
E("KIR","P3","MFA_MedicalForm","Kiribati MFA medical form, B5","HIV","administrative_test","Have you ever had an abnormal or reactive HIV blood test?","","yes","history question","Medical form B5","S2","medium","applicability not established")
# ---------------- PLW
E("PLW","P1","Title13_PNCA","13 PNCA §1362(a)","GENERIC_FITNESS","administrative_test","a certificate of freedom from communicable diseases","","no","nonresident workers must hold certificate (issued <=30 days before entry by physician in country of origin)","13 PNCA §1362","S1","high","no disease named",scope="country of origin physician")
E("PLW","P1","Title13_PNCA","13 PNCA §1364(c)","GENERIC_FITNESS","diagnostic","Such notification shall be made upon any positive finding of any contagious disease, including any social disease.","","no","post-arrival exam finding triggers notification to President (possible revocation §1365)","13 PNCA §1364(c)","S1","high","'social disease' not a named disease")
E("PLW","P1","Title13_PNCA","13 PNCA §1364(b)","GENERIC_FITNESS","diagnostic","No such notification shall be made if the continued presence of the entry permit holder or family member in the Republic would not, in all medical probability, result in substantial danger to the health of the inhabitants of the Republic","","no","individualised medical-probability test","13 PNCA §1364(b)","S1","high","",safe="individual assessment")
# ---------------- VUT
E("VUT","P1","ImmigAct2010","Immigration Act No.17 of 2010, visa criteria (d)","GENERIC_FITNESS","diagnostic","is not suffering from a contagious or other disease, or a mental condition, which makes his or her presence in Vanuatu a risk to the health of the community in Vanuatu","","no","visa grant criterion","Immigration Act 2010, visa criteria para (d)","S1","high","")
E("VUT","P1","Immig_ResEmpChecklist","Residency visa - employee checklist item 6","GENERIC_FITNESS","administrative_test","Recent Medical Certificate","","no","required document","Checklist item 6","S2","high","links pathway to Medical_Report form")
E("VUT","P1","Immig_MedicalReport","Vanuatu Immigration Medical Report, item 23","HIV","administrative_test","Human immunodeficiency virus result","","yes","test result field","Medical Report item 23","S2","medium","form identity as the 'Medical Certificate' of the checklist inferred from same issuer")
E("VUT","P1","Immig_MedicalReport","Vanuatu Immigration Medical Report, item 22","HBV","administrative_test","Hepatitis B antigen test result","","yes","test result field","item 22","S2","medium","")
E("VUT","P1","Immig_MedicalReport","Vanuatu Immigration Medical Report, item 20","SYPH","administrative_test","VDRL test result— only in Clinically indicated","","yes","conditional test","item 20","S2","medium","")
E("VUT","P1","Immig_MedicalReport","Vanuatu Immigration Medical Report, history c)","TB","administrative_test","Tuberculosis?","","yes","history question (item c)","Medical history item c) Tuberculosis?","S2","medium","chest x-ray item 21 not counted as TB screening")
# ---------------- WSM
E("WSM","P1","MPMC100_TR","MPMC100 Temporary Residence application, Q16","HD","administrative_test","Do you or your spouse suffer from any communicable illness or disease (eg HIV, T.B., Leprosy, STD)?","","yes","self-declaration; no consequence stated on form","MPMC100 question 16","S2","high","HD named as example in declaration question")
E("WSM","P1","MPMC100_TR","MPMC100 Q16","HIV","administrative_test","eg HIV, T.B., Leprosy, STD","","yes","self-declaration","MPMC100 question 16","S2","high","")
E("WSM","P1","MPMC100_TR","MPMC100 Q16","TB","administrative_test","eg HIV, T.B., Leprosy, STD","","yes","self-declaration","MPMC100 question 16","S2","high","")
E("WSM","P1","ImmigAct2020_AG","Immigration Act 2020 s27(1)(d)","GENERIC_FITNESS","diagnostic","does not suffer from a physical or psychological condition that poses a threat to the health or safety of the community","","no","required health criterion","s27(1)(d)","S1","high","")
# ---------------- GUY
E("GUY","P1","ImmigrationAct_sela","Immigration Act Cap. 14:02 s5(1)(b)","HD","diagnostic","is suffering from any variety of leprosy, or from a communicable disease, as certified by a Government Medical Officer","","yes","prohibited immigrant","Cap. 14:02 prohibited-immigrant section, para (b)","S1","high","no treatment-state distinction; 'suffering from' wording does not establish exclusion of cured persons",clin="not addressed")
E("GUY","P1","MOHA_landing_permit_req","MoHA requirements for landing permit for employment, item 10","GENERIC_FITNESS","administrative_test","Medical Report for applicant","","no","required document; report content not retrieved","item 10","S2","high","")
# ---------------- URY
E("URY","P2","Dec394_2009","Decreto 394/2009 (residencia), literal C) Salud","GENERIC_FITNESS","administrative_test","Deberá acreditarse mediante Carné de Salud que el titular está apto para residir en el país","Fitness to reside must be shown by a Health Card (Carné de Salud)","no","health card required","Decreto 394/2009 art. on requisitos, C) Salud","S1","high","no disease named; Carné de Salud content not retrieved")

def PW(j, pw, ptype, name, exam, loc, srcs, dss, dssb, hd, hdb, cons, notes=""):
    return dict(pathway_id=f"PW-B9-{j}-{pw}", jurisdiction_id="J-"+j, pathway_type=ptype, pathway_name=name, medical_exam_required=exam, exam_location=loc,
        governing_source_ids=";".join(SID(j, s) for s in srcs), dss_status=dss, dss_basis=dssb, hd_named_status=hd, hd_basis=hdb, consequence_values=cons, notes=notes)
U = "unresolved"; P = "documented_positive"; A = "documented_absence"; C = "confirmed"; N = "not_established"; X = "cannot_determine"
POLICY_PATHWAYS = [
 PW("USA","P2","P2_residence_for_workers","Employment-based immigrant visa / adjustment of status","yes","panel physician abroad; civil surgeon in US",["eCFR42CFR34","CDC_TI_civil_hansens","CDC_TI_panel_hansens","USCIS_PM_v8pB_ch6"],P,"42 CFR 34.2(b) names HD, TB, syphilis, gonorrhea",C,"42 CFR 34.2(b)(4) 'Hansen's disease, infectious'; CDC TIs","Class A inadmissibility (INA 212(a)(1)(A)(i)); Class B noted"),
 PW("USA","P1","P1_work","Nonimmigrant temporary work visas (e.g. H, L)","conditional","consular-directed panel physician",["eCFR22CFR41_108","INA212_uscode","eCFR42CFR34"],P,"22 CFR 41.108: exam only in listed cases; when performed, 42 CFR 34 named list applies",C,"42 CFR 34.2(b) applies to all visa applicants examined (34.1)","inadmissibility if Class A","exam not routine for temporary workers"),
 PW("CAN","P1","P1_work","Work permit (IME required for public-health-sensitive occupations / designated-country residence / >6 months)","conditional","panel physician",["IRPR_s30","IRPA_s38","PanelMemberGuide"],P,"Panel Members' Guide: syphilis screening required 15+; HIV testing in IME",C,"Panel Members' Guide physical exam: 'Screen for ... leprosy'","IRPA s38 health inadmissibility (generic)","HD named as screening item only; no HD-specific consequence located"),
 PW("AUS","P1","P1_work","Temporary/permanent visas subject to PIC 4005/4007 health criteria (subclass mapping not extracted)","yes","panel physician",["PIC4005_extract","PIC4007_extract","Form26","HA_health_page"],P,"PIC 4005(1)(a) 'free from tuberculosis'",N,"HD not named in PIC 4005/4007 extracts, Form 26 or health page read","visa refusal if health criterion not met; PIC 4007 waiver"),
 PW("NZL","P1","P1_work","Temporary entry (work) visas >12 months","yes","panel physician",["OM_A4_15","OM_A4_25","INZ1007"],P,"INZ1007 compulsory tests: syphilis, Hep B, Hep C, HIV",N,"HD not named in INZ1007, A4.10, A4.15, A4.25","acceptable standard of health; medical waiver"),
 PW("NZL","P2","P2_residence_for_workers","Residence class visas (skilled migrant etc.)","yes","panel physician",["OM_A4_10","INZ1007"],P,"A4.10.1 names MDR/XDR-TB; INZ1007 test panel",N,"HD not named in A4.10 or INZ1007","not acceptable standard of health"),
 PW("PNG","P1","P1_work","Employment visa / work permit","conditional","officer-required examination",["WB_MigrationAct1978"],U,"statute generic (disease/mental illness); ICA medical form inaccessible",X,"medical form not obtained","refusal of entry"),
 PW("FJI","P1","P1_work","Work permit","yes","not stated",["ImmigrationAct2003_sherloc","permit_to_work_form"],U,"statute generic; Department medical report form not retrieved",X,"medical report form content not obtained","prohibited immigrant"),
 PW("FSM","P1","P1_work","Alien worker entry permit","not_determined","",["Code_T50_Ch01"],U,"§107 generic 'serious communicable disease'; immigration regulations PDF scanned/not read",X,"regulations not read","deny/revoke permit; deportation"),
 PW("KIR","P1","P1_work","Work visa","not_determined","",["MFA_WorkVisa","MFA_ImmigProcedures2020"],U,"MFA medical form names TB/HIV but its applicability to work visa not established; Act (natlex) 403",X,"no applicable source naming or excluding HD located",""),
 PW("KIR","P3","P3_other","MFA medical form (pathway applicability unknown)","not_determined","",["MFA_MedicalForm"],U,"form names TB, HIV, Hep B/C history questions; applicability not established",X,"",""),
 PW("MHL","P1","P1_work","Work permit","not_determined","",["RMI_live_work_permits"],U,"official page lists no health terms; legislation inaccessible (natlex 403, rmiparliament 502)",X,"",""),
 PW("NRU","P1","P1_work","Work / business visa","not_determined","",["ImmigrationAct2014","Justice_immig_div"],U,"Act: officer may require medical exam; no named condition; regulations/forms not read",X,"",""),
 PW("PLW","P1","P1_work","Nonresident worker entry permit","yes","country of origin + Bureau of Public Health on arrival",["Title13_PNCA"],U,"statute generic (communicable/contagious/social disease); DOL Rules 2019 inaccessible (403)",X,"implementing rules not obtained","revocation of entry permit"),
 PW("SLB","P1","P1_work","Working resident visa","not_determined","",["Immig_WRV_long"],U,"official page no health terms located; Immigration Act 2012 (natlex) 403",X,"",""),
 PW("TON","P1","P1_work","Employment visa","not_determined","",["Embassy_employment_visa"],U,"embassy page no health terms; Immigration Regulations fetch failed",X,"",""),
 PW("TUV","P1","P1_work","Work permit / entry","not_determined","",[],U,"tuvalu-legislation.tv returned 502; no source read",X,"",""),
 PW("VUT","P1","P1_work","Residency visa - employee","yes","not stated",["ImmigAct2010","Immig_ResEmpChecklist","Immig_MedicalReport"],P,"Medical Report form: HIV, Hep B antigen, VDRL (clinically indicated)",N,"HD not named in Medical Report form or Act","visa refusal"),
 PW("WSM","P1","P1_work","Temporary Resident Permit for Employment","yes","not stated (MPMC102/103)",["MPMC100_TR","ImmigAct2020_AG"],P,"MPMC100 Q16 names HIV, TB, leprosy, STD",C,"MPMC100 Q16 'Leprosy'","not stated on form; s27 health criteria"),
 PW("GUY","P1","P1_work","Employment visa / landing permit for employment","yes","not stated",["ImmigrationAct_sela","MOHA_landing_permit_req","ISS_employment_visa"],P,"Cap 14:02 prohibited immigrant: 'any variety of leprosy'",C,"Cap 14:02 s5(1)(b)","prohibited immigrant"),
 PW("ARG","P1","P1_work","Residencia temporaria - trabajador migrante","no","",["Ley25871_texto","Serv_res_temp_trabajador"],A,"Ley 25.871 art. 29 impediments contain no health ground; official requirement list names no medical document",N,"",""),
 PW("BOL","P1","P1_work","Permanencia temporal por trabajo","not_determined","",["Ley370_cepal"],U,"Ley 370 no health ground located; reglamento DS 1923 inaccessible",X,"",""),
 PW("BRA","P1","P1_work","Autorização de residência para trabalho","not_determined","",["MJ_res_laboral"],U,"Lei 13.445 / Decreto 9.199 text inaccessible (429 / portal page)",X,"",""),
 PW("CHL","P1","P1_work","Residencia temporal - actividades remuneradas","not_determined","",["Ley21325_DO","SM_actividades_remuneradas"],U,"no health ground or medical document located, but identity of DO text (law vs reglamento) and full requirement list not confirmed",X,"",""),
 PW("COL","P1","P1_work","Visa M / V (trabajo)","no","",["Res5477_2022"],A,"Res. 5477/2022 full text: requirements include health insurance only; no medical examination or named condition",N,"",""),
 PW("ECU","P1","P1_work","Residencia temporal de trabajo","not_determined","",["LOMH_2026_text","Tramite_res_temp_trabajo"],U,"LOMH art.137 read only partially; tramite page no medical terms",X,"",""),
 PW("PER","P1","P1_work","Calidad migratoria trabajador residente","no","",["gobpe_trabajador_residente","DL1350"],A,"official requirement list (gob.pe) contains no medical document; DL 1350 has generic sanitary-authority entry restriction only",N,"","generic: sanitary authority may restrict entry"),
 PW("PRY","P1","P1_work","Residencia temporal","not_determined","",["DGM_residencia_temporal"],U,"requirement list read to item 10 without health document, end of list not confirmed",X,"",""),
 PW("SUR","P1","P1_work","Verblijfsvergunning / werkvergunning","not_determined","",["Vreemdelingenwet1991","WV_folder_EN","VZ_FAQ"],U,"no health terms located in Vreemdelingenwet 1991 text or folders; implementing decree not read",X,"",""),
 PW("URY","P2","P2_residence_for_workers","Residencia legal temporaria / permanente","yes","Uruguay (MSP-authorised providers)",["Dec394_2009","GUB_res_temporaria"],U,"Carné de Salud (fitness) required; tests not named in sources read",X,"Carné de Salud content not obtained","residence requirement"),
 PW("VEN","P1","P1_work","Visa transeúnte laboral (TL)","not_determined","",[],U,"SAIME portal DNS failure; OAS copy of Ley 37.944 503",X,"",""),
]
def CS(j, pw, a, o, m, c, r, srcs, n=""): return dict(jurisdiction_id="J-"+j, pathway_id=f"PW-B9-{j}-{pw}", active_untreated=a, on_treatment=o, mdt_completed=m, cured=c, residual_sequelae=r, source_ids=";".join(SID(j,s) for s in srcs), notes=n)
na = "not_addressed"
CLINICAL_STATUS_RULES = [
 CS("USA","P2","mentioned_excluded","mentioned_conditional","mentioned_exempt",na,na,["CDC_TI_civil_hansens","CDC_TI_panel_hansens"],"untreated=Class A; >=7 days therapy -> Class B; treated/partially treated = Class B (not inadmissible). 'cured' not separately addressed"),
 CS("USA","P1","mentioned_excluded","mentioned_conditional","mentioned_exempt",na,na,["CDC_TI_panel_hansens"],"same TI applies when NIV exam is performed"),
 CS("CAN","P1",na,na,na,na,na,["PanelMemberGuide"],"leprosy screening item; no treatment-state rule located"),
 CS("WSM","P1",na,na,na,na,na,["MPMC100_TR"],"declaration only"),
 CS("GUY","P1",na,na,na,na,na,["ImmigrationAct_sela"],"'suffering from any variety of leprosy' - no treatment-state distinction; ambiguity does not establish exclusion of cured persons"),
]
def SG(j, pw, d, i, r, w, a, srcs, n=""): return dict(jurisdiction_id="J-"+j, pathway_id=f"PW-B9-{j}-{pw}", discretion=d, individual_assessment=i, retest=r, waiver=w, appeal_review=a, source_ids=";".join(SID(j,s) for s in srcs), notes=n)
ea, ns = "explicit_available", "no_statement_located"
SAFEGUARDS = [
 SG("USA","P2",ea,ea,ns,ea,ea,["INA212_uscode","eCFR42CFR34","CDC_TI_panel_hansens"],"INA 212(g) waiver; 42 CFR 34.8 reexamination by medical board on appeal; CDC panel TI Class A waiver for first-week-treatment applicants"),
 SG("USA","P1",ea,ns,ns,ea,ea,["INA212_uscode","eCFR42CFR34"],"waiver per INA 212 (NIV 212(d)(3) not extracted); 34.8 reexamination"),
 SG("CAN","P1",ns,ns,ns,ns,ns,["IRPA_s38"],"not extracted"),
 SG("AUS","P1",ns,ns,ns,ea,ns,["PIC4007_extract"],"PIC 4007(2) waiver of cost limb"),
 SG("NZL","P1",ns,ns,ns,ns,ns,["OM_A4_15"],"medical waiver sections (A4.60/A4.65) not retrieved"),
 SG("NZL","P2",ns,ea,ns,ns,ns,["OM_A4_10"],"medical assessor referral"),
 SG("PLW","P1",ns,ea,ns,ns,ns,["Title13_PNCA"],"§1364(b) medical-probability individual test"),
 SG("WSM","P1",ns,ns,ns,ns,ns,["ImmigAct2020_AG"],"s23(5) Minister's prohibited-immigrant declaration not reviewable (not specific to health)"),
 SG("GUY","P1",ns,ns,ns,ns,ns,["ImmigrationAct_sela"],"Minister's discretionary exemptions in Act not extracted"),
]
def TR(j, pw, t, v, srcs, n=""): return dict(jurisdiction_id="J-"+j, pathway_id=f"PW-B9-{j}-{pw}", trigger_type=t, trigger_value=v, source_ids=";".join(SID(j,s) for s in srcs), notes=n)
NATIONALITY_ORIGIN_TRIGGERS = [
 TR("CAN","P1","prior_residence","residence >=6 months in designated country/territory (higher communicable disease incidence)",["IRPR_s30"],"triggers IME for work permits >6 months / agricultural work"),
 TR("PNG","P1","nationality","Philippines: medical, radiology and NBI reports validated by PNG Embassy Manila",["ICA_medical_character_r"],"document-validation trigger, not a disease test"),
 TR("PLW","P1","country_of_origin","certificate issued by physician licensed in worker's country of origin",["Title13_PNCA"],"exam location rule"),
]
OBSERVED_DENIALS = []
def PR(j, pw, s, m, srcs, n=""): return dict(jurisdiction_id="J-"+j, pathway_id=f"PW-B9-{j}-{pw}", risk_status=s, mechanism=m, source_ids=";".join(SID(j,x) for x in srcs), notes=n)
POST_TREATMENT_RISK = [
 PR("USA","P2","treatment_distinction_exempts_or_conditional","treated/partially treated HD = Class B; >=7 days therapy reclassifies from Class A",["CDC_TI_civil_hansens","CDC_TI_panel_hansens"]),
 PR("USA","P1","treatment_distinction_exempts_or_conditional","same CDC TI when exam performed",["CDC_TI_panel_hansens"]),
 PR("CAN","P1","cannot_determine","leprosy named as screening item; no consequence or treatment rule located",["PanelMemberGuide"]),
 PR("WSM","P1","cannot_determine","leprosy named in declaration question; no consequence stated",["MPMC100_TR"]),
 PR("GUY","P1","no_treatment_distinction_named_condition","'suffering from any variety of leprosy' prohibited immigrant; no treatment-state wording",["ImmigrationAct_sela"],"not an explicit post-treatment exclusion"),
 PR("AUS","P1","not_applicable_hd_not_named","",["PIC4005_extract"]),
 PR("NZL","P1","not_applicable_hd_not_named","",["INZ1007"]), PR("NZL","P2","not_applicable_hd_not_named","",["OM_A4_10"]),
 PR("VUT","P1","not_applicable_hd_not_named","",["Immig_MedicalReport"]),
 PR("ARG","P1","not_applicable_hd_not_named","",["Ley25871_texto"]), PR("COL","P1","not_applicable_hd_not_named","",["Res5477_2022"]), PR("PER","P1","not_applicable_hd_not_named","",["gobpe_trabajador_residente"]),
]
for p in POLICY_PATHWAYS:
    if p["hd_named_status"] == "cannot_determine" and not any(r["pathway_id"] == p["pathway_id"] for r in POST_TREATMENT_RISK):
        POST_TREATMENT_RISK.append(dict(jurisdiction_id=p["jurisdiction_id"], pathway_id=p["pathway_id"], risk_status="cannot_determine", mechanism="", source_ids=p["governing_source_ids"], notes="HD status cannot be determined"))
EN = "English (official legal language); local-language search not applicable"
JMETA = {
 "J-USA": ("en", "", "P1/P2 English; P3 applicability via 22 CFR 41.108"), "J-CAN": ("en;fr", "", "French official-language search run"),
 "J-AUS": ("en", "", ""), "J-NZL": ("en", "", "Māori local-language search not run (requirements published in English)"),
 "J-PNG": ("en", "ICA medical examination form inaccessible; PacLII 403 (Act via Wayback)", "Tok Pisin search not run"),
 "J-FJI": ("en", "Department medical report form content not retrieved", "Fijian/Hindi search not run"),
 "J-FSM": ("en", "Immigration regulations PDF scanned (no text layer); no medical requirement located", EN),
 "J-KIR": ("en", "Immigration Act 2019 (natlex) 403; medical form applicability unknown", EN),
 "J-MHL": ("en", "legislation sites 403/502", EN), "J-NRU": ("en", "regulations/forms not located", EN),
 "J-PLW": ("en", "Division of Labor Rules 2019 inaccessible (403); statute generic", EN), "J-SLB": ("en", "Immigration Act 2012 (natlex) 403", EN),
 "J-TON": ("en", "Immigration Regulations fetch failed (000)", EN + "; Tongan search not run"), "J-TUV": ("en", "tuvalu-legislation.tv 502; no source read", EN),
 "J-VUT": ("en;fr", "", "French query run"), "J-WSM": ("en", "", EN + "; Samoan search not run"),
 "J-GUY": ("en", "", ""), "J-ARG": ("es", "", ""), "J-BOL": ("es", "reglamento DS 1923 inaccessible (000)", ""),
 "J-BRA": ("pt", "planalto.gov.br 429 / portal page instead of law text", ""), "J-CHL": ("es", "law text identity and full requirement list not confirmed", ""),
 "J-COL": ("es", "", ""), "J-ECU": ("es", "LOMH art.137 only partially read within budget", ""), "J-PER": ("es", "", ""),
 "J-PRY": ("es", "end of requirement list not confirmed", ""), "J-SUR": ("nl;en", "implementing decree not read", ""),
 "J-URY": ("es", "Carné de Salud content not obtained", ""), "J-VEN": ("es", "SAIME DNS failure; OAS 503", ""),
}
