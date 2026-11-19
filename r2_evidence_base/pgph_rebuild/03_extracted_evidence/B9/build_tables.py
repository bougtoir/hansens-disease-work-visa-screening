"""Build B9 extraction tables from raw ledgers + agent adjudications. Every source_wording_original is
verified verbatim (whitespace/quote-normalised) against the text extracted from the saved raw file."""
import csv, re, sys, datetime, hashlib, subprocess
from pathlib import Path
from bs4 import BeautifulSoup
D = Path(__file__).resolve().parent; RB = D.parents[1]; RAW = RB/"01_raw_search_logs"/"B9"
sys.path.insert(0, str(RB/"data")); import schema as S
NOW = "2026-11-19T09:10:00Z"; REV = "Devin B9 agent"
frame = {r["jurisdiction_id"]: r for r in csv.DictReader(open(RB/"data/JURISDICTION_FRAME.csv"))}
MINE = sorted(j for j, r in frame.items() if r["search_batch"].split()[0] == "B9")
def iso(ts): return f"{ts[0:4]}-{ts[4:6]}-{ts[6:8]}T{ts[9:11]}:{ts[11:13]}:{ts[13:15]}Z"
S1 = set("eCFR42CFR34 INA212_uscode eCFR22CFR41_108 IRPR_s30 IRPA_s38 PIC4005_extract PIC4007_extract ImmigrationAct2003_sherloc Code_T50_Ch01 Title13_PNCA ImmigrationAct2014 ImmigAct2010 ImmigAct2020_AG Ley25871_texto Ley370_cepal Ley21325_DO Decreto177 Res5477_2022 LOMH_reform_RO268 LOMH_2026_text ImmigrationAct_sela DL1350 Vreemdelingenwet1991 Dec394_2009 Ley18250 Lei13445_b Decreto9199_b".split())
S2 = set("CDC_TI_civil_hansens CDC_TI_panel_hansens USCIS_PM_v8pB_ch6 USCIS_PM_v8pB_ch3b PanelMemberGuide Form26 INZ1007 INZ1121 OM_A4_10 OM_A4_15 OM_A4_25 permit_to_work_form MFA_MedicalForm Immig_MedicalReport Immig_ResEmpChecklist MPMC100_TR MCIL_FEEP MOHA_landing_permit_req EntryPermitApp RevImmigReg MFA_WorkVisa MFA_ImmigProcedures2020".split())
LANG = dict(ARG="es", BOL="es", CHL="es", COL="es", ECU="es", PER="es", PRY="es", URY="es", VEN="es", BRA="pt", SUR="nl")
NOTE = {"WB_MigrationAct1978": "PacLII consolidated text retrieved via Wayback Machine id_ capture 20250130103048 (PacLII returned 403 directly)",
        "ImmigrationAct_sela": "Laws of Guyana Cap. 14:02 revised-edition text reproduced on sela.org; official mola.gov.gy copy not retrieved",
        "ImmigrationAct2003_sherloc": "Act text as hosted by UNODC SHERLOC", "MPMC100_TR": "Samoa MPMC form hosted on samoa.tradeportal.org (Government of Samoa trade portal)",
        "ICA_medical_exam_form": "URL of ICA medical examination form returned an HTML page, not the form -> form content inaccessible",
        "Lei13445_b": "browser capture returned Planalto portal home page (not the law text) -> law text inaccessible", "Decreto9199_b": "browser capture returned Planalto portal home page (not the decree text) -> inaccessible",
        "WB_PacLII_MigrationAct1978": "Wayback availability-API JSON (locator only, not a legal text)", "RevImmigReg": "scanned PDF without text layer; not read",
        "AUS_PIC4007_extract": ""}
BAD = {"ICA_medical_exam_form", "Lei13445_b", "Decreto9199_b", "WB_PacLII_MigrationAct1978", "RevImmigReg"}
# ---- SOURCE_REGISTER from FETCH_LEDGER
src, byshort, texts = [], {}, {}
for i, r in enumerate(csv.DictReader(open(RAW/"FETCH_LEDGER.csv")), 1):
    if r["local_path"] and not (RB.parent/r["local_path"]).exists():
        r["error"] = "file discarded by agent: saved during temp-file collision between parallel fetches (tool bug), content not verifiable; see WORKING_NOTES.md"; r["local_path"] = ""
    ok = bool(r["local_path"]); sid = f"SRC-B9-{r['iso3']}-{i:03d}"
    if ok: byshort[(r["iso3"], r["short"])] = sid
    tier = "S7" if (not ok or r["short"] in BAD) else ("S4" if r["short"].startswith("WB_") else "S1" if r["short"] in S1 else "S2" if r["short"] in S2 else "S3")
    note = NOTE.get(r["short"], "") or ("" if ok else f"fetch failed: HTTP {r['http_status'] or 'none'} {r['error'][:120]}".strip())
    src.append(dict(source_id=sid, jurisdiction_id="J-"+r["iso3"], title=r["short"], issuing_authority="see url domain", source_tier=tier,
        doc_type=("pdf" if r["local_path"].endswith(".pdf") else "html" if ok else ""), language=LANG.get(r["iso3"], "en"), url=r["url"], archive_url=(r["url"] if "web.archive.org" in r["url"] else ""),
        retrieved_utc=iso(r["retrieved_utc"]), local_path=r["local_path"], bytes=r["bytes"] if ok else "0", sha256=r["sha256"],
        effective_date="", status_in_force=("not_determined" if ok else "inaccessible"), redistribution="public official text; redistribution terms not verified",
        notes=(note + f" | fetch mode {r['mode']}, content_type {r['content_type']}").strip(" |")))
    if ok:
        p = RB.parent/r["local_path"]; data = p.read_bytes(); assert hashlib.sha256(data).hexdigest() == r["sha256"]
        t = subprocess.run(["pdftotext", "-raw", str(p), "-"], capture_output=True, text=True).stdout if data[:4] == b"%PDF" else BeautifulSoup(data, "html.parser").get_text(" ")
        texts[sid] = t
def nz(s): return re.sub(r"\s+", " ", s.replace("\u2019", "'").replace("\u2018", "'").replace("\u201c", '"').replace("\u201d", '"').replace("\u2014", "-").replace("\u2013", "-")).strip().lower()
NT = {k: nz(v) for k, v in texts.items()}
def SID(iso3, short): return byshort[(iso3, short)]
# ---- EVIDENCE
EV = []
def E(j, pw, short, prov, cond, ctype, orig, eng, named, effect, loc, tier, conf, note, clin="", safe="", trig="", scope="", other=""):
    sid = SID(j, short)
    if nz(orig) not in NT[sid]: raise SystemExit(f"NOT VERBATIM in {sid} {short}: {orig[:80]}")
    EV.append(dict(evidence_id=f"EV-B9-{j}-{len([e for e in EV if e['jurisdiction_id']=='J-'+j])+1:03d}", jurisdiction_id="J-"+j, pathway_id=f"PW-B9-{j}-{pw}", source_id=sid,
        provision_id=prov, exam_country_scope=scope, condition_code=cond, condition_other_text=other, criterion_type=ctype, source_wording_original=orig,
        source_wording_english=eng or orig, explicit_named=named, operative_effect=effect, authority_level=tier, effective_from="", effective_to="",
        clinical_state_rule=clin, safeguard_rule=safe, trigger_rule=trig, excerpt_locator=loc, evidence_tier=tier, confidence=conf, classification_note=note,
        reviewer=REV, record_created_at=NOW, modified_at="", supersedes="", reason_for_change="", provenance_class="new_collected"))
exec(open(D/"adjudications.py", encoding="utf-8").read())
def W(name, rows):
    cols = getattr(S, name)
    with open(D/f"{name}.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols); w.writeheader(); [w.writerow({c: r.get(c, "") for c in cols}) for r in rows]
W("SOURCE_REGISTER", src); W("EVIDENCE_ROWS", EV)
# ---- SEARCH_LOG from QUERY_LEDGER
okurls = {s["url"] for s in src if s["local_path"]}
sl, nq = [], {}
for q in csv.DictReader(open(RAW/"QUERY_LEDGER.csv")):
    urls = [u for u in q["result_urls"].split(";") if u]; opened = [u for u in urls if u in okurls]
    nq[q["jurisdiction_id"]] = nq.get(q["jurisdiction_id"], 0) + 1
    sl.append(dict(search_id=q["search_id"], jurisdiction_id=q["jurisdiction_id"], batch="B9", stage=q["stage"], language=q["language"], timestamp_utc=q["timestamp_utc"],
        query=q["query"], engine_or_database=q["engine"], result_url=";".join(urls), hit_title="see raw file", relevant_excerpt="",
        opened_full_text=("yes" if opened else "no"), access_status="results_returned", translation_method=("query written and results read directly by agent in " + q["language"]) if q["language"] != "en" else "n/a (English)",
        notes=f"raw results: 01_raw_search_logs/B9/{q['raw_file']}" + (f"; opened: {';'.join(opened)}" if opened else "")))
W("SEARCH_LOG", sl)
for n in ["POLICY_PATHWAYS", "CLINICAL_STATUS_RULES", "SAFEGUARDS", "NATIONALITY_ORIGIN_TRIGGERS", "OBSERVED_DENIALS", "POST_TREATMENT_RISK"]: W(n, globals()[n])
nsrc = {}; best = {}
for s in src:
    if s["local_path"] and s["source_tier"] != "S7":
        nsrc[s["jurisdiction_id"]] = nsrc.get(s["jurisdiction_id"], 0) + 1
        best[s["jurisdiction_id"]] = min(best.get(s["jurisdiction_id"], "S9"), s["source_tier"])
js = []
for j in MINE:
    p = [x for x in POLICY_PATHWAYS if x["jurisdiction_id"] == j]; meta = JMETA[j]
    order = ["documented_positive", "documented_absence", "unresolved"]; d = min((x["dss_status"] for x in p), key=order.index)
    h = "confirmed" if any(x["hd_named_status"] == "confirmed" for x in p) else ("not_established" if all(x["hd_named_status"] == "not_established" for x in p) else "cannot_determine")
    js.append(dict(jurisdiction_id=j, batch="B9", english_pass_done="yes", local_pass_done="yes", local_languages_used=meta[0], n_queries=nq.get(j, 0),
        n_sources_opened=nsrc.get(j, 0), best_source_tier=best.get(j, "S7"), dss_status=d, hd_named_status=h, unresolved_reason=meta[1], hours_spent="",
        notes=meta[2] + " | hours_spent not tracked per jurisdiction (batch run approx. 1.5 h wall-clock total); no budget extension used"))
W("JURISDICTION_STATUS", js)
print("rows: src", len(src), "ev", len(EV), "search", len(sl), "pathways", len(POLICY_PATHWAYS))
