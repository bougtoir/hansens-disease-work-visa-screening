import collections
import csv
import datetime as dt
import hashlib
import json
import re
import sys
from pathlib import Path
from urllib.parse import urlsplit

sys.dont_write_bytecode = True
import importlib.util
root = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("schema", root / "data/schema.py")
schema = importlib.util.module_from_spec(spec)
spec.loader.exec_module(schema)
logs = root / "01_raw_search_logs/B3"
output = Path(__file__).parent
data = {n: [] for n in schema.TABLES}
queries = sorted([(p, json.loads(p.read_text())) for p in logs.glob("*.json") if not p.name.startswith("fetch_")], key=lambda q: (q[1]["at"], q[0].name))
fetches = {p.stem[6:]: json.loads(p.read_text()) for p in sorted(logs.glob("fetch_*.json"))}
checked = {}
statuses = {}
# Fixed at the close of the batch extraction so a new checkout reproduces the CSVs.
adjudicated = dt.datetime.fromisoformat("2026-11-19T09:10:05+00:00")
P = "documented_positive"
A = "documented_absence"
U = "unresolved"
H = "confirmed"
N = "not_established"
C = "cannot_determine"
ns = "no_statement_located"
yes = "explicit_available"
na = "not_addressed"
source_counts = collections.Counter()
e_counts = collections.Counter()

def row(table, **values):
    assert set(values) <= set(schema.TABLES[table]), (table, set(values) - set(schema.TABLES[table]))
    data[table].append(values)

def source(key, title, authority, doc, *, date="", status="currentness_not_independently_established", notes="", tier=None, language=None, archive="", jurisdiction=None):
    if key in checked:
        return checked[key]["source_id"]
    f = fetches[key]
    p = root / f["file"]
    sha = hashlib.sha256(p.read_bytes()).hexdigest()
    assert sha == f["sha256"] and p.stat().st_size == f["bytes"], key
    j = jurisdiction or f["jur"]
    source_counts[j] += 1
    record = dict(source_id=f"S-{j}-{source_counts[j]:03}", jurisdiction_id="J-"+j, title=title,
                  issuing_authority=authority, source_tier=tier or f["tier"], doc_type=doc, language=language or f["lang"],
                  url=f["url"], archive_url=archive, retrieved_utc=f["retrieved"],
                  local_path=str(p.relative_to(root.parent)), bytes=f["bytes"], sha256=sha, effective_date=date,
                  status_in_force=status, redistribution="local_only", notes=notes)
    checked[key] = record
    row("SOURCE_REGISTER", **record)
    return record["source_id"]

def sr(keys):
    return ";".join(checked[k]["source_id"] for k in keys.split() if k)

def path(j, num, name, dss, hd, keys, exam, loc, consequence, basis, notes="", *, hb=""):
    row("POLICY_PATHWAYS", pathway_id=f"PW-{j}-{num}", jurisdiction_id="J-"+j,
        pathway_type={"P1":"P1_work","P2":"P2_residence_for_workers","P3":"P3_other"}[num.split("-")[0]], pathway_name=name,
        medical_exam_required=exam, exam_location=loc, governing_source_ids=sr(keys), dss_status=dss, dss_basis=basis,
        hd_named_status=hd, hd_basis=hb or ("Explicitly named in a cited primary source." if hd==H else
        "No explicit wording in the sources for this pathway." if hd==N else "Available evidence cannot settle explicit inclusion."),
        consequence_values=consequence, notes=notes)

def e(j, pw, key, code, original, locator, effect, *, en="", other="", kind="diagnostic", scope="not_stated", note="", confidence="high", trigger="", safeguard="", clinical="", named=None):
    s = checked[key]
    e_counts[j] += 1
    row("EVIDENCE_ROWS", evidence_id=f"E-{j}-{e_counts[j]:03}", jurisdiction_id="J-"+j, pathway_id=f"PW-{j}-{pw}",
        source_id=s["source_id"], provision_id=locator, condition_code=code,
        condition_other_text=other, criterion_type=kind, source_wording_original=original,
        source_wording_english=en or (original if s["language"]=="en" else ""), explicit_named=named or ("no" if code=="GENERIC_FITNESS" else "yes"),
        exam_country_scope=scope, operative_effect=effect, authority_level=s["source_tier"], effective_from=s["effective_date"],
        clinical_state_rule=clinical or "not_addressed", safeguard_rule=safeguard or ns, trigger_rule=trigger or "not_addressed",
        excerpt_locator=locator, evidence_tier=s["source_tier"], confidence=confidence, classification_note=note,
        reviewer="B3 batch extraction", record_created_at=adjudicated.isoformat(timespec="seconds"),
        modified_at=adjudicated.isoformat(timespec="seconds"),
        provenance_class="recovered_historical" if fetches[key].get("original") else "new_collected")

def saf(j, pw, keys, *, discretion=ns, individual=ns, retest=ns, waiver=ns, review=ns, notes=""):
    row("SAFEGUARDS", jurisdiction_id="J-"+j, pathway_id=f"PW-{j}-{pw}", source_ids=sr(keys),
        discretion=discretion, individual_assessment=individual, retest=retest, waiver=waiver, appeal_review=review, notes=notes)

def trig(j,pw,keys,kind,value,notes=""):
    row("NATIONALITY_ORIGIN_TRIGGERS", jurisdiction_id="J-"+j, pathway_id=f"PW-{j}-{pw}",
        trigger_type=kind, trigger_value=value, source_ids=sr(keys), notes=notes)

def post(j,pw,keys,risk,notes="",mechanism="",*,active=na,on=na,mdt=na,cured=na,residual=na):
    row("CLINICAL_STATUS_RULES", jurisdiction_id="J-"+j, pathway_id=f"PW-{j}-{pw}", active_untreated=active,
        on_treatment=on, mdt_completed=mdt,cured=cured,residual_sequelae=residual,source_ids=sr(keys),notes=notes)
    row("POST_TREATMENT_RISK",jurisdiction_id="J-"+j,pathway_id=f"PW-{j}-{pw}",risk_status=risk,
        mechanism=mechanism,source_ids=sr(keys),notes=notes)

def settle(j,dss,hd,notes="",reason="",additional=""):
    statuses[j] = (dss, hd, notes, reason, additional)

for f in sorted(output.glob(".group*.py"))[:int(sys.argv[1]) if len(sys.argv)>1 else None]:
    exec(compile(f.read_text(), str(f), "exec"))

for i, (file, q) in enumerate(queries, 1):
    result = q.get("result","")
    hits = re.findall(r"^## \d+\. (.+?)\nURL: (.+)$", result, re.M)
    at = q["at"]
    matched = [k for k, g in fetches.items() if g.get("discovered_in") == file.name]
    opened = [k for k in matched if k in checked]
    full = [k for k in opened if not checked[k]["doc_type"].startswith(("partial", "metadata"))]
    candidate_url = fetches[opened[0]]["url"] if opened else (hits[0][1] if hits else "")
    candidate_title = checked[opened[0]]["title"] if opened else (hits[0][0] if hits else "")
    status = "opened" if full else "partial_open" if opened else "search_error" if q.get("error") else "hits_not_adjudicated_from_snippets" if hits else "no_identifiable_hit"
    note = f"Full returned search output: {file.name}. "
    if matched:
        note += "Retrieval records: " + ", ".join(f"fetch_{k}.json" for k in matched) + ". "
    if opened:
        note += "Documents inspected: " + ", ".join(checked[k]["source_id"] for k in opened) + ". "
    note += q.get("notes","")
    row("SEARCH_LOG", search_id=f"Q-B3-{i:04}", jurisdiction_id="J-"+q["jur"], batch="B3", stage=q["stage"],
        language=q["lang"], timestamp_utc=at, query=q["query"], engine_or_database="Devin web search",
        result_url=candidate_url, hit_title=candidate_title, opened_full_text="yes" if full else "no",
        access_status=status, translation_method="original English" if q["lang"]=="en" else "model-assisted search terms and extraction translation; source wording retained",
        notes=note.strip())

for j in sorted({q["jur"] for _,q in queries}):
    jq = [q for _,q in queries if q["jur"] == j]
    kind = {q["stage"] for q in jq}
    opened = {s["source_id"]:s for s in checked.values() if s["jurisdiction_id"]=="J-"+j}
    refs = {key for p in data["POLICY_PATHWAYS"] if p["jurisdiction_id"]=="J-"+j for key in p["governing_source_ids"].split(";")}
    opened.update((s["source_id"],s) for s in checked.values() if s["source_id"] in refs)
    dss, hd, notes, reason, extra = statuses.get(j,(U,C,"Interim checkpoint; searches executed.","Full-text adjudication pending.",""))
    row("JURISDICTION_STATUS", jurisdiction_id="J-"+j,batch="B3",english_pass_done="yes" if "P1_english" in kind else "no",
        local_pass_done="yes" if "P2_local" in kind else "no",
        local_languages_used=";".join(sorted({q["lang"] for q in jq if q["stage"]=="P2_local"})),n_queries=len(jq),
        n_sources_opened=len(opened),best_source_tier=min((s["source_tier"] for s in opened.values()),default="S7"),
        dss_status=dss,hd_named_status=hd,unresolved_reason=reason,hours_spent="",notes=" ".join(filter(None,(notes,extra))))

if not sys.argv[1:]:
    retained = {
        "ARE_dha_decree_pdf": ("DHA document download response","Portal error or shell returned instead of the Cabinet resolution."),
        "ARE_uae_rendered": ("UAE legislation portal metadata for Cabinet Resolution 7/2008","Legislation metadata returned without operative text."),
        "AZE_dangerous_list_alt": ("Azerbaijan legal database result returned for dangerous diseases lead","Retrieved text concerned pharmaceutical regulation, not the migration disease list."),
        "PSE_initial_exam": ("Palestinian Ministry of Health initial examination PDF request","HTML portal response instead of regulation PDF."),
        "PSE_work_regulation": ("An-Najah law portal result for Decision 45/2004","Landing index instead of operative work-permit text."),
        "PSE_work_regulation_alt": ("An-Najah law portal alternate result for Decision 45/2004","Landing index instead of operative work-permit text."),
        "DZA_work_detail_rendered": ("Algerian employment ministry foreign labour page","Rendered retrieval did not expose substantive text."),
        "ARE_denial_lead2": ("Leprosy among migrant workers: ensuring proper treatment","Secondary publication retained for triage only."),
    }
    for key, f in fetches.items():
        if key in checked or not f.get("file"):
            continue
        title, detail = retained.get(key,(key.replace("_"," "), "Snapshot retained with its retrieval log. No adjudication depends on it."))
        source(key,title,f.get("authority") or urlsplit(f["url"]).hostname,
               "retained retrieval snapshot",status="not_used_for_adjudication",
               notes=detail + " See fetch_" + key + ".json for the acquisition record.",
               tier="S7" if key in {"ARE_dha_decree_pdf","PSE_initial_exam","PSE_work_regulation","PSE_work_regulation_alt","DZA_work_detail_rendered"} else None)
    actual = {str(p.relative_to(root)) for p in (root / "02_primary_sources/B3").iterdir() if p.is_file()}
    assert actual == {f["file"] for f in fetches.values() if f.get("file")}, "Source/log inventory mismatch"

for name, cols in schema.TABLES.items():
    with (output / (name + ".csv")).open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=cols, lineterminator="\n")
        writer.writeheader()
        writer.writerows(data[name])
