"""Final legal-consequence source audit of the HD-named jurisdictions.

Input: 05_analysis/LEGAL_CONSEQUENCE_AUDIT_INPUT.csv (provision-level adjudication, one row per jurisdiction).
Codes: D = explicit operative rule attaches the consequence to HD (named or by incorporation);
L = linked procedural rule (HD makes the person unfit/failed in one located source and a second located
source attaches the consequence to that failure); U = only a generic provision located, HD not shown to
fall within it; N = not located. Only D and L count as documented consequences.
Outputs: the source-linked audit matrix, the regex autoscreen (kept for provenance) and the audited
EXCLUSION_CONSEQUENCES.csv consumed downstream.
"""
from pathlib import Path
import pandas as pd

RB = Path(__file__).resolve().parents[1]
M, RES, QC = RB / "04_validated_master", RB / "05_analysis/results", RB / "07_audits/FINAL_QC"
QC.mkdir(exist_ok=True)
CONS = {"unfit": "unfit_certificate_possible", "visa_permit_refusal_inadmissibility": "visa_refusal_inadmissibility",
        "entry_denial": "entry_denial", "removal_deportation": "removal_deportation",
        "work_authorisation_denial_revocation": "work_authorisation_denial_or_revocation"}
SAFE = ["appeal", "waiver", "retest"]
CODES = {"D", "L", "U", "N", "D-unavailable"}
STATUS = {"EXPLICIT OPERATIVE RULE", "LINKED PROCEDURAL RULE", "REASONED BUT UNCONFIRMED INTERPRETATION",
          "NOT LOCATED", "HISTORICAL-ONLY", "CONTRADICTORY"}

au = pd.read_csv(RB / "05_analysis/LEGAL_CONSEQUENCE_AUDIT_INPUT.csv", dtype=str).fillna("")
# Versioned rectifications (2026-11-19 legal re-check) applied on top of the frozen audit input; provenance kept per row.
rx = pd.read_csv(RB / "05_analysis/LEGAL_RECTIFICATIONS_20261119.csv", dtype=str).fillna("")
rxe = pd.read_csv(RB / "05_analysis/LEGAL_RECTIFICATION_EVIDENCE.csv", dtype=str).fillna("").set_index("evidence_id")
au["rectifications"] = ""; au["legacy_codes"] = ""
for x in rx.itertuples():
    i = au.index[au.jurisdiction_id == x.jurisdiction_id]
    assert len(i) == 1 and au.at[i[0], x.field] == x.old_value, (x.rectification_id, x.field, au.loc[i, x.field].tolist())
    assert all(e in rxe.index for e in x.evidence_ids.split(";")), x.rectification_id
    au.at[i[0], x.field] = x.new_value
    au.at[i[0], "rectifications"] = ";".join(filter(None, [au.at[i[0], "rectifications"], x.rectification_id]))
    au.at[i[0], "legacy_codes"] = ";".join(filter(None, [au.at[i[0], "legacy_codes"], f"{x.field}={x.old_value}"]))
    au.at[i[0], "key_evidence_ids"] = ";".join(dict.fromkeys(au.at[i[0], "key_evidence_ids"].split(";") + x.evidence_ids.split(";")))
ev = pd.read_csv(M / "EVIDENCE_ROWS_ADJUDICATED.csv", dtype=str).fillna("").set_index("evidence_id")
sr = pd.read_csv(M / "SOURCE_REGISTER.csv", dtype=str).fillna("").drop_duplicates("source_id").set_index("source_id")
tc = pd.read_csv(RB / "07_audits/TARGETED_CHECKS/SOURCES.csv", dtype=str).fillna("").set_index("file")
tcc = pd.read_csv(RB / "07_audits/TARGETED_CHECKS/TARGETED_CHECKS.csv", dtype=str).fillna("").set_index("check_id")
auto = pd.read_csv(RES / "EXCLUSION_CONSEQUENCES.csv", dtype=str)
if "consequence_basis" not in auto.columns:  # first pass on fresh gate4b output
    auto.to_csv(RES / "EXCLUSION_CONSEQUENCES_AUTOSCREEN.csv", index=False)
auto = pd.read_csv(RES / "EXCLUSION_CONSEQUENCES_AUTOSCREEN.csv", dtype=str)

assert set(au.jurisdiction_id) == set(auto.jurisdiction_id) and len(au) == len(auto), "audit must cover every HD-named jurisdiction"
for c in list(CONS) + SAFE:
    assert set(au[c]) <= CODES, (c, set(au[c]) - CODES)
assert all(s.split(" (")[0] in STATUS for s in au.evidence_status), set(au.evidence_status)
for ids in au.key_evidence_ids:
    for e in ids.split(";"):
        assert e.startswith(("TC-", "OBSERVED")) or e in ev.index or e in rxe.index, e

rows = []
for r in au.itertuples():
    for e in r.key_evidence_ids.split(";"):
        base = {k: getattr(r, k) for k in au.columns}
        if e in ev.index:
            x = ev.loc[e]; s = sr.loc[x.source_id] if x.source_id in sr.index else pd.Series(dtype=str)
            base.update(evidence_id=e, source_id=x.source_id, document=s.get("title", ""), issuing_institution=s.get("issuing_authority", ""),
                        document_date=s.get("effective_date", ""), status_in_force=s.get("status_in_force", ""), url=s.get("url", ""),
                        local_path=s.get("local_path", ""), sha256=s.get("sha256", ""), redistribution=s.get("redistribution", ""),
                        quote_original=x.source_wording_original, quote_english=x.source_wording_english or "(original English)",
                        locator=x.excerpt_locator, criterion_type=x.criterion_type, clinical_state_rule=x.clinical_state_rule)
        elif e.startswith("TC-"):
            c = tcc.loc[e]; fs = [f for f in c.source_files.split(";") if f in tc.index]
            base.update(evidence_id=e, source_id="targeted check", document=c.question, locator=c.locator, quote_english=c.finding,
                        status_in_force=c.status, local_path=";".join("pgph_rebuild/07_audits/TARGETED_CHECKS/sources/" + f for f in fs),
                        sha256=";".join(tc.loc[f, "sha256"] for f in fs), url=";".join(tc.loc[f, "url"] for f in fs),
                        redistribution=";".join(tc.loc[f, "redistribution"] for f in fs))
        elif e in rxe.index:
            x = rxe.loc[e]
            base.update(evidence_id=e, source_id="LEGAL_RECTIFICATION_EVIDENCE.csv", document=x.locator, locator=x.locator, quote_original=x.quote,
                        quote_english=x.quote, status_in_force=x.temporal_validity, url=x.url, local_path=x.local_path, sha256=x.sha256)
        else:
            base.update(evidence_id=e, source_id="OBSERVED_DENIALS.csv")
        base.update(ai_second_pass="", final_adjudication="see code columns", human_signoff="PENDING")
        rows.append(base)
pd.DataFrame(rows).to_csv(QC / "R2_LEGAL_CONSEQUENCE_SOURCE_AUDIT.csv", index=False)
au.to_csv(RES / "LEGAL_CONSEQUENCE_AUDIT_RECTIFIED.csv", index=False)

out = auto.set_index("jurisdiction_id"); a = au.set_index("jurisdiction_id")
for k, col in CONS.items():
    out[col] = a[k].isin(["D", "L"]).reindex(out.index)
out["consequence_basis"] = a.basis
out["consequence_text"] = "audited (gate4d); regex screen kept in EXCLUSION_CONSEQUENCES_AUTOSCREEN.csv"
out.reset_index().to_csv(RES / "EXCLUSION_CONSEQUENCES.csv", index=False)

cnt = []
gcc = a.index[a.mechanism_id == "GCC-WAFID"]
def n_mech(v):  # GCC-WAFID counted once if any member documents the consequence
    d = v.isin(["D", "L"]); return int(d[~d.index.isin(gcc)].sum() + d[gcc].any())
for k in list(CONS) + SAFE:
    v = a[k]
    cnt.append(dict(consequence=k, direct=int((v == "D").sum()), explicitly_unavailable=int((v == "D-unavailable").sum()), linked=int((v == "L").sum()),
                    documented=int(v.isin(["D", "L"]).sum()), generic_only_unknown=int((v == "U").sum()),
                    not_located=int((v == "N").sum()), documented_mechanisms=n_mech(v),
                    auto_screen=int((auto.set_index("jurisdiction_id")[CONS[k]] == "True").sum()) if k in CONS else None))
pd.DataFrame(cnt).to_csv(RES / "LEGAL_CONSEQUENCE_COUNTS.csv", index=False)
print(pd.DataFrame(cnt).to_string(index=False))
