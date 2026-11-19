"""Derive the per-jurisdiction Hansen's-disease exclusion-consequence variables (normalized; raw tables untouched)."""
import re
import pandas as pd
from pathlib import Path
RB = Path(__file__).resolve().parents[1]; M = RB / "04_validated_master"; OUT = RB / "05_analysis/results"
d = pd.read_csv(M / "JURISDICTION_MASTER.csv", dtype=str, keep_default_na=False)
ev = pd.read_csv(M / "EVIDENCE_ROWS_ADJUDICATED.csv", dtype=str, keep_default_na=False)
pw = pd.read_csv(M / "POLICY_PATHWAYS.csv", dtype=str, keep_default_na=False)
pt = pd.read_csv(M / "POST_TREATMENT_RISK.csv", dtype=str, keep_default_na=False)
od = pd.read_csv(M / "OBSERVED_DENIALS.csv", dtype=str, keep_default_na=False)
hd = ev[(ev.condition_code == "HD") & (ev.explicit_named == "yes") & (ev.applicable_to_pathway == "yes")]
P = {"unfit_certificate_possible": r"unfit|not fit|unsuitab", "visa_refusal_inadmissibility": r"inadmiss|refus|prohibited immigrant|exclu|denial of (visa|permit)|not (be )?(issued|granted)|visa",
     "entry_denial": r"entry|admission|landing", "removal_deportation": r"deport|remov|expul|repatriat|cancel|revocation|revoke"}
rows = []
for j in sorted(d[d.hd_final == "confirmed"].jurisdiction_id):
    e = hd[hd.jurisdiction_id == j]; txt = " ".join(e.operative_effect) + " " + " ".join(pw[pw.jurisdiction_id == j].consequence_values)
    r = dict(jurisdiction_id=j, disease_named=True, medical_exam_required=bool((pw[pw.jurisdiction_id == j].medical_exam_required.str.lower().str.startswith("yes")).any()))
    for k, p in P.items(): r[k] = bool(re.search(p, txt, re.I))
    r["treatment_history_exclusion"] = bool((pt[pt.jurisdiction_id == j].risk_status == "explicit_post_treatment_exclusion").any())
    r["actual_documented_refusal"] = bool((od.jurisdiction_id == j).any())
    r["evidence_ids"] = ";".join(e.evidence_id); r["consequence_text"] = txt.strip()[:500]
    rows.append(r)
x = pd.DataFrame(rows); x["work_authorisation_denial_or_revocation"] = x.consequence_text.str.contains("work_permit_(denial|revocation)", regex=True)
adj = pd.read_csv(RB / "05_analysis/CONSEQUENCE_ADJUDICATION.csv", dtype=str)
for r in adj.itertuples():
    assert r.jurisdiction_id in set(x.jurisdiction_id) and r.field in x.columns, r
    x.loc[x.jurisdiction_id == r.jurisdiction_id, r.field] = r.value == "True"
x["source_note"] = "consequence as stated in located source; absence of a value is not evidence that no consequence exists"
x.to_csv(OUT / "EXCLUSION_CONSEQUENCES.csv", index=False)
print(x.drop(columns=["evidence_ids", "consequence_text"]).sum(numeric_only=True).to_dict(), len(x))
