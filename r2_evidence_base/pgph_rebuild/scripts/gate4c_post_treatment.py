"""Collapse pathway-level POST_TREATMENT_RISK to one jurisdiction-level status for HD_NAMED jurisdictions."""
import pandas as pd
from pathlib import Path
RB = Path(__file__).resolve().parents[1]; M = RB / "04_validated_master"; OUT = RB / "05_analysis/results"
PRIORITY = ["explicit_post_treatment_exclusion", "treatment_distinction_exempts_or_conditional",
            "no_treatment_distinction_named_condition", "cannot_determine", "not_applicable_hd_not_named"]
d = pd.read_csv(M / "JURISDICTION_MASTER.csv", dtype=str, keep_default_na=False)
pt = pd.read_csv(M / "POST_TREATMENT_RISK.csv", dtype=str, keep_default_na=False)
adj = pd.read_csv(RB / "05_analysis/POST_TREATMENT_ADJUDICATION.csv", dtype=str)
hd = sorted(d[d.hd_final == "confirmed"].jurisdiction_id)
rows = []
for j in hd:
    s = set(pt[pt.jurisdiction_id == j].risk_status) or {"cannot_determine"}
    st = next(p for p in PRIORITY if p in s); basis = "pathway_priority"
    if j in set(adj.jurisdiction_id):
        st = adj.set_index("jurisdiction_id").loc[j, "risk_status"]; basis = "lead_adjudication"
    rows.append(dict(jurisdiction_id=j, post_treatment_status=st, pathway_statuses=";".join(sorted(s)), basis=basis))
x = pd.DataFrame(rows); x.to_csv(OUT / "POST_TREATMENT_JURISDICTION.csv", index=False)
assert len(x) == len(hd)
print(x.post_treatment_status.value_counts().to_dict()); print(x[x.pathway_statuses.str.contains(";")].to_string(index=False))
