"""Gate 4: merge validated batch tables into the master and derive every count from evidence rows (no hard-coded values)."""
import json, sys
from pathlib import Path
import pandas as pd
RB = Path(__file__).resolve().parents[1]; sys.path.insert(0, str(RB / "data"))
import schema as S
EX, M, OUT = RB / "03_extracted_evidence", RB / "04_validated_master", RB / "05_analysis/results"
M.mkdir(exist_ok=True); OUT.mkdir(parents=True, exist_ok=True)
frame = pd.read_csv(RB / "data/JURISDICTION_FRAME.csv")
T = {}
for name, cols in S.TABLES.items():
    parts = [pd.read_csv(p, dtype=str, keep_default_na=False) for p in sorted(EX.glob(f"B*/{name}.csv"))]
    T[name] = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=cols)
    T[name].to_csv(M / f"{name}.csv", index=False)
st = T["JURISDICTION_STATUS"]
assert set(st.jurisdiction_id) == set(frame.jurisdiction_id), "status coverage mismatch"
val_p = RB / "07_audits/VALIDATION_PASS2/VALIDATION_ROWS.csv"
val = pd.read_csv(val_p, dtype=str, keep_default_na=False) if val_p.exists() else pd.DataFrame()
hd = st.set_index("jurisdiction_id").hd_named_status.copy()
if len(val):
    v = val[val.item_type.str.contains("hd", case=False)]
    for _, r in v.iterrows():
        if r.agree.lower() not in ("yes", "true") and r.pass2_value in S.ENUMS["hd_named_status"]:
            hd[r.jurisdiction_id] = r.pass2_value
d = frame.set_index("jurisdiction_id").join(st.set_index("jurisdiction_id")[["dss_status", "best_source_tier", "local_pass_done"]]).assign(hd_final=hd)
d["category"] = "U"
d.loc[d.dss_status == "documented_positive", "category"] = "B"
d.loc[d.hd_final == "confirmed", "category"] = "A"
d.loc[(d.dss_status == "documented_absence") & (d.hd_final != "confirmed"), "category"] = "D"
d.reset_index().to_csv(M / "JURISDICTION_MASTER.csv", index=False)
N = len(d); suff = int((d.dss_status != "unresolved").sum())
counts = [("frame_N", N), ("sufficient_evidence_N", suff), ("unresolved_N", N - suff)]
counts += [(f"category_{c}", int((d.category == c).sum())) for c in "ABDU"]
counts += [(f"hd_{k}", int((d.hd_final == k).sum())) for k in S.ENUMS["hd_named_status"]]
pt = T["POST_TREATMENT_RISK"]; counts += [(f"post_treatment_{k}", int(pt[pt.risk_status == k].jurisdiction_id.nunique())) for k in S.ENUMS["risk_status"]]
counts.append(("observed_denial_cases", len(T["OBSERVED_DENIALS"])))
sg = T["SAFEGUARDS"]
for k in ["waiver", "appeal_review", "retest"]:
    counts.append((f"safeguard_{k}_explicit_available", int(sg[sg[k] == "explicit_available"].jurisdiction_id.nunique())))
counts.append(("sources_registered", len(T["SOURCE_REGISTER"])))
counts.append(("sources_saved_locally", int((T["SOURCE_REGISTER"].local_path != "").sum())))
pd.DataFrame(counts, columns=["quantity", "value"]).to_csv(OUT / "COUNTS.csv", index=False)
ev = T["EVIDENCE_ROWS"].copy(); ev["applicable_to_pathway"] = "yes"
ov = pd.read_csv(RB / "05_analysis/ADJUDICATION_OVERRIDES.csv", dtype=str)
assert set(ov.evidence_id) <= set(ev.evidence_id), "override for unknown evidence_id"
for r in ov.itertuples(): ev.loc[ev.evidence_id == r.evidence_id, r.field] = r.value
ev.to_csv(M / "EVIDENCE_ROWS_ADJUDICATED.csv", index=False)
named = ev[(ev.explicit_named == "yes") & (ev.applicable_to_pathway == "yes")]
B = d[d.dss_status == "documented_positive"].index
cnt = lambda x: x[x.jurisdiction_id.isin(B)].groupby("condition_code").jurisdiction_id.nunique().reindex(S.CONDITION_CODES, fill_value=0)
freq, diag = cnt(named), cnt(named[named.criterion_type == "diagnostic"])
pd.DataFrame({"condition_code": freq.index, "n_jurisdictions_named": freq.values, "n_named_as_diagnostic_criterion": diag.values,
              "denominator_documented_positive": len(B), "denominator_frame": N, "unresolved_jurisdictions": N - suff}).to_csv(OUT / "DISEASE_FREQUENCIES.csv", index=False)
hdj = set(d[d.hd_final == "confirmed"].index)
rec = named[named.condition_code == "HD"].groupby("jurisdiction_id").agg(n_rows=("evidence_id", "size"), evidence_ids=("evidence_id", ";".join),
      criterion_types=("criterion_type", lambda x: ";".join(sorted(set(x)))), source_ids=("source_id", lambda x: ";".join(sorted(set(x)))))
rec = rec.reindex(sorted(hdj | set(rec.index))); rec["hd_named_final"] = [j in hdj for j in rec.index]
rec["in_frequency_numerator"] = [j in B and j in set(named[named.condition_code == "HD"].jurisdiction_id) for j in rec.index]
rec.to_csv(RB / "07_audits/HD_COUNT_RECONCILIATION.csv")
assert int(freq["HD"]) == len(hdj) == int(rec.in_frequency_numerator.sum()), (int(freq["HD"]), len(hdj))
old = d[["iso3", "name", "r1_label_historical", "category", "hd_final", "dss_status"]].reset_index()
old["changed"] = old.r1_label_historical.fillna("") != old.category
old.to_csv(OUT / "OLD_VS_NEW.csv", index=False)
miss = d.groupby(["region", "dss_status"]).size().unstack(fill_value=0); miss.to_csv(OUT / "MISSINGNESS_AND_BOUNDS.csv")
cr = T["CLINICAL_STATUS_RULES"].merge(T["SAFEGUARDS"], on=["jurisdiction_id", "pathway_id"], how="outer", suffixes=("", "_sg"))
cr.to_csv(OUT / "LEGAL_CLINICAL_CROSS_TABS.csv", index=False)
json.dump(dict(counts), open(OUT / "counts.json", "w"), indent=1); print(dict(counts))
