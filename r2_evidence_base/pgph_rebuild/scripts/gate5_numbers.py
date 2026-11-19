"""Single source of truth for every number reported in the R2 manuscript, letters, tables and figures."""
import json
import pandas as pd, numpy as np
from scipy import stats
from pathlib import Path
RB = Path(__file__).resolve().parents[1]; ROOT = RB.parent; M = RB / "04_validated_master"; R = RB / "05_analysis/results"
rd = lambda p: pd.read_csv(p, dtype=str, keep_default_na=False)
d = rd(M / "JURISDICTION_MASTER.csv"); N = len(d)
C = json.load(open(R / "counts.json")); fq = pd.read_csv(R / "DISEASE_FREQUENCIES.csv").set_index("condition_code")
ex = pd.read_csv(R / "EXCLUSION_CONSEQUENCES.csv"); pt = rd(R / "POST_TREATMENT_JURISDICTION.csv")
sl = rd(M / "SEARCH_LOG.csv"); sr = rd(M / "SOURCE_REGISTER.csv"); ev = rd(M / "EVIDENCE_ROWS_ADJUDICATED.csv")
va = rd(RB / "07_audits/VALIDATION_PASS2/VALIDATION_ROWS.csv"); od = rd(M / "OBSERVED_DENIALS.csv")
sg = rd(M / "SAFEGUARDS.csv"); og = rd(M / "NATIONALITY_ORIGIN_TRIGGERS.csv"); tc = rd(RB / "07_audits/TARGETED_CHECKS/TARGETED_CHECKS.csv")
name = d.set_index("jurisdiction_id").name
A = sorted(d[d.category == "A"].jurisdiction_id, key=lambda j: name[j]); GCC = ["J-ARE", "J-BHR", "J-KWT", "J-OMN", "J-QAT", "J-SAU"]
pct = lambda k, n: round(100 * k / n, 1)
cat = d.category.value_counts().reindex(["A", "B", "D", "U"], fill_value=0).astype(int).to_dict()
reg = d.groupby("region").agg(n=("category", "size"), unresolved=("category", lambda x: int((x == "U").sum())), A=("category", lambda x: int((x == "A").sum())))
sub = d.groupby("search_batch").agg(n=("category", "size"), unresolved=("category", lambda x: int((x == "U").sum())))
pos = cat["A"] + cat["B"]
freq = {k: dict(named=int(fq.loc[k, "n_jurisdictions_named"]), diagnostic=int(fq.loc[k, "n_named_as_diagnostic_criterion"]),
                pct_of_positive=pct(int(fq.loc[k, "n_jurisdictions_named"]), pos)) for k in fq.index}
exc_cols = [c for c in ex.columns if c not in ("jurisdiction_id", "evidence_ids", "consequence_text", "source_note", "consequence_basis")]
exc = {c: int(ex[c].astype(bool).sum()) for c in exc_cols}
la = pd.read_csv(RB / "05_analysis/LEGAL_CONSEQUENCE_AUDIT_INPUT.csv", dtype=str).fillna("")
legal_lists = {k: {c: [name[j] for j in la[la[k] == c].jurisdiction_id] for c in ("D", "L", "U")} for k in ("unfit", "visa_permit_refusal_inadmissibility", "entry_denial", "removal_deportation", "work_authorisation_denial_revocation")}
legal_counts = pd.read_csv(R / "LEGAL_CONSEQUENCE_COUNTS.csv").set_index("consequence").drop(columns="auto_screen").to_dict("index")
exc_lists = {c: [name[j] for j in ex[ex[c].astype(bool)].jurisdiction_id] for c in exc_cols}
ptc = pt.post_treatment_status.value_counts().to_dict(); ptl = {k: [name[j] for j in pt[pt.post_treatment_status == k].jurisdiction_id] for k in ptc}
yes = lambda c: sg[c].str.startswith("explicit_available")
saf = {c: sorted({name[j] for j in sg[yes(c)].jurisdiction_id}) for c in ("waiver", "appeal_review", "retest")}
saf_A = {c: sorted({name[j] for j in sg[yes(c) & sg.jurisdiction_id.isin(A)].jurisdiction_id}) for c in saf}
ogA = og[og.jurisdiction_id.isin(A) & (og.trigger_type != "none") & (og.trigger_type != "")]
origin = {name[j]: sorted(set(g.trigger_type)) for j, g in ogA.groupby("jurisdiction_id")}
stage = sl.stage.value_counts().to_dict()
lang_en = int((sr.language.str.lower().str.startswith("en")).sum())
# burden comparison (WHO GHO NTD_LEPR2, latest 2019-2025)
g = pd.DataFrame(json.load(open(ROOT / "data/raw/who_gho/NTD_LEPR2.json"))["value"])
g = g[(g.SpatialDimType == "COUNTRY") & g.TimeDim.between(2019, 2025) & g.NumericValue.notna()].sort_values("TimeDim").groupby("SpatialDim").tail(1).set_index("SpatialDim")
d["ncdr"] = d.iso3.map(g.NumericValue)
def comp(df):
    a, o = df[df.category == "A"].ncdr.dropna(), df[df.category.isin(["B", "D"])].ncdr.dropna()
    u, p = stats.mannwhitneyu(a, o, alternative="two-sided"); rb = 1 - 2 * u / (len(a) * len(o))
    rng = np.random.default_rng(20261119); bs = [np.median(rng.choice(a, len(a))) - np.median(rng.choice(o, len(o))) for _ in range(5000)]
    return dict(n_A=len(a), n_other=len(o), median_A=round(float(a.median()), 2), median_other=round(float(o.median()), 2),
                diff=round(float(a.median() - o.median()), 2), ci=[round(float(x), 2) for x in np.percentile(bs, [2.5, 97.5])],
                U=float(u), p=round(float(p), 3), rank_biserial=round(float(-rb), 3), A_missing=[name[j] for j in df[(df.category == "A") & df.ncdr.isna()].jurisdiction_id])
d1 = pd.concat([d[~d.jurisdiction_id.isin(GCC)], d[d.jurisdiction_id == "J-SAU"]])
burden = dict(all=comp(d), gcc_once=comp(d1), coverage=int(d.ncdr.notna().sum()), years=[int(g.TimeDim.min()), int(g.TimeDim.max())])
out = dict(N=N, frame=d.frame_status.value_counts().to_dict(), categories=cat, pct={k: pct(v, N) for k, v in cat.items()},
           documented_positive=pos, documented_positive_pct=pct(pos, N), sufficient=C["sufficient_evidence_N"], hd_named=A and len(A),
           hd_named_names=[name[j] for j in A], hd_mechanisms=len(A) - len(GCC) + 1, hd_cannot_determine=C["hd_cannot_determine"],
           hd_not_established=C["hd_not_established"], absence_names=[name[j] for j in sorted(d[d.category == "D"].jurisdiction_id, key=lambda j: name[j])],
           freq=freq, exclusion=exc, exclusion_lists=exc_lists, legal_counts=legal_counts, legal_lists=legal_lists, post_treatment=ptc, post_treatment_lists=ptl,
           safeguards={k: len(v) for k, v in saf.items()}, safeguards_any=len(set(sg[yes("waiver") | yes("appeal_review") | yes("retest")].jurisdiction_id)),
           safeguards_A=saf_A, origin_triggers_A=origin, denials=dict(cases=len(od), jurisdictions=sorted({name[j] for j in od.jurisdiction_id})),
           region=reg.to_dict("index"), batch=sub.to_dict("index"),
           search=dict(queries=len(sl), stage=stage, opened_full_text=int((sl.opened_full_text == "yes").sum()), sources=len(sr),
                       sources_saved=C["sources_saved_locally"], sources_english=lang_en, sources_other=len(sr) - lang_en, evidence_rows=len(ev),
                       criterion=ev.criterion_type.value_counts().to_dict(), provenance=ev.provenance_class.value_counts().to_dict()),
           validation=dict(items=len(va), agree=int((va.agree == "yes").sum()), disagreements=va[va.agree != "yes"][["jurisdiction_id", "item_type", "resolution"]].to_dict("records")),
           targeted_checks=tc[["check_id", "status"]].to_dict("records"), burden=burden)
json.dump(out, open(R / "rebuild_numbers.json", "w"), indent=1, ensure_ascii=False, default=int)
print(json.dumps({k: out[k] for k in ("categories", "hd_named", "hd_mechanisms", "exclusion", "post_treatment", "safeguards", "safeguards_A", "origin_triggers_A", "burden", "validation")}, ensure_ascii=False, default=int)[:3500])
