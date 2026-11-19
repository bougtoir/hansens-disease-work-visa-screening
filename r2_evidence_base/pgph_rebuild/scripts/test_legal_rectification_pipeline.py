"""Regression test: final outputs must derive from the frozen legal audit input PLUS the versioned rectifications.
Fails if any analysis script reads the pre-rectification input, or if any output matches the unrectified codes
where a rectification changes them. Expected values are recomputed here independently; none are hard-coded."""
import json, re, sys
from pathlib import Path
import pandas as pd
RB = Path(__file__).resolve().parents[1]; ROOT = RB.parent; RES = RB / "05_analysis/results"
F = ["unfit", "visa_permit_refusal_inadmissibility", "entry_denial", "removal_deportation", "work_authorisation_denial_revocation", "appeal", "waiver", "retest"]
CONS = {"unfit": "unfit_certificate_possible", "visa_permit_refusal_inadmissibility": "visa_refusal_inadmissibility", "entry_denial": "entry_denial",
        "removal_deportation": "removal_deportation", "work_authorisation_denial_revocation": "work_authorisation_denial_or_revocation"}
fails = []
def check(ok, msg): (None if ok else fails.append(msg)); print(("PASS " if ok else "FAIL ") + msg)

# 1. Static: only scripts that apply (gate4d) or explicitly compare (gate10, gate11) rectifications may read the raw input.
allowed = {"gate4d_legal_audit.py", "gate10_legal_resolution.py", "gate11_final_numeric_audit.py", "test_legal_rectification_pipeline.py"}  # apply or compare rectifications explicitly
for p in sorted((RB / "scripts").glob("*.py")) + sorted((RB / "08_submission").glob("*.py")) + sorted((ROOT / "r2/manuscript").glob("*.py")):
    if p.name in allowed: continue
    txt = p.read_text(errors="ignore")
    reads = [l for l in txt.splitlines() if "LEGAL_CONSEQUENCE_AUDIT_INPUT" in l and re.search(r"read_csv|open\(|DictReader|Path\(", l)]
    check(not reads, f"{p.name} does not read the pre-rectification legal input")

# 2. Independent re-application of rectifications.
raw = pd.read_csv(RB / "05_analysis/LEGAL_CONSEQUENCE_AUDIT_INPUT.csv", dtype=str).fillna("").set_index("jurisdiction_id")
rx = pd.read_csv(RB / "05_analysis/LEGAL_RECTIFICATIONS_20261119.csv", dtype=str).fillna("")
exp = raw.copy()
for x in rx.itertuples():
    check(raw.at[x.jurisdiction_id, x.field] == x.old_value, f"{x.rectification_id}: old value matches frozen input")
    exp.at[x.jurisdiction_id, x.field] = x.new_value
rect = pd.read_csv(RES / "LEGAL_CONSEQUENCE_AUDIT_RECTIFIED.csv", dtype=str).fillna("").set_index("jurisdiction_id")
check(set(rect.index) == set(raw.index), "rectified matrix covers the same jurisdictions as the frozen input")
check(all((rect.loc[exp.index, c] == exp[c]).all() for c in F + ["evidence_status", "current_validity"]), "rectified matrix equals independent re-application")

dl = lambda df, k: set(df.index[df[k].isin(["D", "L"])])
flips = [k for k in F if dl(raw, k) != dl(exp, k)]
check(len(flips) > 0, f"rectifications change documented membership for: {flips}")

# 3. Counts, jurisdiction-level data, numbers JSON.
cnt = pd.read_csv(RES / "LEGAL_CONSEQUENCE_COUNTS.csv").set_index("consequence")
N = json.load(open(RES / "rebuild_numbers.json"))
ex = pd.read_csv(RES / "EXCLUSION_CONSEQUENCES.csv").set_index("jurisdiction_id")
for k in F:
    check(int(cnt.at[k, "documented"]) == len(dl(exp, k)), f"LEGAL_CONSEQUENCE_COUNTS {k} = rectified ({len(dl(exp, k))})")
    check(N["legal_counts"][k]["documented"] == len(dl(exp, k)), f"rebuild_numbers {k} = rectified")
    if k in flips: check(int(cnt.at[k, "documented"]) != len(dl(raw, k)) or dl(raw, k) != dl(exp, k), f"{k}: output not equal to unrectified membership")
    if k in CONS: check(set(ex.index[ex[CONS[k]].astype(str) == "True"]) == dl(exp, k), f"EXCLUSION_CONSEQUENCES {CONS[k]} membership = rectified")
name = pd.read_csv(RB / "04_validated_master/JURISDICTION_MASTER.csv", dtype=str).set_index("jurisdiction_id")["name"]
for k in CONS:
    got = set(N["legal_lists"][k]["D"] + N["legal_lists"][k]["L"]); check(got == {name[j] for j in dl(exp, k)}, f"legal_lists {k} names = rectified")
for k, c in [("appeal", "appeal_review"), ("waiver", "waiver"), ("retest", "retest")]:
    check(set(N["safeguards_A"][c]) == {name[j] for j in dl(exp, k)}, f"safeguards_A {c} = rectified {k}")

# 4. Table 1 / Figure 4 rows (same generator), manuscript, response, cover letter (when built).
sys.path.insert(0, str(RB / "08_submission")); sys.path.insert(0, str(ROOT / "r2/manuscript"))
try:
    import spec_rebuild as SPR
    rows = SPR.table1_rows(); body = {r[0]: r for r in rows[1:]}
    LAB = {"unfit": "unfit certificate", "visa_permit_refusal_inadmissibility": "visa refusal", "entry_denial": "entry denial", "removal_deportation": "removal", "work_authorisation_denial_revocation": "work authorisation"}
    for j in sorted(raw.index):
        if j in ("J-ARE", "J-BHR", "J-KWT", "J-OMN", "J-QAT"): continue
        r = body["GCC states (n = 6)" if j == "J-SAU" else name[j]]
        for k, lab in LAB.items():
            if j == "J-SAU":
                g = [x for x in ("J-ARE", "J-BHR", "J-KWT", "J-OMN", "J-QAT", "J-SAU") if x in dl(exp, k)]
                cell = [p for p in r[3].split("; ") if p.startswith(lab)]
                check(bool(cell) == bool(g) and (not g or len(g) == 6 or all(name[x] in cell[0] for x in g)), f"Table 1 GCC {k} = rectified")
            else:
                check(any(p.startswith(lab) for p in r[3].split("; ")) == (j in dl(exp, k)), f"Table 1 {name[j]} {k} = rectified")
except Exception as e:
    check(False, f"Table 1 generator check raised {e!r}")
out = RB / "08_submission/out"
if (out / "PGPH_R2_cleaned.docx").exists():
    from docx import Document
    T = lambda f: "\n".join([p.text for p in Document(out / f).paragraphs] + [c.text for t in Document(out / f).tables for r in t.rows for c in r.cells])
    C = T("PGPH_R2_cleaned.docx")
    m = {k: re.search(re.escape(l) + r" (\d+) \(", C) for k, l in [("unfit", "an unfit or not-cleared result in"), ("visa_permit_refusal_inadmissibility", "visa or residence-permit refusal or inadmissibility in"),
         ("entry_denial", "entry denial in"), ("removal_deportation", "removal, deportation or repatriation in"), ("work_authorisation_denial_revocation", "denial or revocation of work authorisation in")]}
    for k, v in m.items(): check(v is not None and int(v[1]) == len(dl(exp, k)), f"manuscript Results {k} = rectified")
    ab = re.search(r"stated visa refusal or inadmissibility in (\d+), entry denial in (\d+), removal in (\d+) and denial or revocation of work authorisation in (\d+)", C)
    check(ab is not None and [int(v) for v in ab.groups()] == [len(dl(exp, k)) for k in ("visa_permit_refusal_inadmissibility", "entry_denial", "removal_deportation", "work_authorisation_denial_revocation")], "abstract consequence counts = rectified")
    nrx, nj = len(rx), rx.jurisdiction_id.nunique()
    for f in ("PGPH_R2_Response_to_Reviewers.docx", "PGPH_R2_Cover_Letter.docx"):
        check(f"{nrx} versioned corrections in {nj} jurisdictions" in T(f), f"{f} states {nrx} corrections in {nj} jurisdictions")
    check("Zimbabwe’s consequences were not counted" in C and "October 2016 version of the Ministry of Health" in C, "manuscript carries ZWE and BRN qualifications")
print(f"\n{'ALL PASS' if not fails else str(len(fails)) + ' FAILURES'}"); sys.exit(1 if fails else 0)
