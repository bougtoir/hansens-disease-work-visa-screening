"""Track B Figures 1-6 (PLOS: TIFF 300 dpi LZW, <=7.5 in wide, no in-image titles) from the validated rebuild data."""
import json, sys
from pathlib import Path
import numpy as np, pandas as pd, geopandas as gpd
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Patch
from PIL import Image
RB = Path(__file__).resolve().parents[1]; ROOT = RB.parent; M = RB / "04_validated_master"; R = RB / "05_analysis/results"
OUT = RB / "08_submission/out/figures"; OUT.mkdir(parents=True, exist_ok=True)
sys.path.insert(0, str(RB / "08_submission")); import spec_rebuild as SPR
N = json.load(open(R / "rebuild_numbers.json")); d = pd.read_csv(M / "JURISDICTION_MASTER.csv", dtype=str, keep_default_na=False)
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 8}); W = 7.5
COL = {"A": "#c0392b", "B": "#2e86c1", "D": "#27ae60", "U": "#b3b3b3"}
LAB = {"A": "A: Hansen's disease named", "B": "B: screening documented, Hansen's disease not established", "D": "D: documented absence", "U": "U: unresolved"}
c = N["categories"]
def save(fig, n):
    p = OUT / f"Figure_{n}.png"; fig.savefig(p, dpi=300, facecolor="white"); plt.close(fig)
    Image.open(p).convert("RGB").save(OUT / f"Figure_{n}.tiff", compression="tiff_lzw", dpi=(300, 300))
def box(ax, x, y, w, h, txt, fc, fs=7.5):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.01", fc=fc, ec="#333333", lw=0.6))
    ax.text(x + w / 2, y + h / 2, txt, ha="center", va="center", fontsize=fs)
def arrow(ax, x1, y1, x2, y2): ax.annotate("", (x2, y2), (x1, y1), arrowprops=dict(arrowstyle="->", lw=0.7))
# Figure 1: flow
s, v = N["search"], N["validation"]; st = s["stage"]
fig, ax = plt.subplots(figsize=(W, 4.6)); ax.set_xlim(0, 1); ax.set_ylim(0.1, 1); ax.axis("off")
box(ax, 0.25, 0.86, 0.5, 0.11, f"Analytical frame: {N['N']} jurisdictions\n(193 UN Member States, Holy See, State of Palestine,\nTaiwan, Kosovo, Hong Kong SAR, Macao SAR)", "#ebf5fb")
box(ax, 0.2, 0.63, 0.6, 0.15, f"Searches in 9 regional batches, passes P1–P4\n{s['queries']:,} logged queries (P1 {st.get('P1_english',0)}, P2 {st.get('P2_local',0)}, P3 {st.get('P3_applicability',0)}, P4 {st.get('P4_ambiguity',0)}); {s['opened_full_text']} full-text openings\n{s['sources']} sources registered, {s['sources_saved']} archived with SHA-256\n{s['evidence_rows']} provision-level evidence rows", "#f8f9f9")
box(ax, 0.2, 0.45, 0.6, 0.1, f"Independent verification pass: {v['agree']}/{v['items']} items agreed\nTargeted legal checks: GCC, Thailand, Guyana, Liberia", "#f8f9f9")
arrow(ax, 0.5, 0.86, 0.5, 0.78); arrow(ax, 0.5, 0.63, 0.5, 0.55)
xs = [0.01, 0.26, 0.51, 0.76]
for x, k in zip(xs, "ABDU"):
    arrow(ax, 0.5, 0.45, x + 0.11, 0.33)
    extra = f"\n({N['hd_mechanisms']} mechanisms;\nGCC counted once)" if k == "A" else ""
    box(ax, x, 0.13, 0.22, 0.2, f"Category {k}\nn = {c[k]} ({N['pct'][k]}%)\n" + LAB[k].split(": ", 1)[1].replace(", ", ",\n") + extra, COL[k] + "40", fs=7)
fig.tight_layout(); save(fig, 1)
# Figure 2: batch bars
b = d.groupby(["search_batch", "category"]).size().unstack(fill_value=0).reindex(columns=list("ABDU"), fill_value=0).sort_index()
fig, ax = plt.subplots(figsize=(W, 3.6)); left = np.zeros(len(b))
for k in "ABDU":
    ax.barh(range(len(b)), b[k], left=left, color=COL[k], label=f"{LAB[k]} (n = {c[k]})"); left += b[k].values
ax.set_yticks(range(len(b)), [i.split(" ", 1)[1] + f" (n = {int(b.loc[i].sum())})" for i in b.index]); ax.invert_yaxis()
ax.set_xlabel("Number of jurisdictions"); ax.spines[["top", "right"]].set_visible(False)
ax.legend(fontsize=6, loc="upper center", bbox_to_anchor=(0.35, -0.18), ncol=2, frameon=False)
fig.subplots_adjust(left=0.33, right=0.97, top=0.97, bottom=0.3); save(fig, 2)
# Figure 3: map
w = gpd.read_file(ROOT / "data/raw/naturalearth/ne_50m_admin_0_countries.zip")
w["iso"] = w.ISO_A3.where(w.ISO_A3 != "-99", w.ADM0_A3).replace({"KOS": "XKX", "SDS": "SSD"})
w = w[w.iso != "ATA"].to_crs("ESRI:54030")
g = w.merge(d[["iso3", "category"]], left_on="iso", right_on="iso3", how="left")
miss = sorted(set(d.iso3) - set(g.dropna(subset=["category"]).iso3)); json.dump(miss, open(OUT / "map_unmatched_iso3.json", "w"))
fig, ax = plt.subplots(figsize=(W, 3.9)); ax.axis("off")
g[g.category.isna()].plot(ax=ax, color="#f2f2f2", ec="white", lw=0.2)
for k, cl in COL.items():
    sub = g[g.category == k]
    if len(sub): sub.plot(ax=ax, color=cl, ec="white", lw=0.2)
    sm = sub[sub.geometry.area < 2.5e9]
    if len(sm): ax.scatter(sm.representative_point().x, sm.representative_point().y, s=6, color=cl, ec="black", lw=0.2, zorder=5)
ax.legend(handles=[Patch(color=cl, label=f"{LAB[k]} (n = {c[k]})") for k, cl in COL.items()], fontsize=6, loc="lower left", frameon=True)
fig.tight_layout(); save(fig, 3)
# Figure 4: HD matrix (from Table 1 rows)
rows = SPR.table1_rows(); hdr, body = rows[0], rows[1:]
CONS = ["visa refusal/inadmissibility", "entry denial", "removal", "work authorisation denied/revoked", "unfit certificate"]
cols = ["Unfit\ncertificate", "Visa refusal/\ninadmissibility", "Entry\ndenial", "Removal", "Work auth.\ndenied/revoked", "Treatment\nstatus", "Waiver", "Appeal/\nreview", "Re-exam-\nination", "Origin-\nrelated rule"]
def cell(txt, key):
    for part in txt.split("; "):
        if part.startswith(key): return "Yes" + (" (" + part.split(" (", 1)[1] if " (" in part else "")
    return ""
cells = []
for r in body:
    cons, sg, o = r[3], r[5], r[6]
    org = [x for x in o.split("; ") if x.replace(" ", "_") in SPR.ORIGIN]
    cells.append([cell(cons, "unfit certificate"), cell(cons, "visa refusal"), cell(cons, "entry denial"), cell(cons, "removal"), cell(cons, "work authorisation"),
                  r[4].replace("Treatment history excluded", "History\nexcluded").replace("Conditional/exempting", "Conditional/\nexempting").replace("No distinction", "None").replace("Cannot determine", "Cannot\ndetermine"),
                  cell(sg, "waiver"), cell(sg, "appeal"), cell(sg, "re-examination"), "Yes" if org else ""])
cc = lambda v, j: ("#d5f5e3" if j in (6, 7, 8) else "#fdebd0" if j == 9 else "#f5b7b1") if v.startswith("Yes") else {"None": "#f5b7b1", "History\nexcluded": "#c0392b80", "Conditional/\nexempting": "#d5f5e3", "Cannot\ndetermine": "#eaecee"}.get(v, "#ffffff")
cells = [[x.replace(" (", "\n(") for x in r] for r in cells]
fig, ax = plt.subplots(figsize=(W, 7.2)); ax.axis("off")
tb = ax.table(cellText=cells, rowLabels=[r[0] for r in body], colLabels=cols, cellColours=[[cc(x, j) for j, x in enumerate(r)] for r in cells], cellLoc="center", bbox=[0, 0, 1, 1])
tb.auto_set_font_size(False); tb.set_fontsize(5.6)
for (i, j), cl in tb.get_celld().items():
    cl.set_edgecolor("#999999"); cl.set_linewidth(0.4)
    if i == 0: cl.set_text_props(weight="bold")
fig.subplots_adjust(left=0.17, right=0.99, top=0.99, bottom=0.01); save(fig, 4)
# Figure 5: burden
gj = pd.DataFrame(json.load(open(ROOT / "data/raw/who_gho/NTD_LEPR2.json"))["value"])
yr = N["burden"]["years"]; gj = gj[(gj.SpatialDimType == "COUNTRY") & gj.TimeDim.between(*yr) & gj.NumericValue.notna()].sort_values("TimeDim").groupby("SpatialDim").tail(1).set_index("SpatialDim")
d["ncdr"] = d.iso3.map(gj.NumericValue); floor = 0.01
grp = [("A", d[d.category == "A"].ncdr.dropna()), ("B and D", d[d.category.isin(["B", "D"])].ncdr.dropna())]
fig, ax = plt.subplots(figsize=(4.5, 3.6)); rng = np.random.default_rng(1)
for i, (lab, x) in enumerate(grp):
    ax.scatter(i + rng.uniform(-0.15, 0.15, len(x)), np.maximum(x, floor), s=10, alpha=0.7, color=COL["A"] if lab == "A" else COL["B"], ec="none")
    ax.hlines(max(x.median(), floor), i - 0.25, i + 0.25, color="black", lw=1.2)
ax.set_yscale("log"); ax.set_xticks([0, 1], [f"Category A\n(n = {len(grp[0][1])})", f"Categories B and D\n(n = {len(grp[1][1])})"])
ax.set_ylabel("New case detection rate per 1,000,000"); ax.spines[["top", "right"]].set_visible(False); ax.set_xlim(-0.6, 1.6)
fig.tight_layout(); save(fig, 5)
# Figure 6: disease frequencies
t2 = SPR.table2_rows()[1:]; t2 = [r for r in t2 if r[0] != "Other named conditions"][::-1]
fig, ax = plt.subplots(figsize=(W, 3.8)); y = np.arange(len(t2))
ax.barh(y, [int(r[1]) for r in t2], color="#aed6f1", label="Named in any applicable criterion")
ax.barh(y, [int(r[3]) for r in t2], color="#2471a3", label="Named as a diagnostic criterion")
for yi, r in zip(y, t2): ax.text(int(r[1]) + 0.5, yi, r[1], va="center", fontsize=7)
ax.set_yticks(y, [r[0] for r in t2]); ax.set_xlabel(f"Jurisdictions (of {N['documented_positive']} with documented disease-specific screening)")
ax.spines[["top", "right"]].set_visible(False); ax.legend(fontsize=7, frameon=False, loc="lower right")
fig.tight_layout(); save(fig, 6)
print("figures OK; unmatched map iso3:", miss)
