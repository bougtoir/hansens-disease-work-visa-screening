#!/usr/bin/env python3
"""Create Sankey/alluvial flow diagram showing data accessibility for the scoping review.
Uses matplotlib since plotly/kaleido has dependency issues in this environment."""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
from matplotlib.path import Path
import os

OUTPUT_DIR = "/home/ubuntu/scoping_review/figures"
os.makedirs(OUTPUT_DIR, exist_ok=True)

plt.rcParams['font.family'] = 'DejaVu Sans'


def _draw_flow_band(ax, x0, y0_start, y0_end, x1, y1_start, y1_end, color, alpha=0.35):
    """Draw a curved flow band between two vertical bars."""
    xm = (x0 + x1) / 2

    verts_top = [
        (x0, y0_end),
        (xm, y0_end),
        (xm, y1_end),
        (x1, y1_end),
    ]
    verts_bot = [
        (x1, y1_start),
        (xm, y1_start),
        (xm, y0_start),
        (x0, y0_start),
    ]

    codes_top = [Path.MOVETO, Path.CURVE4, Path.CURVE4, Path.CURVE4]
    codes_bot = [Path.LINETO, Path.CURVE4, Path.CURVE4, Path.CURVE4]

    all_verts = verts_top + verts_bot + [(x0, y0_end)]
    all_codes = codes_top + codes_bot + [Path.CLOSEPOLY]

    path = Path(all_verts, all_codes)
    patch = mpatches.PathPatch(path, facecolor=color, edgecolor='none', alpha=alpha, zorder=2)
    ax.add_patch(patch)


def _draw_node(ax, x, y, w, h, label, color, fontsize=8, text_color='white', bold=True):
    """Draw a node rectangle with centered label."""
    rect = FancyBboxPatch((x - w/2, y), w, h,
                           boxstyle="round,pad=0.02",
                           facecolor=color, edgecolor='black', linewidth=0.8, zorder=5)
    ax.add_patch(rect)
    weight = 'bold' if bold else 'normal'
    ax.text(x, y + h/2, label, ha='center', va='center',
            fontsize=fontsize, color=text_color, fontweight=weight, zorder=6,
            linespacing=1.2)


def _get_jp_font():
    """Find a Japanese-capable font."""
    import matplotlib.font_manager as fm
    for pattern in ['Noto Sans CJK JP', 'Noto Serif CJK JP', 'Noto Sans CJK',
                    'IPAGothic', 'IPAPGothic', 'TakaoPGothic']:
        matches = [f.name for f in fm.fontManager.ttflist if f.name == pattern]
        if matches:
            return matches[0]
    # Broader search
    for f in fm.fontManager.ttflist:
        if 'CJK' in f.name and 'JP' in f.name:
            return f.name
    return 'DejaVu Sans'


def _create_sankey_diagram(lang='en'):
    """Create data accessibility Sankey/flow diagram."""
    is_ja = (lang == 'ja')
    jp_font = _get_jp_font() if is_ja else 'DejaVu Sans'

    # Set font globally for this figure
    if is_ja:
        plt.rcParams['font.family'] = jp_font
    else:
        plt.rcParams['font.family'] = 'DejaVu Sans'

    fig, ax = plt.subplots(figsize=(14, 8))
    ax.set_xlim(-1.0, 11.5)
    ax.set_ylim(-0.5, 10.5)
    ax.axis('off')

    # ---- Text definitions ----
    if is_ja:
        title_main = 'データ到達フロー：2段階検索戦略'
        title_sub = '労働ビザ医療要件について197カ国を調査'
        lbl_total = '197カ国\n調査対象'
        lbl_eng_ok = '英語で特定\n110カ国\n(55.8%)'
        lbl_eng_no = '英語情報\n不十分\n87カ国\n(44.2%)'
        lbl_disease = '疾病特異的\nスクリーニング\n確認\n58カ国'
        lbl_noreq = '疾病特異的\n要件なし\n52カ国'
        lbl_multi = '他言語で確認\n5カ国 (2.5%)'
        lbl_unres = '最終的に\n到達不可\n82カ国\n(41.6%)'
        lbl_hansen = 'ハンセン病\nスクリーニング\n20カ国'
        lbl_nohansen = 'ハンセン病\nなし\n43カ国'
        stage_total = '合計'
        stage1 = '第1段階：\n英語検索'
        stage2 = '第2段階：\n多言語調査'
        stage3 = 'ハンセン病\n結果'
        summary_lines = [
            'データ到達：115カ国 (58.4%)',
            '  英語情報源：110 (55.8%)',
            '  他言語：5 (2.5%)',
            '到達不可：82カ国 (41.6%)',
        ]
        fontprop = {'fontfamily': jp_font}
    else:
        title_main = 'Data Accessibility Flow: Two-Stage Search Strategy'
        title_sub = '197 countries examined for work visa medical requirements'
        lbl_total = '197\nCountries\nExamined'
        lbl_eng_ok = 'Characterized\nvia English\n110 countries\n(55.8%)'
        lbl_eng_no = 'English\nInsufficient\n87 countries\n(44.2%)'
        lbl_disease = 'Disease-Specific\nScreening\nConfirmed\n58 countries'
        lbl_noreq = 'No Disease-\nSpecific\nRequirement\n52 countries'
        lbl_multi = 'Confirmed via\nOther Languages\n5 countries (2.5%)'
        lbl_unres = 'Ultimately\nUnresolvable\n82 countries\n(41.6%)'
        lbl_hansen = "Hansen's Disease\nScreening\n20 countries"
        lbl_nohansen = "No Hansen's\nDisease\n43 countries"
        stage_total = 'Total'
        stage1 = 'Stage 1:\nEnglish Search'
        stage2 = 'Stage 2:\nMultilingual\nResearch'
        stage3 = "Hansen's Disease\nOutcome"
        summary_lines = [
            'Data Reached: 115 countries (58.4%)',
            '  English sources: 110 (55.8%)',
            '  Other languages: 5 (2.5%)',
            'Unreachable: 82 countries (41.6%)',
        ]
        fontprop = {}

    # ---- Title ----
    ax.text(5.25, 10.2, title_main,
            fontsize=15, fontweight='bold', ha='center', color='#1A237E', **fontprop)
    ax.text(5.25, 9.75, title_sub,
            fontsize=11, ha='center', color='#455A64', style='italic', **fontprop)

    # ---- Layout parameters ----
    node_w = 1.6
    total_h = 8.0
    gap = 0.3
    gap2 = 0.2

    # Column 1: Total (x=0)
    _draw_node(ax, 0.0, 0.5, node_w, total_h, lbl_total,
               '#1565C0', fontsize=12, text_color='white')

    # Column 2: Stage 1 results (x=3)
    h_eng = total_h * 110 / 197
    h_insuf = total_h * 87 / 197

    y_eng = 0.5 + h_insuf + gap
    y_insuf = 0.5

    _draw_node(ax, 3.0, y_eng, node_w, h_eng, lbl_eng_ok,
               '#43A047', fontsize=9)
    _draw_node(ax, 3.0, y_insuf, node_w, h_insuf, lbl_eng_no,
               '#E65100', fontsize=9)

    # Flow: Total -> Stage 1
    _draw_flow_band(ax, 0.0 + node_w/2, 0.5 + h_insuf + gap, 0.5 + total_h,
                    3.0 - node_w/2, y_eng, y_eng + h_eng, '#43A047', 0.3)
    _draw_flow_band(ax, 0.0 + node_w/2, 0.5, 0.5 + h_insuf,
                    3.0 - node_w/2, y_insuf, y_insuf + h_insuf, '#E65100', 0.3)

    # Column 3: Breakdown (x=6)
    h_dis = total_h * 58 / 197
    h_noreq = total_h * 52 / 197
    h_multi = total_h * 5 / 197
    h_unres = total_h * 82 / 197
    h_multi_actual = max(h_multi, 0.5)

    y_unres = 0.5
    y_multi = y_unres + h_unres + gap2
    y_noreq = y_multi + h_multi_actual + gap2
    y_dis = y_noreq + h_noreq + gap2

    _draw_node(ax, 6.0, y_dis, node_w, h_dis, lbl_disease,
               '#2E7D32', fontsize=8)
    _draw_node(ax, 6.0, y_noreq, node_w, h_noreq, lbl_noreq,
               '#66BB6A', fontsize=8, text_color='black')
    _draw_node(ax, 6.0, y_multi, node_w, h_multi_actual, lbl_multi,
               '#FF9800', fontsize=7.5, text_color='black')
    _draw_node(ax, 6.0, y_unres, node_w, h_unres, lbl_unres,
               '#C62828', fontsize=8)

    # Flow: English OK -> disease-specific + no requirement
    _draw_flow_band(ax, 3.0 + node_w/2, y_eng + h_noreq, y_eng + h_eng,
                    6.0 - node_w/2, y_dis, y_dis + h_dis, '#2E7D32', 0.25)
    _draw_flow_band(ax, 3.0 + node_w/2, y_eng, y_eng + h_noreq,
                    6.0 - node_w/2, y_noreq, y_noreq + h_noreq, '#66BB6A', 0.25)

    # Flow: English insufficient -> multilingual + unresolved
    _draw_flow_band(ax, 3.0 + node_w/2, y_insuf + h_unres, y_insuf + h_insuf,
                    6.0 - node_w/2, y_multi, y_multi + h_multi_actual, '#FF9800', 0.25)
    _draw_flow_band(ax, 3.0 + node_w/2, y_insuf, y_insuf + h_unres,
                    6.0 - node_w/2, y_unres, y_unres + h_unres, '#C62828', 0.25)

    # Column 4: Hansen's disease outcome (x=9)
    h_hansen = max(total_h * 20 / 197, 0.7)
    h_nohansen = total_h * 43 / 197

    y_nohansen = y_dis
    y_hansen = y_dis + h_nohansen + gap2

    _draw_node(ax, 9.0, y_hansen, node_w, h_hansen, lbl_hansen,
               '#D32F2F', fontsize=8)
    _draw_node(ax, 9.0, y_nohansen, node_w, h_nohansen, lbl_nohansen,
               '#4CAF50', fontsize=8)

    # Flow: disease-specific -> Hansen's + no Hansen's
    h_hansen_flow = total_h * 20 / 197
    _draw_flow_band(ax, 6.0 + node_w/2, y_dis + h_dis - h_hansen_flow, y_dis + h_dis,
                    9.0 - node_w/2, y_hansen, y_hansen + h_hansen, '#D32F2F', 0.25)
    h_no_hansen_from_dis = total_h * 38 / 197
    _draw_flow_band(ax, 6.0 + node_w/2, y_dis, y_dis + h_no_hansen_from_dis,
                    9.0 - node_w/2, y_nohansen, y_nohansen + h_no_hansen_from_dis, '#4CAF50', 0.2)

    # Flow: multilingual confirmed -> no Hansen's
    _draw_flow_band(ax, 6.0 + node_w/2, y_multi, y_multi + h_multi_actual,
                    9.0 - node_w/2, y_nohansen + h_no_hansen_from_dis,
                    y_nohansen + h_nohansen, '#FF9800', 0.2)

    # ---- Stage column labels ----
    label_y = 9.3
    ax.text(0.0, label_y, stage_total, ha='center', fontsize=10,
            fontweight='bold', color='#1565C0', **fontprop)
    ax.text(3.0, label_y, stage1, ha='center', fontsize=9,
            fontweight='bold', color='#1976D2', **fontprop)
    ax.text(6.0, label_y, stage2, ha='center', fontsize=9,
            fontweight='bold', color='#E65100', **fontprop)
    ax.text(9.0, label_y, stage3, ha='center', fontsize=9,
            fontweight='bold', color='#C62828', **fontprop)

    # ---- Summary box ----
    summary_text = '\n'.join(summary_lines)
    props = dict(boxstyle='round,pad=0.5', facecolor='#F5F5F5', edgecolor='#9E9E9E', alpha=0.9)
    ax.text(10.8, 2.5, summary_text, fontsize=8, verticalalignment='top',
            bbox=props, linespacing=1.5, **fontprop)

    plt.tight_layout()
    suffix = 'ja' if is_ja else 'en'
    out_png = os.path.join(OUTPUT_DIR, f"fig6_sankey_{suffix}.png")
    out_tiff = os.path.join(OUTPUT_DIR, f"fig6_sankey_{suffix}.tiff")
    plt.savefig(out_png, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.savefig(out_tiff, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()
    print(f"Saved: {out_png}")


def create_sankey_en():
    _create_sankey_diagram('en')


def create_sankey_ja():
    _create_sankey_diagram('ja')


if __name__ == '__main__':
    create_sankey_en()
    create_sankey_ja()
    print("Sankey diagrams created successfully!")
