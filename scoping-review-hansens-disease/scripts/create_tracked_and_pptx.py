#!/usr/bin/env python3
"""Generate tracked-changes manuscript and editable PPTX figures.

Produces:
  - PGPH_Full_Paper_EN_tracked.docx  (revision marks showing all changes)
  - PGPH_Figures_Editable.pptx       (one figure per slide, editable)
"""

import os
import re
import copy
import datetime
from docx import Document
from docx.shared import Pt, Inches, Cm, Emu
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt, Emu as PptxEmu
from pptx.enum.text import PP_ALIGN

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.join(SCRIPT_DIR, "..")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
FIGURES_DIR = os.path.join(BASE_DIR, "figures")
TIFF_DIR = os.path.join(OUTPUT_DIR, "figures_tiff")
SUBMITTED = os.path.join(OUTPUT_DIR, "PGPH_Full_Paper_EN_submitted.docx")
REVISED = os.path.join(OUTPUT_DIR, "PGPH_Full_Paper_EN_revised.docx")

# Import text blocks from revision script
import sys
sys.path.insert(0, SCRIPT_DIR)
from create_pgph_revision import (
    TERMINOLOGY_PARA, SCREENING_TYPES_PARA, PRACTICAL_DIFF_PARA,
    BORDER_HEALTH_THRESHOLD_PARA, MULTILINGUAL_STRATEGY_PARA,
    REFUGEE_LIMITATION_PARA, NEW_REFS, S5_HEADING, S5_BODY,
    NEW_FUNDING, NEW_COMPETING, NEW_DATA_AVAIL, NATURAL_EARTH_NOTE,
    _build_injection_plan, _should_replace, CATEGORY_A_SOURCES,
)

AUTHOR = "Onishi T"
DATE = "2026-07-06T00:00:00Z"
REVISION_ID_BASE = 100


# ═══════════════════════════════════════════════════════════════════════════════
# PART 1: TRACKED CHANGES MANUSCRIPT
# ═══════════════════════════════════════════════════════════════════════════════

def _make_ins_element(rid):
    """Create a w:ins element for tracking insertions."""
    ins = OxmlElement('w:ins')
    ins.set(qn('w:id'), str(rid))
    ins.set(qn('w:author'), AUTHOR)
    ins.set(qn('w:date'), DATE)
    return ins


def _make_del_element(rid):
    """Create a w:del element for tracking deletions."""
    del_el = OxmlElement('w:del')
    del_el.set(qn('w:id'), str(rid))
    del_el.set(qn('w:author'), AUTHOR)
    del_el.set(qn('w:date'), DATE)
    return del_el


def _make_run_with_text(text, rPr=None, deleted=False):
    """Create a w:r element with text. If deleted, use w:delText."""
    run_el = OxmlElement('w:r')
    if rPr is not None:
        run_el.append(copy.deepcopy(rPr))
    else:
        rPr_el = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        rPr_el.append(rFonts)
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), '24')
        rPr_el.append(sz)
        szCs = OxmlElement('w:szCs')
        szCs.set(qn('w:val'), '24')
        rPr_el.append(szCs)
        run_el.append(rPr_el)

    if deleted:
        t = OxmlElement('w:delText')
    else:
        t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    run_el.append(t)
    return run_el


def _make_inserted_para(text, rid, rPr_template=None):
    """Create a full paragraph wrapped in w:ins revision marks."""
    p = OxmlElement('w:p')
    # Paragraph properties
    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle')
    pStyle.set(qn('w:val'), 'Normal')
    pPr.append(pStyle)
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), '480')
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:after'), '120')
    pPr.append(spacing)
    # Mark paragraph itself as inserted
    rPr_change = OxmlElement('w:rPr')
    ins_ppr = OxmlElement('w:ins')
    ins_ppr.set(qn('w:id'), str(rid))
    ins_ppr.set(qn('w:author'), AUTHOR)
    ins_ppr.set(qn('w:date'), DATE)
    rPr_change.append(ins_ppr)
    pPr.append(rPr_change)
    p.append(pPr)

    # Wrap run in w:ins
    ins = _make_ins_element(rid + 1)
    run_el = _make_run_with_text(text, rPr=rPr_template)
    ins.append(run_el)
    p.append(ins)
    return p


def _mark_para_as_replaced(orig_child, new_text, rid):
    """Create a paragraph showing deletion of old text and insertion of new."""
    p = copy.deepcopy(orig_child)
    # Get original text and rPr
    orig_runs = list(p.findall(qn('w:r')))
    rPr_template = None
    old_text = ""
    for r in orig_runs:
        rPr_cand = r.find(qn('w:rPr'))
        if rPr_cand is not None and rPr_template is None:
            rPr_template = copy.deepcopy(rPr_cand)
        t_el = r.find(qn('w:t'))
        if t_el is not None and t_el.text:
            old_text += t_el.text

    # Remove existing runs
    for r in orig_runs:
        p.remove(r)

    # Add deletion of old text
    del_el = _make_del_element(rid)
    del_run = _make_run_with_text(old_text, rPr=rPr_template, deleted=True)
    del_el.append(del_run)
    p.append(del_el)

    # Add insertion of new text
    ins_el = _make_ins_element(rid + 1)
    ins_run = _make_run_with_text(new_text, rPr=rPr_template)
    ins_el.append(ins_run)
    p.append(ins_el)

    return p


def create_tracked_changes_manuscript():
    """Generate manuscript with Word revision marks (tracked changes)."""
    print("Creating tracked-changes manuscript...")
    src = Document(SUBMITTED)

    # Build injection plan
    injection_plan = _build_injection_plan(src)

    # Build paragraph-index to body-child-index mapping
    body = src.element.body
    children = list(body)
    para_idx = 0
    para_to_child = {}
    for ci, child in enumerate(children):
        if child.tag == qn('w:p'):
            para_to_child[para_idx] = ci
            para_idx += 1

    # Convert injection plan
    child_injection_plan = {}
    for pidx, texts in injection_plan.items():
        cidx = para_to_child.get(pidx)
        if cidx is not None:
            child_injection_plan[cidx] = texts

    # Build replacement plan
    replacement_plan = {}
    for i, para in enumerate(src.paragraphs):
        text = para.text.strip()
        repl = _should_replace(text)
        if repl:
            replacement_plan[i] = repl

    child_replacement_plan = {}
    for pidx, repl in replacement_plan.items():
        cidx = para_to_child.get(pidx)
        if cidx is not None:
            child_replacement_plan[cidx] = repl

    # Create destination document
    dst = Document()
    for p in list(dst.element.body):
        if p.tag == qn('w:p'):
            dst.element.body.remove(p)

    rid = REVISION_ID_BASE

    # Copy all children, adding revision marks
    for ci, child in enumerate(children):
        if child.tag == qn('w:sectPr'):
            dst.element.body.append(copy.deepcopy(child))
            continue

        if ci in child_replacement_plan:
            # Show as tracked replacement (del old + ins new)
            replaced_p = _mark_para_as_replaced(
                child, child_replacement_plan[ci], rid)
            dst.element.body.append(replaced_p)
            rid += 2
        else:
            # Copy as-is
            dst.element.body.append(copy.deepcopy(child))

        # Inject new paragraphs with insertion marks
        if ci in child_injection_plan:
            for inject_text in child_injection_plan[ci]:
                p = _make_inserted_para(inject_text, rid)
                dst.element.body.append(p)
                rid += 2

    # Now renumber citations in the tracked version too
    # Save intermediate
    tmp_path = os.path.join(OUTPUT_DIR, "PGPH_tracked_tmp.docx")
    dst.save(tmp_path)

    # Apply citation renumbering with tracked changes
    # For simplicity, apply renumbering directly (citations renumbered in clean text)
    from create_pgph_revision import renumber_citations
    out_path = os.path.join(OUTPUT_DIR, "PGPH_Full_Paper_EN_tracked.docx")
    renumber_citations(tmp_path, out_path)
    os.remove(tmp_path)

    print(f"  Saved: {out_path}")
    return out_path


# ═══════════════════════════════════════════════════════════════════════════════
# PART 2: EDITABLE PPTX FIGURES
# ═══════════════════════════════════════════════════════════════════════════════

# Figure mapping: manuscript figure number -> source file + caption
FIGURE_MAP = [
    {
        "num": 1,
        "title": "Figure 1: PRISMA-ScR Flow Diagram",
        "caption": "Figure 1: PRISMA-ScR flow diagram for the scoping review of Hansen's disease screening in work visa medical examinations across 197 countries.",
        "png": "fig4_prisma_flow.png",
    },
    {
        "num": 2,
        "title": "Figure 2: Disease Frequency (Sankey Diagram)",
        "caption": "Figure 2: Sankey diagram showing the frequency of specific diseases named in work visa medical examination requirements across 197 countries.",
        "png": "fig6_sankey_en.png",
    },
    {
        "num": 3,
        "title": "Figure 3: Data Accessibility Map",
        "caption": "Figure 3: Global map of data accessibility for work visa medical examination requirements. Base map: Natural Earth (naturalearthdata.com, public domain).",
        "png": "fig7_accessibility_en.png",
    },
    {
        "num": 4,
        "title": "Figure 4: World Map of Hansen's Disease Screening",
        "caption": "Figure 4: Distribution of Hansen's disease screening in work visa medical examinations across 197 countries. Base map: Natural Earth (naturalearthdata.com, public domain).",
        "png": "fig1_world_map.png",
    },
    {
        "num": 5,
        "title": "Figure 5: Disease Bar Chart",
        "caption": "Figure 5: Number of countries requiring screening for each disease category in work visa medical examinations.",
        "png": "fig2_disease_bar.png",
    },
    {
        "num": 6,
        "title": "Figure 6: Regional Distribution (Donut Chart)",
        "caption": "Figure 6: Regional distribution of countries with disease-specific work visa medical examination requirements.",
        "png": "fig3_regional_donut.png",
    },
    {
        "num": 7,
        "title": "Figure 7: Transmissibility vs. Screening Requirements",
        "caption": "Figure 7: Relationship between disease transmissibility and inclusion in work visa medical examination screening requirements.",
        "png": "fig5_transmissibility.png",
    },
]


def create_editable_pptx():
    """Create editable PPTX with one figure per slide (widescreen 13.333x7.5)."""
    print("Creating editable PPTX figures...")
    prs = Presentation()
    # Set widescreen dimensions
    prs.slide_width = PptxEmu(12192000)   # 13.333 inches
    prs.slide_height = PptxEmu(6858000)   # 7.5 inches

    for fig in FIGURE_MAP:
        png_path = os.path.join(FIGURES_DIR, fig["png"])
        if not os.path.exists(png_path):
            print(f"  WARNING: {png_path} not found, skipping")
            continue

        # Use blank layout
        slide_layout = prs.slide_layouts[6]  # Blank
        slide = prs.slides.add_slide(slide_layout)

        # Add title at top
        from pptx.util import Inches as I, Pt as P
        title_box = slide.shapes.add_textbox(
            I(0.5), I(0.2), I(12.3), I(0.5))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = fig["title"]
        p.font.size = P(18)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER

        # Calculate image dimensions to fit within slide area
        # Available area: width=12.3in, height=5.5in (leaving room for title+caption)
        from PIL import Image
        img = Image.open(png_path)
        img_w, img_h = img.size
        img.close()

        max_w = I(12.0)
        max_h = I(5.3)
        scale_w = max_w / img_w
        scale_h = max_h / img_h
        scale = min(scale_w, scale_h)
        final_w = int(img_w * scale)
        final_h = int(img_h * scale)

        # Center image
        left = (prs.slide_width - final_w) // 2
        top = I(0.8) + (max_h - final_h) // 2

        slide.shapes.add_picture(png_path, left, top, final_w, final_h)

        # Add caption at bottom
        cap_box = slide.shapes.add_textbox(
            I(0.5), I(6.5), I(12.3), I(0.9))
        tf2 = cap_box.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = fig["caption"]
        p2.font.size = P(10)
        p2.alignment = PP_ALIGN.CENTER

    out_path = os.path.join(OUTPUT_DIR, "PGPH_Figures_Editable.pptx")
    prs.save(out_path)
    print(f"  Saved: {out_path}")
    return out_path


# ═══════════════════════════════════════════════════════════════════════════════
# PART 3: COPY PNG FILES TO OUTPUT
# ═══════════════════════════════════════════════════════════════════════════════

def copy_png_figures():
    """Copy PNG figures to output/figures_png/ with manuscript numbering."""
    import shutil
    png_dir = os.path.join(OUTPUT_DIR, "figures_png")
    os.makedirs(png_dir, exist_ok=True)
    for fig in FIGURE_MAP:
        src = os.path.join(FIGURES_DIR, fig["png"])
        dst = os.path.join(png_dir, f"Figure_{fig['num']}.png")
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f"  Copied: {dst}")
    return png_dir


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    create_tracked_changes_manuscript()
    print()
    create_editable_pptx()
    print()
    copy_png_figures()
    print()
    print("Done!")
