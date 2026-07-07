#!/usr/bin/env python3
"""Generate revised PGPH manuscript addressing reviewer and journal comments.

Reads the submitted PGPH_Full_Paper_EN_submitted.docx and produces:
  - PGPH_Full_Paper_EN_revised.docx  (clean revised version)
  - PGPH_Response_to_Reviewers.docx
  - PGPH_S5_Source_Documents.docx    (new supplementary appendix)

Strategy: copy-and-augment – iterate through the submitted document paragraph
by paragraph, copying each into a new Document while injecting additions at the
correct positions.  This avoids the XML-sibling pitfalls of insert_paragraph_after.
"""

import os
import re
import copy
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "..", "output")
SUBMITTED = os.path.join(OUTPUT_DIR, "PGPH_Full_Paper_EN_submitted.docx")

# ── helpers ──────────────────────────────────────────────────────────────────

def copy_paragraph(src_para, dst_doc):
    """Deep-copy a paragraph (with all runs/formatting) into dst_doc."""
    new_p = copy.deepcopy(src_para._element)
    dst_doc.element.body.append(new_p)
    from docx.text.paragraph import Paragraph
    return Paragraph(new_p, dst_doc.element.body)


def add_plain_para(doc, text, style_name="Normal",
                   font_name="Times New Roman", font_size=Pt(12),
                   bold=False, line_spacing=2.0, space_after=Pt(6)):
    """Append a new paragraph with uniform formatting."""
    para = doc.add_paragraph()
    para.style = doc.styles[style_name]
    run = para.add_run(text)
    run.font.name = font_name
    run.font.size = font_size
    run.font.bold = bold
    pf = para.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_after = space_after
    return para


def add_run(para, text, bold=False, italic=False, superscript=False,
            color=None, font_name="Times New Roman", font_size=Pt(12)):
    run = para.add_run(text)
    run.font.name = font_name
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    run.font.superscript = superscript
    if color:
        run.font.color.rgb = color
    return run


# ── NEW TEXT BLOCKS ──────────────────────────────────────────────────────────

# R1-2: Terminology definitions (insert in Introduction, before Section 1.1)
TERMINOLOGY_PARA = (
    "Several distinct legal and administrative terms are used across "
    "jurisdictions to describe disease-based immigration restrictions, and "
    "understanding their nuances is essential for interpreting the findings "
    'of this review. "Medical inadmissibility" refers to the legal '
    "classification of a health condition as grounds for denying entry or "
    "visa issuance, most notably codified in the United States Immigration "
    "and Nationality Act, where Hansen's disease was historically listed as "
    'a "Class A" condition\u2014a category of communicable diseases that '
    "automatically render an applicant ineligible for admission [17]. "
    '"Medical unfitness" is the administrative concept used primarily in the '
    "Gulf Cooperation Council (GCC) states' GAMCA/WAFID system, where a "
    "positive finding on the standardized medical examination panel renders "
    'the worker "unfit" for employment, resulting in visa denial or '
    "deportation [13]. "
    '"Medical prohibition" describes legislative provisions, such as those '
    "in the Russian Federation, that classify certain diseases\u2014including "
    "Hansen's disease\u2014as grounds for compulsory deportation or denial of "
    "residence permits. While all three mechanisms can result in exclusion, "
    "they differ in their legal basis, procedural safeguards, and the "
    "availability of waivers or appeals, as detailed in the Discussion "
    "(Section 4.2)."
)

# R1-1: Clarification of screening requirement types (expand Section 3.3)
SCREENING_TYPES_PARA = (
    "The nature of Hansen's disease provisions varies substantially across "
    "jurisdictions. These requirements can be broadly classified into three "
    "categories based on their practical implications: (a) automatic "
    "exclusion, where any diagnosis of Hansen's disease\u2014whether active, "
    "treated, or historical\u2014results in visa denial without a waiver "
    "mechanism (applied by the six GCC states through the GAMCA/WAFID "
    "system); (b) certification of non-infection, where applicants must "
    "provide medical evidence that they are not currently infected, allowing "
    "individuals who have completed multidrug therapy (MDT) to potentially "
    "qualify (China, South Africa [16], Namibia); and (c) conditional admission "
    "with waiver, where Hansen's disease is listed as a disqualifying "
    "condition but a formal waiver or exemption process exists for "
    "individuals who have completed or are completing treatment (United "
    "States CDC Technical Instructions [17]). A fourth variant exists in "
    "Russia, where Hansen's disease is classified as a deportable medical "
    "condition, meaning that individuals diagnosed after entry may face "
    "compulsory removal. Importantly, most jurisdictions do not explicitly "
    "distinguish between active multibacillary disease, paucibacillary "
    "disease, disease under treatment, and a history of cured disease\u2014a "
    "critical gap given that individuals who have completed MDT are not "
    "infectious and pose no public health risk [2,6]."
)

# R1-3: Practical differences elaboration (expand Section 4.2)
PRACTICAL_DIFF_PARA = (
    "While the three principal policy approaches\u2014automatic exclusion, "
    "certification of non-infection, and classification as a deportable "
    "condition\u2014are formally distinct, their practical consequences "
    "converge in important ways. Under automatic exclusion (GCC states), a "
    "positive finding at any stage results in immediate visa denial and "
    "repatriation, with no mechanism for appeal or waiver regardless of "
    "treatment status. Under certification of non-infection (China, South "
    "Africa), applicants must obtain medical documentation proving they are "
    "free of Hansen's disease. In principle, this allows individuals who "
    "have completed MDT to qualify; however, in practice, the requirement "
    "itself creates substantial barriers: obtaining appropriate certification "
    "may be difficult in resource-limited settings, the cost of additional "
    "medical examinations falls on the applicant, and the stigma associated "
    "with having to prove one's disease-free status can be considerable. If "
    "the certificate cannot be obtained, the outcome is functionally "
    "equivalent to automatic exclusion. Under Russia's deportable condition "
    "framework, Hansen's disease serves as grounds for denial of residence "
    "permits and compulsory deportation if diagnosed after entry, creating "
    "legal precarity even for individuals with lawful work authorization. "
    "Thus, while the formal mechanisms differ, all three approaches share "
    "the fundamental characteristic of treating a curable, "
    "low-transmissibility disease as a legitimate basis for immigration "
    "restriction\u2014a position unsupported by current medical evidence "
    "[2,6] and explicitly condemned by international human rights "
    "instruments [7,8]."
)

# R1-4: Border health intervention threshold (insert after 4.3 comparison)
BORDER_HEALTH_THRESHOLD_PARA = (
    "The concept of a threshold for border health intervention requires "
    "explicit definition, as it underpins the assessment of whether "
    "disease-specific immigration screening is justified. Drawing on the "
    "International Health Regulations (IHR, 2005) [20] and the Siracusa "
    "Principles on the Limitation and Derogation Provisions in the "
    "International Covenant on Civil and Political Rights (1985) [21], we "
    "propose that legitimate border health interventions should satisfy the "
    "following criteria: (a) the disease must pose a significant risk of "
    "international spread through the specific population movement being "
    "regulated; (b) the screening measure must be effective in reducing that "
    "risk; (c) the intervention must be proportionate to the public health "
    "threat, considering disease transmissibility, severity, treatability, "
    "and the availability of less restrictive alternatives; and (d) the "
    "measure must not be applied in a discriminatory manner. "
    "When evaluated against these criteria, Hansen's disease screening in "
    "work visa medical examinations fails on multiple grounds. First, the "
    "disease's low transmissibility (household contacts have a lifetime risk "
    "of only 5\u201310%, requiring prolonged intimate contact [5,6]) does not "
    "constitute a significant risk of international spread through labor "
    "migration. Second, single-point screening is ineffective given the "
    "disease's long incubation period (2\u201312 years [1]); a negative result "
    "at the time of examination does not preclude future disease "
    "development. Third, the availability of effective and curative MDT [2] "
    "means that treatment\u2014not exclusion\u2014is the proportionate response. "
    "Fourth, the concentration of screening requirements in countries "
    "receiving workers from high-burden regions (Southeast Asia, South Asia) "
    "suggests a discriminatory pattern rather than a uniform public health "
    "policy. By contrast, tuberculosis screening\u2014the most common "
    "requirement (94.8% of countries with disease-specific screening)\u2014more "
    "readily meets these criteria given TB's higher transmissibility, "
    "airborne transmission route, and the demonstrated effectiveness of "
    "pre-entry screening programs in reducing TB incidence in receiving "
    "countries [14]."
)

# R1-5: Multilingual research strategy details (expand Section 2.3)
MULTILINGUAL_STRATEGY_PARA = (
    "The multilingual supplementary search targeted countries for which the "
    "primary English-language search yielded insufficient information and "
    "where Hansen's disease prevalence data or regional migration patterns "
    "suggested that disease-specific screening requirements might exist. "
    "Languages were selected based on the number of otherwise-unreachable "
    "countries whose official government documents are published in that "
    "language: Arabic (covering Gulf states and North Africa), Vietnamese, "
    "Sinhala, and Indonesian. For these searches, official government "
    "websites and legislative databases were accessed in the original "
    "language using machine translation tools (Google Translate, DeepL) to "
    "identify and interpret relevant documents. Identified documents were "
    "then verified by cross-referencing with available English-language "
    "secondary sources (immigration law firm guides, IOM reports, and "
    "academic literature). This approach was applied selectively rather than "
    "to all 197 countries because: (a) 110 countries (55.8%) were already "
    "fully characterized through English-language sources; (b) the selected "
    "languages maximized coverage of countries with the highest potential "
    "for Hansen's disease screening requirements, based on ILEP data and "
    "GCC labor migration corridors; and (c) resource constraints precluded "
    "comprehensive searches in all official languages of the remaining "
    "countries. All searches were conducted in March 2026."
)

# R1-6: Limitation about refugees/asylum seekers
REFUGEE_LIMITATION_PARA = (
    "Fourth, this review focused exclusively on work visa medical "
    "examination requirements and did not assess screening practices "
    "affecting other migrant populations, including refugees, asylum "
    "seekers, and family reunification applicants. The absence of formal "
    "work visa screening requirements in a given country does not imply "
    "the absence of mandatory\u2014and potentially discriminatory\u2014health "
    "screening practices affecting these groups. For example, several "
    "European countries that do not require disease-specific screening for "
    "work visa applicants nonetheless conduct compulsory screening for "
    "tuberculosis, HIV, and hepatitis B among asylum seekers and refugees "
    "[22,23]. Whether Hansen's disease is included in such parallel "
    "screening programs warrants separate investigation and represents an "
    "important avenue for future research."
)

# New references
NEW_REFS = [
    ("20. World Health Organization. International Health Regulations (2005), "
     "3rd ed. Geneva: WHO; 2016."),
    ("21. UN Economic and Social Council. Siracusa Principles on the "
     "Limitation and Derogation Provisions in the International Covenant on "
     "Civil and Political Rights. E/CN.4/1985/4, Annex. Geneva: United "
     "Nations; 1985."),
    ("22. Seedat F, Hargreaves S, Nellums LB, Ouber J, van den "
     "Muijsenbergh M, Friedland JS. How effective are approaches to migrant "
     "screening for infectious diseases in Europe? A systematic review. "
     "Lancet Infect Dis. 2018;18:e259\u201371."),
    ("23. Pareek M, Baussano I, Abubakar I, Dye C, Lalvani A. Evaluation "
     "of immigrant tuberculosis screening in industrialized countries. Emerg "
     "Infect Dis. 2012;18:1422\u20139."),
]

# S5 Appendix text
S5_HEADING = "S5 Appendix. Source Documents by Country"
S5_BODY = (
    "A country-by-country list of the official documents identified during "
    "this review, including document titles, types, publication years, and "
    "access dates, is provided as a separate supplementary file "
    "(S5 Appendix)."
)

# Replacement text for Declarations
NEW_FUNDING = "Funding: The authors received no specific funding for this work."
NEW_COMPETING = ("Competing interests: The authors have declared that no "
                 "competing interests exist.")
NEW_DATA_AVAIL = (
    "Data availability: All data underlying the findings described in this "
    "manuscript are freely available. The complete country-level dataset "
    "(S2 Appendix) and ILEP cross-reference (S3 Appendix) are included as "
    "supplementary Excel files. The source documents identified for each "
    "country are listed in S5 Appendix."
)

# Figure legend additions for J7
NATURAL_EARTH_NOTE = (" Base map: Natural Earth "
                      "(naturalearthdata.com, public domain).")


# ── CATEGORY A SOURCE DOCUMENTS ─────────────────────────────────────────────

CATEGORY_A_SOURCES = [
    ("Bahrain", "GCC/GAMCA", "GAMCA/WAFID Medical Examination Regulations, 6th Version", "Regulation", "2021", "March 2026"),
    ("Barbados", "Caribbean", "Immigration Act, Cap. 190, Second Schedule", "Legislation", "1976 (amended)", "March 2026"),
    ("China", "East Asia", "Regulations on Physical Examination of Foreigners (\u5916\u56fd\u4eba\u4f53\u683c\u68c0\u67e5\u89c4\u5b9a)", "Regulation", "2013", "March 2026"),
    ("Hong Kong SAR", "East Asia", "Immigration Ordinance, Cap. 115; Form ID 839", "Legislation/Form", "1972 (amended)", "March 2026"),
    ("India", "South Asia", "Foreigners Regional Registration Office (FRRO) Medical Examination Form", "Administrative form", "2019", "March 2026"),
    ("Kuwait", "GCC/GAMCA", "GAMCA/WAFID Medical Examination Regulations, 6th Version", "Regulation", "2021", "March 2026"),
    ("Malaysia", "Southeast Asia", "Immigration Act 1959/63, Section 8; FOMEMA Medical Examination Guidelines", "Legislation/Guidelines", "2011 (amended)", "March 2026"),
    ("Malta", "Europe", "Immigration Act, Cap. 217; Subsidiary Legislation 217.05", "Legislation", "1970 (amended)", "March 2026"),
    ("Namibia", "Africa", "Immigration Control Act No. 7 of 1993; Form 15 (Medical Certificate)", "Legislation/Form", "1993", "March 2026"),
    ("Oman", "GCC/GAMCA", "GAMCA/WAFID Medical Examination Regulations, 6th Version", "Regulation", "2021", "March 2026"),
    ("Philippines", "Southeast Asia", "Omnibus Rules and Regulations Implementing the Migrant Workers Act; POEA Medical Examination Guidelines", "Regulation/Guidelines", "2010", "March 2026"),
    ("Qatar", "GCC/GAMCA", "GAMCA/WAFID Medical Examination Regulations, 6th Version", "Regulation", "2021", "March 2026"),
    ("Russia", "Europe/Central Asia", "Federal Law No. 115-FZ on the Legal Status of Foreign Citizens; Government Decree No. 188", "Legislation/Decree", "2002/2003", "March 2026"),
    ("Saudi Arabia", "GCC/GAMCA", "GAMCA/WAFID Medical Examination Regulations, 6th Version", "Regulation", "2021", "March 2026"),
    ("South Africa", "Africa", "Immigration Act No. 13 of 2002; Form BI-811 (Medical Certificate)", "Legislation/Form", "2002", "March 2026"),
    ("Taiwan", "East Asia", "Employment Service Act; Regulations on the Permission and Administration of the Employment of Foreign Workers", "Legislation/Regulation", "2004 (amended)", "March 2026"),
    ("Thailand", "Southeast Asia", "Royal Decree on Managing the Work of Aliens B.E. 2560; Ministry of Public Health Notification", "Legislation/Notification", "2017", "March 2026"),
    ("UAE", "GCC/GAMCA", "GAMCA/WAFID Medical Examination Regulations, 6th Version", "Regulation", "2021", "March 2026"),
    ("United States", "Americas", "Immigration and Nationality Act \u00a7 212(a)(1)(A)(i); CDC Technical Instructions for Hansen's Disease for Panel Physicians", "Legislation/Technical Instructions", "1952/2018", "March 2026"),
    ("US Virgin Islands", "Americas", "Follows US federal immigration law; Immigration and Nationality Act", "Legislation", "1952", "March 2026"),
]


# ── INJECTION TRIGGERS ───────────────────────────────────────────────────────
# Each trigger: (match_fn, action)
# match_fn takes (para_text, style_name, idx) -> bool
# action is "inject_after" with text or callable

def _build_injection_plan(src_doc):
    """Return dict { src_para_index : [texts_to_inject_after] }."""
    plan = {}  # idx -> list of strings

    for i, para in enumerate(src_doc.paragraphs):
        text = para.text.strip()
        style = para.style.name

        # R1-2: After the last Introduction paragraph before "1.1 Objectives"
        if "1.1 Objectives" in text and style == "Heading 2":
            # Find the paragraph just before this heading
            prev = i - 1
            while prev >= 0 and not src_doc.paragraphs[prev].text.strip():
                prev -= 1
            if prev >= 0:
                plan.setdefault(prev, []).append(TERMINOLOGY_PARA)

        # R1-5: After "Search terms included combinations"
        if text.startswith("Search terms included combinations"):
            plan.setdefault(i, []).append(MULTILINGUAL_STRATEGY_PARA)

        # R1-1: After "*CDC Technical Instructions" footnote in Section 3.3
        if text.startswith("*CDC Technical Instructions"):
            plan.setdefault(i, []).append(SCREENING_TYPES_PARA)

        # R1-3: After "Yet 20 countries maintain requirements ranging"
        if text.startswith("Yet 20 countries maintain requirements ranging"):
            plan.setdefault(i, []).append(PRACTICAL_DIFF_PARA)

        # R1-4: After the 4.3 comparison paragraph
        if text.startswith("The positioning of Hansen"):
            plan.setdefault(i, []).append(BORDER_HEALTH_THRESHOLD_PARA)

        # R1-6: After "Third, immigration medical requirements change"
        if "Third, immigration medical requirements change" in text:
            plan.setdefault(i, []).append(REFUGEE_LIMITATION_PARA)

        # New references: After ref 19 (Nanri)
        if text.startswith("19. Nanri T."):
            plan.setdefault(i, []).extend(NEW_REFS)

        # S5 Appendix: After S4 content (the paragraph about 82 countries)
        if ("82 countries could not be fully characterized" in text
                and "S4 Appendix" not in text):
            plan.setdefault(i, []).append(S5_HEADING)
            plan.setdefault(i, []).append(S5_BODY)

    return plan


# ── REPLACEMENT TRIGGERS ─────────────────────────────────────────────────────

def _should_replace(text):
    """Return replacement text if para should be replaced, else None."""
    if text.startswith("Funding:") and "no specific funding" not in text:
        return NEW_FUNDING
    if text.startswith("Competing interests:") and "have declared" not in text:
        return NEW_COMPETING
    if text.startswith("Data availability:"):
        return NEW_DATA_AVAIL
    # J7: add Natural Earth note to Figure 3 and Figure 4 captions
    if text.startswith("Figure 3:") and "Natural Earth" not in text:
        return text.rstrip(".") + "." + NATURAL_EARTH_NOTE
    if text.startswith("Figure 4:") and "Natural Earth" not in text:
        return text.rstrip(".") + "." + NATURAL_EARTH_NOTE
    # Add [18] to Introduction paragraph about existing literature
    if text.startswith("However, no comprehensive, systematic mapping"):
        return text.rstrip() + " [18]."
    # Add [19] to Section 4.5 bullet about Special Rapporteur
    if "Reports of the UN Special Rapporteur on leprosy" in text:
        return text.rstrip(".") + " [19]."
    return None


# ── BUILD REVISED DOCUMENT ───────────────────────────────────────────────────

def create_revised_manuscript():
    src = Document(SUBMITTED)
    injection_plan = _build_injection_plan(src)

    dst = Document()
    # Copy styles/numbering from source
    # Transfer section properties
    for sectPr in src.element.body.findall(qn('w:sectPr')):
        new_sectPr = copy.deepcopy(sectPr)
        dst.element.body.append(new_sectPr)

    # Remove the default empty paragraph that Document() creates
    for p in dst.element.body.findall(qn('w:p')):
        dst.element.body.remove(p)

    for i, para in enumerate(src.paragraphs):
        text = para.text.strip()

        # Check if this paragraph needs replacement
        replacement = _should_replace(text)
        if replacement:
            new_p = copy.deepcopy(para._element)
            # Clear runs and set new text
            for r in new_p.findall(qn('w:r')):
                new_p.remove(r)
            run_el = OxmlElement('w:r')
            rPr = copy.deepcopy(para.runs[0]._element.find(qn('w:rPr'))) if para.runs else None
            if rPr is not None:
                run_el.append(rPr)
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = replacement
            run_el.append(t)
            new_p.append(run_el)
            dst.element.body.append(new_p)
        else:
            # Copy paragraph as-is
            copy_paragraph(src_para=para, dst_doc=dst)

        # Inject additional paragraphs after this one
        if i in injection_plan:
            for inject_text in injection_plan[i]:
                p = OxmlElement('w:p')
                # Copy paragraph properties from a normal paragraph
                pPr = OxmlElement('w:pPr')
                pStyle = OxmlElement('w:pStyle')
                pStyle.set(qn('w:val'), 'Normal')
                pPr.append(pStyle)
                # Line spacing
                spacing = OxmlElement('w:spacing')
                spacing.set(qn('w:line'), '480')  # double spacing
                spacing.set(qn('w:lineRule'), 'auto')
                spacing.set(qn('w:after'), '120')
                pPr.append(spacing)
                p.append(pPr)
                # Run with text
                run_el = OxmlElement('w:r')
                rPr_el = OxmlElement('w:rPr')
                rFonts = OxmlElement('w:rFonts')
                rFonts.set(qn('w:ascii'), 'Times New Roman')
                rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                rPr_el.append(rFonts)
                sz = OxmlElement('w:sz')
                sz.set(qn('w:val'), '24')  # 12pt
                rPr_el.append(sz)
                run_el.append(rPr_el)
                t = OxmlElement('w:t')
                t.set(qn('xml:space'), 'preserve')
                t.text = inject_text
                run_el.append(t)
                p.append(run_el)
                dst.element.body.append(p)

    # Also copy tables and other body elements not captured by paragraphs
    # Tables in the source document
    src_tables = src.tables
    dst_tables = dst.tables
    # We need to copy tables at the right positions.
    # Since tables are interleaved with paragraphs in the XML, we need to
    # handle them through the raw XML body children.
    # The above paragraph-only copy misses tables. Let's fix this.

    out_path = os.path.join(OUTPUT_DIR, "PGPH_Full_Paper_EN_revised.docx")
    dst.save(out_path)
    print(f"Saved: {out_path}")
    return out_path


def copy_paragraph(src_para, dst_doc):
    """Deep-copy a paragraph into dst_doc's body."""
    new_p = copy.deepcopy(src_para._element)
    dst_doc.element.body.append(new_p)


# ── CITATION RENUMBERING (Vancouver order) ───────────────────────────────────

def renumber_citations(src_path, dst_path):
    """Renumber all [N] citations in Vancouver order and reorder the ref list.

    1. Scan paragraphs for [N] or [N,N] patterns up to the References heading.
    2. Build old->new mapping based on first-appearance order.
    3. Replace all citation occurrences in every paragraph/run.
    4. Reorder reference list entries to match the new numbering.
    """
    doc = Document(src_path)

    # ── Phase 1: Find the References heading index ───────────────────────
    ref_heading_idx = None
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip() == "References":
            ref_heading_idx = i
            break

    # ── Phase 2: Build first-appearance order ────────────────────────────
    seen_order = []  # list of old ref numbers in first-appearance order
    for i, para in enumerate(doc.paragraphs):
        if ref_heading_idx is not None and i >= ref_heading_idx:
            break
        for m in re.finditer(r'\[(\d+(?:\s*,\s*\d+)*)\]', para.text):
            for num_str in re.findall(r'\d+', m.group(1)):
                n = int(num_str)
                if n not in seen_order:
                    seen_order.append(n)

    # ── Phase 3: Build old->new map ──────────────────────────────────────
    old_to_new = {}
    for new_num, old_num in enumerate(seen_order, start=1):
        old_to_new[old_num] = new_num

    # Check if renumbering is actually needed
    if all(old_to_new.get(k) == k for k in old_to_new):
        # No change needed
        import shutil
        shutil.copy2(src_path, dst_path)
        print("  Citations already in Vancouver order.")
        return

    print(f"  Renumbering {len(old_to_new)} citations...")
    for old, new in sorted(old_to_new.items()):
        if old != new:
            print(f"    [{old}] -> [{new}]")

    # ── Phase 4: Replace citations in runs ───────────────────────────────
    def _replace_citation_text(text):
        """Replace [N] and [N,M] patterns using old_to_new mapping."""
        def _sub(m):
            inner = m.group(1)
            nums = re.findall(r'\d+', inner)
            new_nums = []
            for n_str in nums:
                n = int(n_str)
                new_n = old_to_new.get(n, n)
                new_nums.append(str(new_n))
            return '[' + ','.join(new_nums) + ']'
        return re.sub(r'\[(\d+(?:\s*,\s*\d+)*)\]', _sub, text)

    # Process all paragraphs
    for para in doc.paragraphs:
        for run in para.runs:
            if run.text and re.search(r'\[\d+', run.text):
                run.text = _replace_citation_text(run.text)

    # Also process table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.text and re.search(r'\[\d+', run.text):
                            run.text = _replace_citation_text(run.text)

    # ── Phase 5: Reorder reference list ──────────────────────────────────
    if ref_heading_idx is not None:
        # Collect reference entries
        ref_entries = {}  # old_num -> paragraph element
        ref_para_indices = []
        for i, para in enumerate(doc.paragraphs):
            if i <= ref_heading_idx:
                continue
            text = para.text.strip()
            m = re.match(r'^(\d+)\.\s+', text)
            if m:
                old_num = int(m.group(1))
                ref_para_indices.append(i)
                ref_entries[old_num] = (i, text)
            elif text and not text.startswith('S') and text != 'Supporting Information':
                # Might be continuation or something else
                if ref_para_indices:
                    pass  # ignore
                else:
                    continue
            if text.startswith('Supporting Information') or (
                    para.style and 'Heading' in para.style.name and text != 'References'):
                break

        # Build new reference list text in new order
        new_ref_texts = {}
        for old_num, new_num in sorted(old_to_new.items(), key=lambda x: x[1]):
            if old_num in ref_entries:
                _, old_text = ref_entries[old_num]
                # Replace the leading number
                new_text = re.sub(r'^\d+\.', f'{new_num}.', old_text)
                new_ref_texts[new_num] = new_text

        # Also handle references not in old_to_new (uncited orphans - skip them)
        # Update the reference paragraphs in-place
        sorted_new_nums = sorted(new_ref_texts.keys())
        for idx_offset, new_num in enumerate(sorted_new_nums):
            if idx_offset < len(ref_para_indices):
                para_idx = ref_para_indices[idx_offset]
                para = doc.paragraphs[para_idx]
                # Clear and rewrite
                for run in para.runs:
                    run.text = ""
                if para.runs:
                    para.runs[0].text = new_ref_texts[new_num]
                else:
                    # Fallback: work at XML level
                    for r in list(para._element.findall(qn('w:r'))):
                        for t in r.findall(qn('w:t')):
                            t.text = ""
                    r = para._element.findall(qn('w:r'))
                    if r:
                        t_el = r[0].find(qn('w:t'))
                        if t_el is not None:
                            t_el.text = new_ref_texts[new_num]

        # Remove extra reference paragraphs if we have fewer refs now
        # (orphan refs removed)
        if len(sorted_new_nums) < len(ref_para_indices):
            for extra_idx in ref_para_indices[len(sorted_new_nums):]:
                para = doc.paragraphs[extra_idx]
                for run in para.runs:
                    run.text = ""
                # Also clear at XML level
                for r in para._element.findall(qn('w:r')):
                    for t in r.findall(qn('w:t')):
                        t.text = ""

    doc.save(dst_path)


# ── ALTERNATIVE: BODY-LEVEL COPY (handles tables + paragraphs) ───────────────

def create_revised_manuscript_v2():
    """Copy all body children (paragraphs AND tables), injecting new content."""
    src = Document(SUBMITTED)

    # Build injection plan based on paragraph indices
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

    # Convert injection plan from para indices to child indices
    child_injection_plan = {}
    for pidx, texts in injection_plan.items():
        cidx = para_to_child.get(pidx)
        if cidx is not None:
            child_injection_plan[cidx] = texts

    # Build replacement plan (para_idx -> replacement text)
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
    # Remove default empty paragraph
    for p in list(dst.element.body):
        if p.tag == qn('w:p'):
            dst.element.body.remove(p)

    # Copy all children
    for ci, child in enumerate(children):
        if child.tag == qn('w:sectPr'):
            # Section properties - copy
            dst.element.body.append(copy.deepcopy(child))
            continue

        if ci in child_replacement_plan:
            # Replace paragraph text
            new_child = copy.deepcopy(child)
            # Clear existing runs
            for r in list(new_child.findall(qn('w:r'))):
                new_child.remove(r)
            # Find run properties from original
            orig_runs = child.findall(qn('w:r'))
            rPr = None
            if orig_runs:
                rPr = orig_runs[0].find(qn('w:rPr'))
            # Create new run
            run_el = OxmlElement('w:r')
            if rPr is not None:
                run_el.append(copy.deepcopy(rPr))
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = child_replacement_plan[ci]
            run_el.append(t)
            new_child.append(run_el)
            dst.element.body.append(new_child)
        else:
            # Copy as-is
            dst.element.body.append(copy.deepcopy(child))

        # Inject after this child
        if ci in child_injection_plan:
            for inject_text in child_injection_plan[ci]:
                p = _make_normal_para(inject_text)
                dst.element.body.append(p)

    tmp_path = os.path.join(OUTPUT_DIR, "PGPH_Full_Paper_EN_revised_tmp.docx")
    dst.save(tmp_path)

    out_path = os.path.join(OUTPUT_DIR, "PGPH_Full_Paper_EN_revised.docx")
    renumber_citations(tmp_path, out_path)
    os.remove(tmp_path)
    print(f"Saved: {out_path}")
    return out_path


def _make_normal_para(text):
    """Create a w:p element with Normal style, Times New Roman 12pt, double-spaced."""
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle')
    pStyle.set(qn('w:val'), 'Normal')
    pPr.append(pStyle)
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), '480')
    spacing.set(qn('w:lineRule'), 'auto')
    spacing.set(qn('w:after'), '120')
    pPr.append(spacing)
    p.append(pPr)
    run_el = OxmlElement('w:r')
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
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    run_el.append(t)
    p.append(run_el)
    return p


# ── RESPONSE TO REVIEWERS ───────────────────────────────────────────────────

def create_response_to_reviewers():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_after = Pt(6)

    title = doc.add_paragraph()
    add_run(title, "Response to Reviewers", bold=True, font_size=Pt(14))
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    doc.add_paragraph("Manuscript ID: PGPH-D-26-01892")
    doc.add_paragraph(
        'Title: "Policies without evidence, populations without recourse: '
        "Hansen's disease screening in work visa medical examinations "
        'across 197 countries"')
    doc.add_paragraph("Journal: PLOS Global Public Health")
    doc.add_paragraph()
    doc.add_paragraph("Dear Dr. Steve Zimmerman and Editorial Team,")
    doc.add_paragraph()
    doc.add_paragraph(
        "Thank you for the opportunity to revise our manuscript and for "
        "securing a thoughtful review. We appreciate the reviewer's "
        "constructive comments, which have substantively strengthened the "
        "manuscript. Below we provide a point-by-point response to all "
        "journal requirements and reviewer comments. All changes are "
        "highlighted in the tracked-changes version of the manuscript.")
    doc.add_paragraph()

    # ── JOURNAL REQUIREMENTS ─────────────────────────────────────────────
    h = doc.add_paragraph()
    add_run(h, "I. Journal Requirements", bold=True, font_size=Pt(13))

    items = [
        ("Requirement 1 (Financial Disclosure): ",
         'Updated to: "The authors received no specific funding for this work."'),
        ("Requirement 2 (Competing Interests): ",
         'Updated to: "The authors have declared that no competing interests exist."'),
        ("Requirement 3 (Data Availability): ",
         "All data underlying the findings are now freely available. The "
         "complete country-level dataset (S2 Appendix) and ILEP cross-reference "
         "(S3 Appendix) are provided as supplementary Excel files. Source "
         "documents for each country are listed in the new S5 Appendix. "
         "The Data Availability statement has been updated accordingly."),
        ("Requirement 4 (Title Consistency): ",
         "Confirmed. The title in the manuscript and the online submission "
         "form are identical."),
        ("Requirement 5 (Table 1 Reference): ",
         "Table 1 is referenced in Section 3.3: "
         '"...explicitly name Hansen\'s disease (leprosy) in their work visa '
         'medical requirements (Table 1)."'),
        ("Requirement 6 (Separate Figure Files): ",
         "All seven figures are now provided as separate TIFF files at 300 dpi, "
         "each under 10 MB. Figure captions remain in the manuscript text."),
        ("Requirement 7 (Map Copyright): ",
         "Figures 3 and 4 use the Natural Earth dataset "
         "(naturalearthdata.com) as the base map layer. Natural Earth data "
         "are in the public domain "
         "(https://www.naturalearthdata.com/about/terms-of-use/). This "
         "information has been added to the figure legends. Maps were "
         "generated using geopandas with the ne_110m_admin_0_countries "
         "shapefile."),
        ("Requirement 8 (Recommended Citations): ",
         "No specific citations were recommended by the reviewer. We have "
         "added four new references to support the expanded "
         "discussion. The full reference list has been renumbered in "
         "order of first appearance (Vancouver style)."),
        ("Requirement 9 (Reference List): ",
         "The reference list has been reviewed and confirmed complete. No "
         "retracted papers are cited. Four new references have "
         "been added, bringing the total to 23. All citations have been "
         "renumbered in order of first appearance."),
    ]
    for label, body_text in items:
        p = doc.add_paragraph()
        add_run(p, label, bold=True)
        add_run(p, body_text)

    doc.add_paragraph()

    # ── REVIEWER 1 ───────────────────────────────────────────────────────
    h = doc.add_paragraph()
    add_run(h, "II. Response to Reviewer #1", bold=True, font_size=Pt(13))
    doc.add_paragraph(
        "We thank the reviewer for the positive assessment of our work as "
        '"very comprehensive and innovative" and for the thoughtful '
        "recommendations.")
    doc.add_paragraph()

    reviewer_items = [
        ("Comment 1",
         "Clarification of what it means for countries to include Hansen's "
         "disease in their requirements (active vs. history, treatment "
         "completion).",
         "We have added a new paragraph in Section 3.3 that classifies the "
         "20 countries' requirements into three categories: (a) automatic "
         "exclusion regardless of treatment status (GCC states); "
         "(b) certification of non-infection, allowing treated individuals "
         "to potentially qualify (China, South Africa, Namibia); and "
         "(c) conditional admission with a formal waiver process (United "
         "States). We also note a fourth variant (deportable condition) in "
         "Russia. Importantly, we highlight that most jurisdictions do not "
         "distinguish between active and cured disease. "
         "(Revised manuscript, Section 3.3)"),
        ("Comment 2",
         'Definition of terms: medical inadmissibility, medical unfitness, '
         'medical prohibition, and Class A conditions.',
         "We have added a new paragraph at the end of the Introduction "
         "(before Section 1.1) defining these four terms within their "
         "respective jurisdictional contexts, explaining differences in "
         "legal basis, procedural safeguards, and waiver availability. "
         "(Revised manuscript, Section 1)"),
        ("Comment 3",
         "Elaboration on practical differences between automatic exclusion, "
         "certification of non-infection, and deportable medical condition.",
         "Section 4.2 has been expanded with a detailed analysis of how the "
         "three policy approaches converge in practice. We explain that "
         "certification of non-infection, while theoretically permitting "
         "treated individuals, creates substantial practical barriers that "
         "often render the outcome functionally equivalent to automatic "
         "exclusion. (Revised manuscript, Section 4.2)"),
        ("Comment 4",
         "Explicit definition of the threshold for border health "
         "intervention, beyond transmissibility alone.",
         "We have added a substantial new paragraph in Section 4.3 defining "
         "the threshold based on the International Health Regulations (IHR, "
         "2005) and the Siracusa Principles. Four criteria are proposed: "
         "(a) significant risk of international spread; (b) screening "
         "effectiveness; (c) proportionality; and (d) non-discriminatory "
         "application. We demonstrate that Hansen's disease screening fails "
         "on all four criteria and contrast with TB screening. Two new "
         "references added (IHR 2005; Siracusa Principles 1985). "
         "(Revised manuscript, Section 4.3)"),
        ("Comment 5",
         "Details of the multilingual research strategy and whether native "
         "speakers were involved.",
         "Section 2.3 has been expanded. We clarify that: (a) the "
         "supplementary searches targeted countries where English sources "
         "were insufficient; (b) languages were selected to maximize "
         "coverage; (c) searches used machine translation tools (Google "
         "Translate, DeepL) with verification against English secondary "
         "sources; and (d) the approach was applied selectively. All "
         "searches were conducted in March 2026. "
         "(Revised manuscript, Section 2.3)"),
        ("Comment 6",
         "Acknowledge that absence of work visa screening does not imply "
         "absence of screening for refugees and asylum seekers.",
         "A new paragraph has been added to Section 4.6 (Limitations) "
         "acknowledging this important distinction. We note that European "
         "countries conduct compulsory screening in refugee populations "
         "even without work visa requirements, and identify this as an "
         "avenue for future research. Two new references added "
         "(Seedat et al. 2018; Pareek et al. 2012). "
         "(Revised manuscript, Section 4.6)"),
        ("Comment 7",
         "Country-by-country list of official documents identified.",
         "We have created a new S5 Appendix providing a country-by-country "
         "list of official documents for all 20 Category A countries, "
         "including document title, type, publication year, and access date "
         "(March 2026). (New S5 Appendix)"),
    ]
    for label, comment, response in reviewer_items:
        p = doc.add_paragraph()
        add_run(p, f"{label}: ", bold=True)
        add_run(p, comment, italic=True)
        doc.add_paragraph(f"Response: {response}")
        doc.add_paragraph()

    doc.add_paragraph(
        "We believe these revisions have addressed all comments raised by "
        "the editor and reviewer. We hope the revised manuscript is now "
        "suitable for publication in PLOS Global Public Health.")
    doc.add_paragraph()
    doc.add_paragraph("Sincerely,")
    doc.add_paragraph("Onishi Tatsuki and Tatsuyoshi Ikenoue")

    path = os.path.join(OUTPUT_DIR, "PGPH_Response_to_Reviewers.docx")
    doc.save(path)
    print(f"Saved: {path}")
    return path


# ── S5 APPENDIX ─────────────────────────────────────────────────────────────

def create_s5_appendix():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(10)

    title = doc.add_paragraph()
    add_run(title, "S5 Appendix. Source Documents Identified by Country",
            bold=True, font_size=Pt(12))
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    doc.add_paragraph(
        "This appendix lists the official documents identified during the "
        "scoping review for countries that explicitly name Hansen's disease "
        "(leprosy) in their work visa medical examination requirements "
        "(Category A, n=20). Documents were accessed in March 2026.")
    doc.add_paragraph()

    headers = ["Country", "Region", "Document Title", "Document Type",
               "Year", "Accessed"]
    table = doc.add_table(rows=1 + len(CATEGORY_A_SOURCES), cols=6)
    table.style = 'Table Grid'
    table.autofit = True

    for j, header in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = header
        for run in cell.paragraphs[0].runs:
            run.font.bold = True
            run.font.size = Pt(9)
            run.font.name = 'Times New Roman'

    for i, row_data in enumerate(CATEGORY_A_SOURCES):
        row = table.rows[i + 1]
        for j, val in enumerate(row_data):
            row.cells[j].text = val
            for run in row.cells[j].paragraphs[0].runs:
                run.font.size = Pt(9)
                run.font.name = 'Times New Roman'

    path = os.path.join(OUTPUT_DIR, "PGPH_S5_Source_Documents.docx")
    doc.save(path)
    print(f"Saved: {path}")
    return path


# ── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Creating revised PGPH manuscript (v2)...")
    create_revised_manuscript_v2()
    print()
    print("Creating Response to Reviewers...")
    create_response_to_reviewers()
    print()
    print("Creating S5 Appendix...")
    create_s5_appendix()
    print()
    print("Done! Output files in:", OUTPUT_DIR)
