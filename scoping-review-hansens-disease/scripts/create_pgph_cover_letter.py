#!/usr/bin/env python3
"""Generate cover letter for PLOS Global Public Health (PGPH) transfer submission."""

from docx import Document
from docx.shared import Pt, Cm
import os

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def create_cover_letter():
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.space_after = Pt(6)

    p = doc.add_paragraph()
    p.add_run('[Date]')

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.add_run('Editorial Office\nPLOS Global Public Health')

    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run('Re: ')
    run.bold = True
    p.add_run(
        'Transfer submission — "Policies without evidence, populations without recourse: '
        'Hansen\u2019s disease screening in work visa medical examinations across 197 countries"'
    )

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.add_run('Dear Editors,')

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.add_run(
        'We are pleased to submit the above manuscript, transferred from PLOS Neglected Tropical '
        'Diseases, for consideration as a Research Article in PLOS Global Public Health. The '
        'manuscript has not been revised since its original submission; we believe it is well suited '
        'to the scope and mission of PLOS Global Public Health and present it for your evaluation.'
    )

    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run('Why this manuscript fits PLOS Global Public Health')
    run.bold = True

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.add_run(
        'This scoping review systematically maps Hansen\u2019s disease (leprosy) screening requirements '
        'in work visa medical examinations across all 197 UN-recognized countries and territories. '
        'We found that 20 countries continue to name Hansen\u2019s disease in immigration medical '
        'requirements, despite the disease being fully curable, of low transmissibility, and '
        'despite UN resolutions calling for elimination of such discrimination. The six Gulf '
        'Cooperation Council states, through the unified GAMCA/WAFID system, collectively screen '
        'over 35 million migrant workers for leprosy as a condition of "medical unfitness."'
    )

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.add_run(
        'The manuscript addresses themes central to PLOS Global Public Health\u2019s mission:'
    )

    items = [
        'Health inequity: Hansen\u2019s disease screening disproportionately burdens migrant workers '
        'from South and Southeast Asia, populations already facing structural disadvantages in '
        'global labor migration.',
        'Evidence\u2013policy gap: Hansen\u2019s disease ranks as the 5th most commonly screened condition '
        'in work visa medical examinations, despite having the lowest transmissibility among '
        'commonly screened diseases\u2014driven by historical stigma rather than epidemiological evidence.',
        'Global health governance: The persistence of these requirements contravenes UN General '
        'Assembly Resolution A/RES/65/215 (2010) and WHO\u2019s Global Leprosy Strategy 2021\u20132030, '
        'highlighting a failure of international policy implementation.',
        'Migration and health: The study contributes to a growing body of work at PLOS Global '
        'Public Health examining how visa and immigration policies function as structural '
        'determinants of health, complementing recent publications in the journal on visa '
        'restrictions and health equity.',
    ]
    for item in items:
        doc.add_paragraph(item, style='List Bullet')

    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run('Key findings')
    run.bold = True

    doc.add_paragraph()

    findings = [
        '20 of 197 countries/territories (10.2%) explicitly name Hansen\u2019s disease in work visa '
        'medical requirements.',
        'The GCC/GAMCA system affects over 35 million migrant workers with automatic exclusion '
        'for leprosy.',
        'Hansen\u2019s disease screening shows no correlation with national leprosy burden, indicating '
        'stigma-driven rather than evidence-based policy.',
        '82 countries (41.6%) could not be fully characterized, revealing a transparency gap in '
        'immigration health policy.',
    ]
    for f in findings:
        doc.add_paragraph(f, style='List Bullet')

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.add_run(
        'All authors have approved the manuscript for submission. The work is original, has not '
        'been published elsewhere, and is not under consideration at any other journal. The '
        'authors declare no competing interests.'
    )

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.add_run(
        'We believe this study will be of interest to PLOS Global Public Health\u2019s readership '
        'of researchers, health professionals, and policy-makers working at the intersection of '
        'migration, infectious disease, and human rights. We look forward to your editorial '
        'consideration.'
    )

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.add_run('Sincerely,')

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.add_run('[Corresponding Author Name]\n[Affiliation]\n[Email]')

    output_path = os.path.join(OUTPUT_DIR, "PGPH_Cover_Letter_EN.docx")
    doc.save(output_path)
    print(f"Saved: {output_path}")
    return output_path


if __name__ == '__main__':
    create_cover_letter()
