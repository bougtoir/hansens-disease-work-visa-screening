#!/usr/bin/env python3
"""Generate PRISMA-ScR Checklist (filled) for PGPH submission as DOCX.

Uses the official PRISMA-ScR template structure and fills in the
'Reported on Page #' column based on the PGPH manuscript content.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def set_cell_shading(cell, color_hex):
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color_hex)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)




def create_prisma_scr_checklist():
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        section.page_width = Inches(11)
        section.page_height = Inches(8.5)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(9)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(
        'Preferred Reporting Items for Systematic reviews and Meta-Analyses '
        'extension for Scoping Reviews (PRISMA-ScR) Checklist'
    )
    run.bold = True
    run.font.size = Pt(11)
    run.font.name = 'Arial'

    doc.add_paragraph()

    # Create the checklist table
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    # Set column widths
    for row in table.rows:
        row.cells[0].width = Cm(3.5)
        row.cells[1].width = Cm(1.2)
        row.cells[2].width = Cm(14.0)
        row.cells[3].width = Cm(5.0)

    # Header row
    headers = ['SECTION', 'ITEM', 'PRISMA-ScR CHECKLIST ITEM', 'REPORTED ON PAGE #']
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.name = 'Arial'
        set_cell_shading(cell, '1A237E')

    # Section header style
    section_bg = 'D0D0E8'

    # Checklist items: (section_label, item_no, description, reported_on)
    # Section headers have item_no = None
    items = [
        # TITLE
        ('TITLE', None, None, None),
        ('Title', '1',
         'Identify the report as a scoping review.',
         'p. 1 (Title page: "...Scoping Review of Administrative Requirements...")'),

        # ABSTRACT
        ('ABSTRACT', None, None, None),
        ('Structured summary', '2',
         'Provide a structured summary that includes (as applicable): background, '
         'objectives, eligibility criteria, sources of evidence, charting methods, '
         'results, and conclusions that relate to the review questions and objectives.',
         'pp. 2\u20133 (Abstract with Background, Objectives, Methods, Results, Conclusions subheadings)'),

        # INTRODUCTION
        ('INTRODUCTION', None, None, None),
        ('Rationale', '3',
         'Describe the rationale for the review in the context of what is already '
         'known. Explain why the review questions/objectives lend themselves to a '
         'scoping review approach.',
         'pp. 4\u20135 (Section 1. Introduction, paragraphs 1\u20135)'),
        ('Objectives', '4',
         'Provide an explicit statement of the questions and objectives being addressed '
         'with reference to their key elements (e.g., population or participants, concepts, '
         'and context) or other relevant key elements used to conceptualize the review '
         'questions and/or objectives.',
         'pp. 5\u20136 (Section 1.1 Objectives: five explicit research questions with PCC elements)'),

        # METHODS
        ('METHODS', None, None, None),
        ('Protocol and registration', '5',
         'Indicate whether a review protocol exists; state if and where it can be accessed '
         '(e.g., a Web address); and if available, provide registration information, '
         'including the registration number.',
         'p. 6 (Section 2.1: "conducted in accordance with PRISMA-ScR... protocol was developed a priori")'),
        ('Eligibility criteria', '6',
         'Specify characteristics of the sources of evidence used as eligibility criteria '
         '(e.g., years considered, language, and publication status), and provide a rationale.',
         'pp. 6\u20137 (Section 2.2: Population, Concept, Context, and Types of evidence sources defined)'),
        ('Information sources*', '7',
         'Describe all information sources in the search (e.g., databases with dates of '
         'coverage and contact with authors to identify additional sources), as well as the '
         'date the most recent search was executed.',
         'pp. 7\u20138 (Section 2.3: seven information sources listed; search conducted January\u2013March 2026)'),
        ('Search', '8',
         'Present the full electronic search strategy for at least 1 database, including '
         'any limits used, such that it could be repeated.',
         'p. 8 (Section 2.3: search terms and Boolean combinations presented)'),
        ('Selection of sources of evidence\u2020', '9',
         'State the process for selecting sources of evidence (i.e., screening and eligibility) '
         'included in the scoping review.',
         'p. 8 (Section 2.4: two-stage selection by geographic region; criteria (a)\u2013(c) described)'),
        ('Data charting process\u2021', '10',
         'Describe the methods of charting data from the included sources of evidence (e.g., '
         'calibrated forms or forms that have been tested by the team before their use, and '
         'whether data charting was done independently or in duplicate) and any processes for '
         'obtaining and confirming data from investigators.',
         'pp. 8\u20139 (Section 2.5: standardized charting form with extracted variables listed)'),
        ('Data items', '11',
         'List and define all variables for which data were sought and any assumptions and '
         'simplifications made.',
         'pp. 8\u20139 (Section 2.5: country name, UN status, region, visa types, diseases, provision type, legal source, year)'),
        ('Critical appraisal of individual sources of evidence\u00a7', '12',
         'If done, provide a rationale for conducting a critical appraisal of included sources '
         'of evidence; describe the methods used and how this information was used in any data '
         'synthesis (if appropriate).',
         'Not applicable \u2014 critical appraisal of individual sources was not conducted, consistent with scoping review methodology [Arksey & O\u2019Malley 2005]'),
        ('Synthesis of results', '13',
         'Describe the methods of handling and summarizing the data that were charted.',
         'pp. 8\u20139 (Section 2.6: descriptive analytical approach; four-category classification; frequency analysis; geographic mapping)'),

        # RESULTS
        ('RESULTS', None, None, None),
        ('Selection of sources of evidence', '14',
         'Give numbers of sources of evidence screened, assessed for eligibility, and included '
         'in the review, with reasons for exclusions at each stage, ideally using a flow diagram.',
         'pp. 9\u201312 (Section 3.1 + Figure 1 PRISMA-ScR flow diagram + Figure 2 Sankey diagram)'),
        ('Characteristics of sources of evidence', '15',
         'For each source of evidence, present characteristics for which data were charted and '
         'provide the citations.',
         'pp. 13\u201315 (Section 3.2\u20133.3; Table 1: 20 countries with Hansen\u2019s disease screening; Table 2: 15 countries without)'),
        ('Critical appraisal within sources of evidence', '16',
         'If done, present data on critical appraisal of included sources of evidence (see item 12).',
         'Not applicable \u2014 see Item 12'),
        ('Results of individual sources of evidence', '17',
         'For each included source of evidence, present the relevant data that were charted '
         'that relate to the review questions and objectives.',
         'pp. 14\u201320 (Table 1: individual country data with region, legal instrument, provision type, year; Table 2: selected comparator countries)'),
        ('Synthesis of results', '18',
         'Summarize and/or present the charting results as they relate to the review questions '
         'and objectives.',
         'pp. 16\u201321 (Sections 3.4\u20133.8; Figures 4\u20137: GCC/GAMCA analysis, disease frequency, regional distribution, transmissibility-screening discordance)'),

        # DISCUSSION
        ('DISCUSSION', None, None, None),
        ('Summary of evidence', '19',
         'Summarize the main results (including an overview of concepts, themes, and types of '
         'evidence available), link to the review questions and objectives, and consider the '
         'relevance to key groups.',
         'pp. 22\u201324 (Sections 4.1\u20134.4: evidence summary, administrative divergence from medical practice, disease comparison, Japan comparative model)'),
        ('Limitations', '20',
         'Discuss the limitations of the scoping review process.',
         'pp. 24\u201325 (Section 4.6: five limitations discussed including 82 unreachable countries, non-public requirements, temporal snapshot, enforcement vs. law, Category D confirmation)'),
        ('Conclusions', '21',
         'Provide a general interpretation of the results with respect to the review questions '
         'and objectives, as well as potential implications and/or next steps.',
         'pp. 26\u201327 (Section 5: four specific policy recommendations; implications for GCC, WHO, UN HRC; Japan as model)'),

        # FUNDING
        ('FUNDING', None, None, None),
        ('Funding', '22',
         'Describe sources of funding for the included sources of evidence, as well as sources '
         'of funding for the scoping review. Describe the role of the funders of the scoping review.',
         'p. 35 (Declarations section)'),
    ]

    for section_label, item_no, description, reported_on in items:
        row = table.add_row()
        if item_no is None:
            # Section header row — merge all cells into one
            merged_cell = row.cells[0].merge(row.cells[3])
            set_cell_shading(merged_cell, section_bg)
            merged_cell.text = section_label
            for paragraph in merged_cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(8)
                    run.font.name = 'Arial'
        else:
            row.cells[0].text = section_label
            row.cells[1].text = item_no
            row.cells[2].text = description
            row.cells[3].text = reported_on

            row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(8)
                        run.font.name = 'Arial'

    # Footnotes
    doc.add_paragraph()
    footnotes = [
        'JBI = Joanna Briggs Institute.',
        '* Where sources of evidence (see: https://jbi.global/scoping-review-network/resources).',
        '\u2020 Include the process of managing and selecting sources of evidence.',
        '\u2021 Describe the methods for charting/extracting data from included sources.',
        '\u00a7 This is an optional item.',
    ]
    for fn in footnotes:
        p = doc.add_paragraph()
        run = p.add_run(fn)
        run.font.size = Pt(7)
        run.font.name = 'Arial'
        run.italic = True

    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run(
        'From: Tricco AC, Lillie E, Zarin W, et al. PRISMA Extension for Scoping Reviews '
        '(PRISMA-ScR): Checklist and Explanation. Ann Intern Med. 2018;169:467\u201373. '
        'doi: 10.7326/M18-0850.'
    )
    run.font.size = Pt(7)
    run.font.name = 'Arial'
    run.italic = True

    # Save
    output_path = os.path.join(OUTPUT_DIR, "PGPH_PRISMA_ScR_Checklist.docx")
    doc.save(output_path)
    print(f"Saved: {output_path}")
    return output_path


if __name__ == '__main__':
    create_prisma_scr_checklist()
