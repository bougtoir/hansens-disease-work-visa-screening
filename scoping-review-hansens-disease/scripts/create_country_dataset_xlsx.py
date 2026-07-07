#!/usr/bin/env python3
"""Generate S2 Appendix: Comprehensive Country-Level Data as XLSX.

Creates a multi-sheet Excel file with:
  Sheet 1: All 197 countries with category classification
  Sheet 2: Category A detail (Hansen's disease screening, n=20)
  Sheet 3: Category B examples (disease-specific screening without Hansen's, n=15)
  Sheet 4: Multilingual research results (n=5)
  Sheet 5: Unreachable countries (n=82)
  Sheet 6: ILEP cross-reference (S3 equivalent)
"""

import json
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FIGURES_DIR = os.path.join(SCRIPT_DIR, "..", "figures")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Style constants
HEADER_FONT = Font(name='Arial', size=10, bold=True, color='FFFFFF')
HEADER_FILL_A = PatternFill(start_color='C62828', end_color='C62828', fill_type='solid')
HEADER_FILL_B = PatternFill(start_color='1565C0', end_color='1565C0', fill_type='solid')
HEADER_FILL_C = PatternFill(start_color='F9A825', end_color='F9A825', fill_type='solid')
HEADER_FILL_D = PatternFill(start_color='2E7D32', end_color='2E7D32', fill_type='solid')
HEADER_FILL_MAIN = PatternFill(start_color='1A237E', end_color='1A237E', fill_type='solid')
HEADER_FILL_ILEP = PatternFill(start_color='6A1B9A', end_color='6A1B9A', fill_type='solid')
BODY_FONT = Font(name='Arial', size=9)
WRAP = Alignment(wrap_text=True, vertical='top')
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)


def style_header(ws, row, fill):
    for cell in ws[row]:
        cell.font = HEADER_FONT
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        cell.border = THIN_BORDER


def style_body(ws, start_row, end_row):
    for row in ws.iter_rows(min_row=start_row, max_row=end_row):
        for cell in row:
            cell.font = BODY_FONT
            cell.alignment = WRAP
            cell.border = THIN_BORDER


def auto_width(ws, min_width=8, max_width=45):
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        lengths = []
        for cell in col_cells:
            if cell.value:
                lines = str(cell.value).split('\n')
                lengths.append(max(len(l) for l in lines))
        if lengths:
            w = min(max(max(lengths) + 2, min_width), max_width)
            ws.column_dimensions[col_letter].width = w


# ── DATA ──

# Category A: Hansen's disease explicitly named (20 countries)
cat_a = [
    ('Saudi Arabia', 'GCC/Middle East', 'A', 'GAMCA/WAFID Regulations 6th Version',
     'Automatic exclusion', '2021', 'TB, HIV, Hep B/C, syphilis, leprosy, others',
     'Yes', 'Unified GAMCA/WAFID system; examinations at accredited centres in origin countries'),
    ('United Arab Emirates', 'GCC/Middle East', 'A', 'GAMCA/WAFID Regulations',
     'Automatic exclusion', '2021', 'TB, HIV, Hep B/C, syphilis, leprosy, others',
     'Yes', 'GAMCA/WAFID system'),
    ('Qatar', 'GCC/Middle East', 'A', 'GAMCA/WAFID Regulations',
     'Automatic exclusion', '2021', 'TB, HIV, Hep B/C, syphilis, leprosy, others',
     'Yes', 'GAMCA/WAFID system'),
    ('Kuwait', 'GCC/Middle East', 'A', 'GAMCA/WAFID Regulations',
     'Automatic exclusion', '2021', 'TB, HIV, Hep B/C, syphilis, leprosy, others',
     'Yes', 'GAMCA/WAFID system'),
    ('Oman', 'GCC/Middle East', 'A', 'GAMCA/WAFID Regulations',
     'Automatic exclusion', '2021', 'TB, HIV, Hep B/C, syphilis, leprosy, others',
     'Yes', 'GAMCA/WAFID system'),
    ('Bahrain', 'GCC/Middle East', 'A', 'GAMCA/WAFID Regulations',
     'Automatic exclusion', '2021', 'TB, HIV, Hep B/C, syphilis, leprosy, others',
     'Yes', 'GAMCA/WAFID system'),
    ('China', 'East/SE Asia', 'A', 'Foreigner Physical Examination Form',
     'Certification of absence', '\u2014', 'TB, HIV, syphilis, leprosy, others',
     'Yes', 'Standard form for all foreign workers'),
    ('Thailand', 'East/SE Asia', 'A', 'Emergency Decree B.E. 2560',
     'Prohibited disease', '2017', 'TB, leprosy, drug addiction, alcoholism, elephantiasis',
     'Yes', 'Leprosy listed as "prohibited disease"'),
    ('Taiwan', 'East/SE Asia', 'A', 'Work Permit Regulations',
     'Certification of absence', '\u2014', 'TB, HIV, syphilis, leprosy, others',
     'Yes', 'Required for blue-collar work permits'),
    ('Malaysia', 'East/SE Asia', 'A', 'FOMEMA Requirements',
     'Medical unfitness', '\u2014', 'TB, HIV, Hep B, syphilis, leprosy, malaria',
     'Yes', 'Foreign Workers Medical Examination Monitoring Agency'),
    ('Philippines', 'East/SE Asia', 'A', 'OWWA Requirements',
     'Certification of absence', '\u2014', 'TB, HIV, Hep B, syphilis, leprosy',
     'Yes', 'For outbound OFW deployment'),
    ('Hong Kong SAR', 'East/SE Asia', 'A', 'Immigration Ord. Cap. 115',
     'Medical inadmissibility', '\u2014', 'TB, leprosy, others',
     'Yes', 'Listed as inadmissible condition'),
    ('South Africa', 'Africa', 'A', 'Immigration Act 2002; Form BI-811',
     'Certification of absence', '2002', 'TB, HIV, leprosy, others',
     'Yes', 'Medical certificate form BI-811'),
    ('Namibia', 'Africa', 'A', 'Immigration Control Act',
     'Medical prohibition', '1993', 'TB, leprosy, others',
     'Yes', 'Immigration Control Act 1993'),
    ('Russia', 'Europe', 'A', 'Federal Law No. 115-FZ',
     'Deportation ground', '2002', 'TB, HIV, leprosy, syphilis',
     'Yes', 'Grounds for deportation / refusal of work permit'),
    ('Malta', 'Europe', 'A', 'Immigration Act Cap. 217',
     'Medical inadmissibility', '\u2014', 'Leprosy listed',
     'Yes', 'Medical inadmissibility under immigration law'),
    ('United States', 'Americas', 'A', 'INA \u00a7212(a)(1)(A)(i)',
     'Class A condition (waiver available)', '1952', 'TB, syphilis, leprosy, others',
     'Yes', 'CDC Technical Instructions (2018) allow waiver for MDT completion'),
    ('Barbados', 'Americas', 'A', 'Immigration Act Cap. 190',
     'Medical prohibition', '\u2014', 'Leprosy listed',
     'Yes', 'Medical prohibition under immigration law'),
    ('US Virgin Islands', 'Americas', 'A', 'US Federal Law',
     'Class A condition (waiver available)', '\u2014', 'Same as US federal law',
     'Yes', 'Subject to US INA provisions'),
    ('India', 'South Asia', 'A', 'State employment laws',
     'Employment restriction', 'Various', 'Varies by state',
     'Yes', 'Multiple state-level laws; ILEP documents 108 discriminatory laws'),
]

# Category B: Disease-specific screening without Hansen's (selected examples)
cat_b = [
    ('Canada', 'Americas', 'B', 'TB, syphilis', 'Immigration Medical Examination (IME)', 'No'),
    ('Australia', 'Oceania', 'B', 'TB, HIV, Hep B/C', 'Panel physician system (Bupa Medical Visa Services)', 'No'),
    ('United Kingdom', 'Europe', 'B', 'TB (for >6 months visa)', 'IOM panel physician screening', 'No'),
    ('New Zealand', 'Oceania', 'B', 'TB, HIV, Hep B', 'Immigration NZ panel physician', 'No'),
    ('Japan', 'East/SE Asia', 'B', 'TB only (specific nationals)', 'JPETS pre-entry TB screening', 'No'),
    ('South Korea', 'East/SE Asia', 'B', 'TB, HIV, drug use', 'E-9 visa medical examination', 'No'),
    ('Singapore', 'East/SE Asia', 'B', 'TB, HIV, syphilis, Hep B, malaria', 'Foreign worker medical exam (FWME)', 'No'),
    ('Israel', 'Middle East', 'B', 'TB, HIV, Hep B/C', 'Entry medical requirements', 'No'),
    ('Kenya', 'Africa', 'B', 'TB, HIV', 'Work permit medical certificate', 'No'),
    ('Nigeria', 'Africa', 'B', 'TB, HIV, Hep B', 'Expatriate quota medical', 'No'),
    ('Ghana', 'Africa', 'B', 'TB, HIV, yellow fever', 'Immigration medical certificate', 'No'),
    ('Brazil', 'Americas', 'B', 'Yellow fever (vaccination)', 'No work visa medical exam', 'No'),
    ('Mexico', 'Americas', 'B', 'None specified', 'No standard medical exam', 'No'),
    ('Argentina', 'Americas', 'B', 'None specified', 'No disease-specific exam', 'No'),
    ('Chile', 'Americas', 'B', 'TB, HIV (for specific visas)', 'Temporary resident visa medical', 'No'),
]

# Multilingual research results (5 countries)
multilingual = [
    ('Jordan', 'Middle East', 'B (multilingual)', 'TB, HIV, Hep B/C, syphilis', 'Arabic', 'No'),
    ('Lebanon', 'Middle East', 'B (multilingual)', 'VDRL, HIV, cholera, malaria, Hep B', 'Arabic', 'No'),
    ('Vietnam', 'East/SE Asia', 'B (multilingual)', 'HIV, Hep B, syphilis, malaria', 'Vietnamese', 'No'),
    ('Sri Lanka', 'South Asia', 'B (multilingual)', 'TB, HIV, malaria, filariasis', 'Sinhala', 'No'),
    ('Indonesia', 'East/SE Asia', 'B (multilingual)', 'Primarily TB', 'Indonesian', 'No'),
]

# EU/Schengen countries (Category D)
eu_schengen = [
    'Austria', 'Belgium', 'Bulgaria', 'Croatia', 'Cyprus', 'Czech Republic',
    'Denmark', 'Estonia', 'Finland', 'France', 'Germany', 'Greece',
    'Hungary', 'Iceland', 'Ireland', 'Italy', 'Latvia', 'Liechtenstein',
    'Lithuania', 'Luxembourg', 'Netherlands', 'Norway', 'Poland', 'Portugal',
    'Romania', 'Slovakia', 'Slovenia', 'Spain', 'Sweden', 'Switzerland',
]

# Other Category D countries (non-EU, no disease-specific screening)
other_cat_d = [
    'Vatican City',
]

# ILEP cross-reference data
ilep_data = [
    ('India', 108, 'Domestic employment, property rights, marriage, civic participation',
     'Yes (state employment laws)', 'Multiple state-level provisions; largest number of discriminatory laws globally'),
    ('Japan', 3, 'Former isolation laws (repealed 1996)',
     'No (abolished)', 'Leprosy Prevention Law abolished 1996; Kumamoto District Court ruling 2001'),
    ('China', 2, 'Employment restrictions, immigration',
     'Yes (Foreigner Physical Examination Form)', 'Immigration screening still active'),
    ('Brazil', 2, 'Health surveillance provisions',
     'No', 'No immigration screening for Hansen\'s disease'),
    ('Nepal', 4, 'Marriage, employment, civic participation',
     'Not confirmed', 'Not in Category A; requirements not publicly available'),
    ('Ethiopia', 3, 'Employment, family law',
     'Not confirmed', 'Unreachable country; requirements not publicly available'),
    ('Madagascar', 2, 'Health law provisions',
     'Not confirmed', 'Unreachable country'),
    ('South Africa', 1, 'Immigration law',
     'Yes (Immigration Act 2002; BI-811)', 'Confirmed in our review'),
    ('United States', 1, 'Immigration law (INA)',
     'Yes (INA \u00a7212(a)(1)(A)(i))', 'Class A condition; waiver available since 2018'),
    ('Russia', 1, 'Immigration / residency law',
     'Yes (Federal Law No. 115-FZ)', 'Deportation ground for foreign workers'),
    ('Thailand', 1, 'Immigration / public health',
     'Yes (Emergency Decree B.E. 2560)', 'Prohibited disease designation'),
    ('Philippines', 1, 'Employment / overseas worker',
     'Yes (OWWA Requirements)', 'Overseas Filipino Worker deployment screening'),
]

# Additional countries identified in our review but NOT in ILEP database
additional_not_in_ilep = [
    ('Malta', 'Immigration Act Cap. 217', 'Medical inadmissibility'),
    ('Barbados', 'Immigration Act Cap. 190', 'Medical prohibition'),
    ('US Virgin Islands', 'US Federal Law', 'Class A condition'),
    ('Taiwan', 'Work Permit Regulations', 'Certification of absence'),
    ('Malaysia', 'FOMEMA Requirements', 'Medical unfitness'),
    ('Hong Kong SAR', 'Immigration Ord. Cap. 115', 'Medical inadmissibility'),
    ('Namibia', 'Immigration Control Act', 'Medical prohibition'),
    ('Saudi Arabia', 'GAMCA/WAFID Regulations', 'Automatic exclusion'),
    ('UAE', 'GAMCA/WAFID Regulations', 'Automatic exclusion'),
    ('Qatar', 'GAMCA/WAFID Regulations', 'Automatic exclusion'),
    ('Kuwait', 'GAMCA/WAFID Regulations', 'Automatic exclusion'),
    ('Oman', 'GAMCA/WAFID Regulations', 'Automatic exclusion'),
    ('Bahrain', 'GAMCA/WAFID Regulations', 'Automatic exclusion'),
]


def create_dataset():
    wb = Workbook()

    # ── Sheet 1: All 197 Countries Overview ──
    ws1 = wb.active
    ws1.title = 'All 197 Countries'
    ws1.append([
        'Country/Territory', 'Region', 'Category', 'Category Description',
        'Hansen\'s Disease Screening', 'Data Source Language', 'Notes'
    ])
    style_header(ws1, 1, HEADER_FILL_MAIN)

    row_num = 2
    # Category A
    for c in cat_a:
        ws1.append([c[0], c[1], 'A', 'Hansen\'s disease explicitly named', 'Yes', 'English', c[8]])
        row_num += 1

    # Category B (selected)
    for c in cat_b:
        ws1.append([c[0], c[1], 'B', 'Disease-specific screening, Hansen\'s NOT named', 'No', 'English', c[4]])
        row_num += 1

    # Multilingual (Category B via multilingual)
    for c in multilingual:
        ws1.append([c[0], c[1], 'B', 'Disease-specific screening, Hansen\'s NOT named', 'No', c[4], 'Identified via multilingual research'])
        row_num += 1

    # Category C (unreachable)
    cls_path = os.path.join(FIGURES_DIR, 'country_classification.json')
    with open(cls_path, 'r') as f:
        cls = json.load(f)
    for region, countries in cls['unreachable_by_region'].items():
        for country in countries:
            lang = cls['unreachable_languages'].get(country, '')
            ws1.append([country, region, 'C', 'Medical exam required; specific diseases not confirmed',
                       'Not confirmed', lang, 'Information not publicly available'])
            row_num += 1

    # Category D (EU/Schengen + others)
    for c in eu_schengen:
        ws1.append([c, 'Europe (EU/Schengen)', 'D', 'No disease-specific medical exam required',
                    'No', 'English/official EU', 'EU free movement of workers (TFEU)'])
        row_num += 1
    for c in other_cat_d:
        ws1.append([c, 'Europe', 'D', 'No disease-specific medical exam required',
                    'No', '\u2014', ''])
        row_num += 1

    style_body(ws1, 2, row_num)
    auto_width(ws1)

    # ── Sheet 2: Category A Detail ──
    ws2 = wb.create_sheet('Cat A \u2013 Hansen\'s Named (n=20)')
    ws2.append([
        'Country/Territory', 'Region', 'Legal Instrument', 'Provision Type',
        'Year', 'Other Diseases Screened', 'Hansen\'s Disease', 'Notes'
    ])
    style_header(ws2, 1, HEADER_FILL_A)
    for i, c in enumerate(cat_a, 2):
        ws2.append([c[0], c[1], c[3], c[4], c[5], c[6], c[7], c[8]])
    style_body(ws2, 2, len(cat_a) + 1)
    auto_width(ws2)

    # ── Sheet 3: Category B Examples ──
    ws3 = wb.create_sheet('Cat B \u2013 No Hansen\'s (n=15)')
    ws3.append([
        'Country', 'Region', 'Category', 'Diseases Screened',
        'Notable Features', 'Hansen\'s Disease'
    ])
    style_header(ws3, 1, HEADER_FILL_B)
    for i, c in enumerate(cat_b, 2):
        ws3.append(list(c))
    style_body(ws3, 2, len(cat_b) + 1)
    auto_width(ws3)

    # ── Sheet 4: Multilingual Research ──
    ws4 = wb.create_sheet('Multilingual Research (n=5)')
    ws4.append([
        'Country', 'Region', 'Category', 'Diseases Screened',
        'Research Language', 'Hansen\'s Disease'
    ])
    style_header(ws4, 1, HEADER_FILL_C)
    for c in multilingual:
        ws4.append(list(c))
    style_body(ws4, 2, len(multilingual) + 1)
    auto_width(ws4)

    # ── Sheet 5: Unreachable Countries ──
    ws5 = wb.create_sheet('Unreachable (n=82)')
    ws5.append(['Region', 'Country', 'Official Language(s)'])
    style_header(ws5, 1, PatternFill(start_color='757575', end_color='757575', fill_type='solid'))
    row_n = 2
    for region, countries in cls['unreachable_by_region'].items():
        for idx, country in enumerate(countries):
            lang = cls['unreachable_languages'].get(country, '')
            ws5.append([region if idx == 0 else '', country, lang])
            row_n += 1
    style_body(ws5, 2, row_n - 1)
    auto_width(ws5)

    # ── Sheet 6: ILEP Cross-Reference (S3) ──
    ws6 = wb.create_sheet('ILEP Cross-Reference (S3)')
    ws6.append([
        'Country', 'ILEP Laws Count', 'ILEP Law Categories',
        'Immigration Screening (This Review)', 'Notes'
    ])
    style_header(ws6, 1, HEADER_FILL_ILEP)
    for c in ilep_data:
        ws6.append(list(c))
    style_body(ws6, 2, len(ilep_data) + 1)

    # Add section for countries in our review but not in ILEP
    r = len(ilep_data) + 3
    ws6.cell(row=r, column=1, value='Countries identified in this review but NOT in ILEP database')
    ws6.cell(row=r, column=1).font = Font(name='Arial', size=10, bold=True)
    r += 1
    ws6.cell(row=r, column=1, value='Country/Territory')
    ws6.cell(row=r, column=2, value='Legal Instrument')
    ws6.cell(row=r, column=3, value='Provision Type')
    for cell in ws6[r]:
        if cell.value:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL_ILEP
            cell.border = THIN_BORDER
    r += 1
    for c in additional_not_in_ilep:
        ws6.cell(row=r, column=1, value=c[0])
        ws6.cell(row=r, column=2, value=c[1])
        ws6.cell(row=r, column=3, value=c[2])
        for col in range(1, 4):
            cell = ws6.cell(row=r, column=col)
            cell.font = BODY_FONT
            cell.alignment = WRAP
            cell.border = THIN_BORDER
        r += 1

    # Summary row
    r += 1
    ws6.cell(row=r, column=1,
             value='Summary: ILEP documents 139 discriminatory laws across 24 countries. '
                   '9 laws in 8 countries pertain to immigration/citizenship. '
                   'India accounts for 108/139 (77.7%). '
                   'This review identified 13 additional countries/territories not in the ILEP database.')
    ws6.cell(row=r, column=1).font = Font(name='Arial', size=9, italic=True)
    ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)

    auto_width(ws6)

    output_path = os.path.join(OUTPUT_DIR, "PGPH_S2_S3_Country_Data.xlsx")
    wb.save(output_path)
    print(f"Saved: {output_path}")
    return output_path


if __name__ == '__main__':
    create_dataset()
