# Scoping Review: Hansen's Disease Screening in Work Visa Medical Examinations

A PRISMA-ScR compliant scoping review analyzing administrative requirements for Hansen's disease (leprosy) screening in work visa medical examinations across 197 countries and territories.

## Key Findings

- **20 countries (10.2%)** explicitly name Hansen's disease in work visa medical requirements
- **110 countries (55.8%)** characterized via English-language sources
- **5 countries (2.5%)** additionally confirmed via multilingual research
- **82 countries (41.6%)** remained unreachable despite multilingual research
- Hansen's disease ranked as the **5th most commonly screened disease** (34.5%) despite low transmissibility

## Repository Structure

```
scoping-review-hansens-disease/
├── scripts/           # Python scripts to generate all figures and documents
│   ├── create_figures.py          # Generate matplotlib figures (bar, donut, scatter, etc.)
│   ├── create_sankey.py           # Sankey diagram for data accessibility flow
│   ├── create_accessibility_map.py # World map of data accessibility
│   ├── create_plos_en.py          # PLoS NTDs full paper (English DOCX)
│   ├── create_plos_ja.py          # PLoS NTDs full paper (Japanese DOCX)
│   ├── create_lancet_en.py        # Lancet Global Health Comment (English DOCX)
│   ├── create_lancet_ja.py        # Lancet Global Health Comment (Japanese DOCX)
│   ├── create_pptx_en.py          # PowerPoint figures (English)
│   ├── create_pptx_ja.py          # PowerPoint figures (Japanese)
│   ├── update_fig1.py             # World map with base layer
│   └── language_research.py       # Multilingual research data
├── figures/           # Generated figures (PNG) and data (JSON)
│   ├── fig1_world_map.png         # Global distribution of Hansen's disease screening
│   ├── fig2_disease_bar.png       # Disease frequency bar chart
│   ├── fig3_regional_donut.png    # Regional distribution donut chart
│   ├── fig4_prisma_flow.png       # PRISMA-ScR flow diagram
│   ├── fig5_transmissibility.png  # Transmissibility vs screening frequency
│   ├── fig6_sankey_en.png         # Data accessibility Sankey (English)
│   ├── fig6_sankey_ja.png         # Data accessibility Sankey (Japanese)
│   ├── fig7_accessibility_en.png  # Data accessibility world map (English)
│   ├── fig7_accessibility_ja.png  # Data accessibility world map (Japanese)
│   └── country_classification.json # Structured country data (197 countries)
├── output/            # Generated documents (DOCX, PPTX)
│   ├── PLoS_NTDs_Full_Paper_EN.docx
│   ├── PLoS_NTDs_Full_Paper_JA.docx
│   ├── Lancet_Global_Health_Comment_EN.docx
│   ├── Lancet_Global_Health_Comment_JA.docx
│   ├── Scoping_Review_Figures_EN.pptx
│   └── Scoping_Review_Figures_JA.pptx
└── docs/              # Reference documents and reports
    ├── scoping_review_report.md   # Full scoping review report (English)
    ├── summary_ja.md              # Summary (Japanese)
    └── target_journals.md         # Target journal analysis
```

## Target Journals

1. **PLoS Neglected Tropical Diseases** (IF 3.4) - Full paper with PRISMA-ScR
2. **The Lancet Global Health** (IF 34.3) - Comment/Viewpoint short report

## Data Categories

| Category | Description | Count |
|----------|-------------|-------|
| A | Hansen's disease explicitly named | 20 |
| B | Disease-specific screening, no Hansen's | 38 |
| C | Medical exam required, diseases unconfirmed | 87 |
| D | No disease-specific medical exam requirement | 52 |
| **Total** | | **197** |

## Data Accessibility

- **English sources:** 110 countries (Cat A + B + D = 20 + 38 + 52)
- **Multilingual research:** 5 additional countries (Jordan, Lebanon, Vietnam, Sri Lanka, Indonesia)
- **Unreachable:** 82 countries (Cat C remainder after multilingual confirmation)

## Requirements

- Python 3.8+
- matplotlib, geopandas, python-docx, python-pptx
- Noto Sans CJK JP font (for Japanese figures)

## License

[To be determined]
