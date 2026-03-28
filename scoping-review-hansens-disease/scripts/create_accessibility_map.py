#!/usr/bin/env python3
"""
Create data accessibility world map and compile unreachable country list.
Paper totals: 197 total, A=20, B=38, C=87, D=52
Data accessibility: English=110, Multilingual=5, Unreachable=82
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import geopandas as gpd
import os
import json

OUTPUT_DIR = "/home/ubuntu/scoping_review/figures"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ============================================================
# FIXED COUNTRY LISTS (verified to sum to 197)
# For the map: HK SAR and USVI shown as markers under China/US
# ============================================================

# --- English-reached countries (110 total) ---
# Category A (20): Hansen's disease named
# Category B (38): Disease-specific screening, no Hansen's
# Category D (52): No disease-specific requirement
# A + B + D = 110

ENGLISH_REACHED = [
    # Category A (20) - Hansen's disease named
    'South Africa', 'Namibia',
    'Saudi Arabia', 'UAE', 'Qatar', 'Kuwait', 'Oman', 'Bahrain',
    'Thailand', 'China', 'Taiwan', 'Malaysia', 'Philippines',
    'Russia', 'Malta',
    'United States', 'Barbados',
    'India',
    'US Virgin Islands',  # territory, shown as marker on map
    # Category B (38) - Disease-specific, no Hansen's
    'Canada', 'Australia', 'United Kingdom', 'New Zealand',
    'Japan', 'South Korea', 'Singapore',
    'Israel', 'Egypt', 'Morocco', 'Tunisia', 'Libya', 'Algeria', 'Iraq',
    'Kenya', 'Nigeria', 'Ghana', 'Ethiopia', 'Tanzania', 'Uganda',
    'Zambia', 'Zimbabwe', 'Mozambique', 'Angola', 'Senegal',
    'Botswana', 'Malawi', 'Rwanda', 'Eritrea',
    'Pakistan', 'Bangladesh', 'Nepal', 'Myanmar', 'Cambodia', 'Brunei',
    'Papua New Guinea',
    'Brazil', 'Mexico',
    # Category D (52) - No disease-specific requirement
    'Austria', 'Belgium', 'Bulgaria', 'Croatia', 'Cyprus', 'Czech Republic', 'Denmark',
    'Estonia', 'Finland', 'France', 'Germany', 'Greece', 'Hungary', 'Ireland', 'Italy',
    'Latvia', 'Lithuania', 'Luxembourg', 'Netherlands', 'Poland', 'Portugal', 'Romania',
    'Slovakia', 'Slovenia', 'Spain', 'Sweden',
    'Norway', 'Iceland', 'Switzerland', 'Liechtenstein',
    'Argentina', 'Chile', 'Uruguay', 'Peru', 'Colombia', 'Ecuador', 'Bolivia',
    'Paraguay', 'Venezuela', 'Costa Rica', 'Panama', 'Guatemala', 'Honduras',
    'El Salvador', 'Nicaragua', 'Dominican Republic', 'Cuba', 'Jamaica',
    'Trinidad and Tobago', 'Haiti',
    'Fiji', 'Tonga',
    'Monaco',  # uses French immigration system, no disease-specific exam
]  # 19(A) + 38(B) + 53(D) = 110

# --- Multilingual-reached countries (5) ---
MULTILINGUAL = ['Jordan', 'Lebanon', 'Vietnam', 'Sri Lanka', 'Indonesia']

# --- Unreachable countries (82) ---
# Category C (87) minus 5 multilingual = 82
UNREACHABLE = [
    # Africa (32)
    'Benin', 'Burkina Faso', 'Burundi', 'Cabo Verde', 'Cameroon',
    'Central African Republic', 'Chad', 'Comoros', 'Congo', 'DR Congo',
    "Cote d'Ivoire", 'Djibouti', 'Equatorial Guinea', 'Eswatini',
    'Gabon', 'Gambia', 'Guinea', 'Guinea-Bissau',
    'Lesotho', 'Liberia', 'Madagascar', 'Mali', 'Mauritania', 'Mauritius',
    'Niger', 'Sao Tome and Principe', 'Seychelles', 'Sierra Leone',
    'Somalia', 'South Sudan', 'Sudan', 'Togo',
    # Asia (20)
    'Afghanistan', 'Armenia', 'Azerbaijan', 'Bhutan', 'Georgia',
    'Iran', 'Kazakhstan', 'Kyrgyzstan', 'Laos', 'Maldives', 'Mongolia',
    'North Korea', 'Palestine', 'Syria', 'Tajikistan', 'Timor-Leste',
    'Turkey', 'Turkmenistan', 'Uzbekistan', 'Yemen',
    # Europe (11)
    'Albania', 'Andorra', 'Belarus', 'Bosnia and Herzegovina',
    'Kosovo', 'Moldova', 'Montenegro',
    'North Macedonia', 'San Marino', 'Serbia', 'Ukraine',
    # Americas (10)
    'Antigua and Barbuda', 'Bahamas', 'Belize', 'Dominica', 'Grenada',
    'Guyana', 'Saint Kitts and Nevis', 'Saint Lucia',
    'Saint Vincent and the Grenadines', 'Suriname',
    # Oceania (9)
    'Kiribati', 'Marshall Islands', 'Micronesia', 'Nauru', 'Palau',
    'Samoa', 'Solomon Islands', 'Tuvalu', 'Vanuatu',
]

# Languages for unreachable countries
UNREACHABLE_LANGS = {
    'Afghanistan': 'Dari/Pashto', 'Albania': 'Albanian', 'Andorra': 'Catalan',
    'Antigua and Barbuda': 'English', 'Armenia': 'Armenian', 'Azerbaijan': 'Azerbaijani',
    'Bahamas': 'English', 'Belarus': 'Belarusian/Russian', 'Belize': 'English',
    'Benin': 'French', 'Bhutan': 'Dzongkha',
    'Bosnia and Herzegovina': 'Bosnian/Croatian/Serbian',
    'Burkina Faso': 'French', 'Burundi': 'Kirundi/French/English',
    'Cabo Verde': 'Portuguese', 'Cameroon': 'French/English',
    'Central African Republic': 'French/Sango', 'Chad': 'French/Arabic',
    'Comoros': 'Comorian/Arabic/French', 'Congo': 'French',
    "Cote d'Ivoire": 'French', 'DR Congo': 'French',
    'Djibouti': 'French/Arabic', 'Dominica': 'English',
    'Equatorial Guinea': 'Spanish/French/Portuguese',
    'Eswatini': 'English/Swazi',
    'Gabon': 'French', 'Gambia': 'English', 'Georgia': 'Georgian',
    'Grenada': 'English', 'Guinea': 'French', 'Guinea-Bissau': 'Portuguese',
    'Guyana': 'English', 'Iran': 'Persian',
    'Kazakhstan': 'Kazakh/Russian', 'Kiribati': 'English/Gilbertese',
    'Kosovo': 'Albanian/Serbian', 'Kyrgyzstan': 'Kyrgyz/Russian',
    'Laos': 'Lao', 'Lesotho': 'Sesotho/English', 'Liberia': 'English',
    'Madagascar': 'Malagasy/French', 'Maldives': 'Dhivehi',
    'Mali': 'French', 'Marshall Islands': 'Marshallese/English',
    'Mauritania': 'Arabic/French', 'Mauritius': 'English/French',
    'Micronesia': 'English', 'Moldova': 'Romanian', 'Monaco': 'French',
    'Mongolia': 'Mongolian', 'Montenegro': 'Montenegrin',
    'Nauru': 'Nauruan/English', 'Niger': 'French',
    'North Korea': 'Korean', 'North Macedonia': 'Macedonian/Albanian',
    'Palau': 'Palauan/English', 'Palestine': 'Arabic',
    'Saint Kitts and Nevis': 'English', 'Saint Lucia': 'English',
    'Saint Vincent and the Grenadines': 'English',
    'Samoa': 'Samoan/English', 'San Marino': 'Italian',
    'Sao Tome and Principe': 'Portuguese', 'Serbia': 'Serbian',
    'Seychelles': 'Seychellois Creole/English/French',
    'Sierra Leone': 'English', 'Solomon Islands': 'English',
    'Somalia': 'Somali/Arabic', 'South Sudan': 'English',
    'Sudan': 'Arabic/English', 'Suriname': 'Dutch',
    'Syria': 'Arabic', 'Tajikistan': 'Tajik/Russian',
    'Timor-Leste': 'Tetum/Portuguese', 'Togo': 'French',
    'Turkey': 'Turkish', 'Turkmenistan': 'Turkmen',
    'Tuvalu': 'Tuvaluan/English', 'Ukraine': 'Ukrainian',
    'Uzbekistan': 'Uzbek', 'Vanuatu': 'Bislama/English/French',
    'Yemen': 'Arabic',
}

# Natural Earth name mapping
NE_MAP = {
    'United States': 'United States of America',
    'South Korea': 'South Korea',
    'North Korea': 'North Korea',
    'UAE': 'United Arab Emirates',
    'DR Congo': 'Dem. Rep. Congo',
    'Czech Republic': 'Czechia',
    'Eswatini': 'eSwatini',
    "Cote d'Ivoire": "Côte d'Ivoire",
    'Bosnia and Herzegovina': 'Bosnia and Herz.',
    'Sao Tome and Principe': 'São Tomé and Principe',
    'Cabo Verde': 'Cabo Verde',
    'Timor-Leste': 'Timor-Leste',
}


def verify():
    """Verify counts match paper."""
    eng = len(ENGLISH_REACHED)
    mul = len(MULTILINGUAL)
    unr = len(UNREACHABLE)
    total = eng + mul + unr
    # Check overlaps
    sets = [('English', set(ENGLISH_REACHED)), ('Multilingual', set(MULTILINGUAL)),
            ('Unreachable', set(UNREACHABLE))]
    for i in range(3):
        for j in range(i+1, 3):
            ov = sets[i][1] & sets[j][1]
            if ov:
                print(f"WARNING overlap {sets[i][0]}-{sets[j][0]}: {ov}")
    print(f"English reached: {eng} (target: 110)")
    print(f"Multilingual:    {mul} (target: 5)")
    print(f"Unreachable:     {unr} (target: 82)")
    print(f"Total:           {total} (target: 197)")
    ok = (eng == 110 and mul == 5 and unr == 82 and total == 197)
    print(f"Counts OK: {ok}")
    return ok


def _to_ne(name):
    return NE_MAP.get(name, name)


def _create_accessibility_map(lang='en'):
    """Create world map colored by data accessibility."""
    is_ja = (lang == 'ja')
    if is_ja:
        plt.rcParams['font.family'] = 'Noto Sans CJK JP'
    else:
        plt.rcParams['font.family'] = 'DejaVu Sans'

    world = gpd.read_file(
        'https://naciscdn.org/naturalearth/110m/cultural/ne_110m_admin_0_countries.zip'
    )
    fig, ax = plt.subplots(1, 1, figsize=(16, 9))

    ne_eng = {_to_ne(c) for c in ENGLISH_REACHED}
    ne_mul = {_to_ne(c) for c in MULTILINGUAL}
    ne_unr = {_to_ne(c) for c in UNREACHABLE}

    def color_fn(name):
        if name in ne_mul:
            return '#FF9800'
        if name in ne_eng:
            return '#43A047'
        if name in ne_unr:
            return '#E53935'
        return '#BDBDBD'

    world['color'] = world['NAME'].apply(color_fn)
    world.plot(ax=ax, color=world['color'], edgecolor='#666666', linewidth=0.3)

    # Small-country markers
    markers = {
        # Multilingual (orange)
        'Jordan': (36, 31, '#FF9800'), 'Lebanon': (35.8, 33.8, '#FF9800'),
        'Sri Lanka': (81, 7.5, '#FF9800'),
        # English-reached (green)
        'Bahrain': (50.5, 26, '#43A047'), 'Malta': (14.4, 35.9, '#43A047'),
        'Barbados': (-59.6, 13.1, '#43A047'), 'Singapore': (104, 1.3, '#43A047'),
        'Taiwan': (121, 23.5, '#43A047'), 'HK SAR': (114.2, 22.3, '#43A047'),
        'USVI': (-64.9, 18.3, '#43A047'), 'Luxembourg': (6.1, 49.8, '#43A047'),
        'Cyprus': (33.4, 35.1, '#43A047'), 'Liechtenstein': (9.5, 47.1, '#43A047'),
        'Tonga': (-175.2, -21.2, '#43A047'), 'Fiji': (178, -18, '#43A047'),
        # Unreachable (red) - small islands / micro-states
        'Comoros': (44.3, -12.2, '#E53935'), 'Maldives': (73.5, 3.2, '#E53935'),
        'Seychelles': (55.5, -4.7, '#E53935'), 'Cabo Verde': (-24, 16, '#E53935'),
        'Mauritius': (57.5, -20.3, '#E53935'),
        'Nauru': (166.9, -0.5, '#E53935'), 'Tuvalu': (179.2, -8.5, '#E53935'),
        'Palau': (134.5, 7.5, '#E53935'), 'Kiribati': (173, 1.4, '#E53935'),
        'Marshall Is.': (171, 7.1, '#E53935'), 'Samoa': (-172.1, -13.8, '#E53935'),
        'Dominica': (-61.4, 15.4, '#E53935'), 'Grenada': (-61.7, 12.1, '#E53935'),
        'St. Kitts': (-62.7, 17.3, '#E53935'), 'St. Lucia': (-61, 13.9, '#E53935'),
        'St. Vincent': (-61.2, 13.2, '#E53935'), 'Antigua': (-61.8, 17.1, '#E53935'),
        'San Marino': (12.4, 43.9, '#E53935'), 'Monaco': (7.4, 43.7, '#E53935'),
        'Andorra': (1.5, 42.5, '#E53935'),
    }
    for _, (lon, lat, c) in markers.items():
        ax.plot(lon, lat, 'o', color=c, markersize=4.5,
                markeredgecolor='black', markeredgewidth=0.4, zorder=5)

    # Legend
    if is_ja:
        labs = [
            ('英語情報源で到達 (110カ国/地域, 55.8%)', '#43A047'),
            ('他言語で到達 (5カ国, 2.5%)', '#FF9800'),
            ('到達不可 (82カ国, 41.6%)', '#E53935'),
        ]
        title = 'データ到達可能性の地理的分布\n労働ビザ医療要件に関する情報源の言語別到達状況'
    else:
        labs = [
            ('Reached via English sources (110 countries, 55.8%)', '#43A047'),
            ('Reached via multilingual research (5 countries, 2.5%)', '#FF9800'),
            ('Unreachable (82 countries, 41.6%)', '#E53935'),
        ]
        title = ('Geographic Distribution of Data Accessibility\n'
                 'Work Visa Medical Requirements by Language of Information Source')

    handles = [mpatches.Patch(facecolor=c, edgecolor='#666', label=l) for l, c in labs]
    ax.legend(handles=handles, loc='lower left', fontsize=9,
              framealpha=0.95, edgecolor='gray', fancybox=True)
    ax.set_xlim(-180, 180)
    ax.set_ylim(-60, 85)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title(title, fontsize=14, fontweight='bold', pad=15)
    plt.tight_layout()

    sfx = 'ja' if is_ja else 'en'
    for ext in ('png', 'tiff'):
        path = os.path.join(OUTPUT_DIR, f"fig7_accessibility_{sfx}.{ext}")
        plt.savefig(path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()
    print(f"Saved fig7_accessibility_{sfx} (png + tiff)")


def get_unreachable_by_region():
    """Unreachable countries by region."""
    africa = set(['Benin', 'Burkina Faso', 'Burundi', 'Cameroon', 'Chad', 'Comoros',
                  'Congo', 'Djibouti', 'Eswatini', 'Gabon', 'Gambia', 'Guinea',
                  'Lesotho', 'Liberia', 'Madagascar', 'Mali', 'Mauritania', 'Mauritius',
                  'Niger', 'Seychelles', 'Sierra Leone', 'Somalia', 'South Sudan',
                  'Sudan', 'Togo', 'Cabo Verde', 'Central African Republic',
                  "Cote d'Ivoire", 'DR Congo', 'Equatorial Guinea', 'Guinea-Bissau',
                  'Sao Tome and Principe'])
    asia = set(['Afghanistan', 'Armenia', 'Azerbaijan', 'Bhutan', 'Georgia',
                'Iran', 'Kazakhstan', 'Kyrgyzstan', 'Laos', 'Maldives', 'Mongolia',
                'North Korea', 'Palestine', 'Syria', 'Tajikistan', 'Timor-Leste',
                'Turkey', 'Turkmenistan', 'Uzbekistan', 'Yemen'])
    europe = set(['Albania', 'Andorra', 'Belarus', 'Bosnia and Herzegovina',
                  'Kosovo', 'Moldova', 'Montenegro',
                  'North Macedonia', 'San Marino', 'Serbia', 'Ukraine'])
    americas = set(['Antigua and Barbuda', 'Bahamas', 'Belize', 'Dominica', 'Grenada',
                    'Guyana', 'Saint Kitts and Nevis', 'Saint Lucia',
                    'Saint Vincent and the Grenadines', 'Suriname'])
    oceania = set(['Kiribati', 'Marshall Islands', 'Micronesia', 'Nauru', 'Palau',
                   'Samoa', 'Solomon Islands', 'Tuvalu', 'Vanuatu'])

    result = {'Africa': [], 'Asia': [], 'Europe': [], 'Americas': [], 'Oceania': []}
    for c in UNREACHABLE:
        placed = False
        for region, rset in [('Africa', africa), ('Asia', asia), ('Europe', europe),
                             ('Americas', americas), ('Oceania', oceania)]:
            if c in rset:
                result[region].append(c)
                placed = True
                break
        if not placed:
            result.setdefault('Other', []).append(c)
    return {k: sorted(v) for k, v in result.items() if v}


def save_data():
    """Save classification as JSON for DOCX/PPTX scripts."""
    data = {
        'english_reached_count': len(ENGLISH_REACHED),
        'multilingual_count': len(MULTILINGUAL),
        'unreachable_count': len(UNREACHABLE),
        'multilingual_countries': MULTILINGUAL,
        'unreachable_by_region': get_unreachable_by_region(),
        'unreachable_languages': {c: UNREACHABLE_LANGS.get(c, '') for c in sorted(UNREACHABLE)},
    }
    out = os.path.join(OUTPUT_DIR, 'country_classification.json')
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"Saved: {out}")


if __name__ == '__main__':
    ok = verify()
    regions = get_unreachable_by_region()
    i = 0
    for region, countries in regions.items():
        print(f"\n{region} ({len(countries)}):")
        for c in countries:
            i += 1
            print(f"  {i:3d}. {c} ({UNREACHABLE_LANGS.get(c, 'N/A')})")

    print("\nCreating maps...")
    _create_accessibility_map('en')
    _create_accessibility_map('ja')
    save_data()
    print("Done!")
