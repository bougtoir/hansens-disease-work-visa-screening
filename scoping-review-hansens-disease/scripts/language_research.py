#!/usr/bin/env python3
"""
Research countries with insufficient English-language data by using official language information.
Identify Category C countries and analyze by official language.
"""

# Category A: Hansen's disease explicitly named (20)
category_a = [
    'South Africa', 'Namibia', 'Saudi Arabia', 'UAE', 'Qatar', 'Kuwait', 'Oman', 'Bahrain',
    'Thailand', 'China', 'Taiwan', 'Malaysia', 'Russia', 'United States', 'Barbados',
    'US Virgin Islands', 'Malta', 'India', 'Philippines', 'Hong Kong SAR'
]

# Category B: Disease-specific screening confirmed, no Hansen's (38)
category_b_confirmed = [
    'Singapore', 'South Korea', 'Japan', 'Canada', 'Australia', 'New Zealand',
    'United Kingdom', 'Israel', 'Nigeria', 'Kenya', 'Ethiopia', 'Ghana', 'Senegal',
    'Brazil', 'Mexico'
]

# Category D: No disease-specific medical exam (EU/Schengen + others) (52)
category_d = [
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
]

# All 197 countries/territories
all_197 = [
    'Algeria', 'Angola', 'Benin', 'Botswana', 'Burkina Faso', 'Burundi', 'Cabo Verde',
    'Cameroon', 'Central African Republic', 'Chad', 'Comoros', 'Congo', 'DR Congo',
    "Cote d'Ivoire", 'Djibouti', 'Egypt', 'Equatorial Guinea', 'Eritrea', 'Eswatini',
    'Ethiopia', 'Gabon', 'Gambia', 'Ghana', 'Guinea', 'Guinea-Bissau', 'Kenya',
    'Lesotho', 'Liberia', 'Libya', 'Madagascar', 'Malawi', 'Mali', 'Mauritania',
    'Mauritius', 'Morocco', 'Mozambique', 'Namibia', 'Niger', 'Nigeria', 'Rwanda',
    'Sao Tome and Principe', 'Senegal', 'Seychelles', 'Sierra Leone', 'Somalia',
    'South Africa', 'South Sudan', 'Sudan', 'Tanzania', 'Togo', 'Tunisia', 'Uganda',
    'Zambia', 'Zimbabwe',
    'Afghanistan', 'Armenia', 'Azerbaijan', 'Bahrain', 'Bangladesh', 'Bhutan', 'Brunei',
    'Cambodia', 'China', 'Cyprus', 'Georgia', 'India', 'Indonesia', 'Iran', 'Iraq',
    'Israel', 'Japan', 'Jordan', 'Kazakhstan', 'Kuwait', 'Kyrgyzstan', 'Laos', 'Lebanon',
    'Malaysia', 'Maldives', 'Mongolia', 'Myanmar', 'Nepal', 'North Korea', 'Oman',
    'Pakistan', 'Philippines', 'Qatar', 'Saudi Arabia', 'Singapore', 'South Korea',
    'Sri Lanka', 'Syria', 'Tajikistan', 'Thailand', 'Timor-Leste', 'Turkey',
    'Turkmenistan', 'UAE', 'Uzbekistan', 'Vietnam', 'Yemen',
    'Taiwan', 'Hong Kong SAR', 'Palestine',
    'Albania', 'Andorra', 'Armenia', 'Austria', 'Azerbaijan', 'Belarus', 'Belgium',
    'Bosnia and Herzegovina', 'Bulgaria', 'Croatia', 'Czech Republic', 'Denmark',
    'Estonia', 'Finland', 'France', 'Georgia', 'Germany', 'Greece', 'Hungary',
    'Iceland', 'Ireland', 'Italy', 'Kazakhstan', 'Kosovo', 'Latvia', 'Liechtenstein',
    'Lithuania', 'Luxembourg', 'Malta', 'Moldova', 'Monaco', 'Montenegro',
    'Netherlands', 'North Macedonia', 'Norway', 'Poland', 'Portugal', 'Romania',
    'Russia', 'San Marino', 'Serbia', 'Slovakia', 'Slovenia', 'Spain', 'Sweden',
    'Switzerland', 'Turkey', 'Ukraine', 'United Kingdom', 'Vatican City',
    'Antigua and Barbuda', 'Argentina', 'Bahamas', 'Barbados', 'Belize', 'Bolivia',
    'Brazil', 'Canada', 'Chile', 'Colombia', 'Costa Rica', 'Cuba', 'Dominica',
    'Dominican Republic', 'Ecuador', 'El Salvador', 'Grenada', 'Guatemala', 'Guyana',
    'Haiti', 'Honduras', 'Jamaica', 'Mexico', 'Nicaragua', 'Panama', 'Paraguay',
    'Peru', 'Saint Kitts and Nevis', 'Saint Lucia', 'Saint Vincent and the Grenadines',
    'Suriname', 'Trinidad and Tobago', 'United States', 'US Virgin Islands', 'Uruguay',
    'Venezuela',
    'Australia', 'Fiji', 'Kiribati', 'Marshall Islands', 'Micronesia', 'Nauru',
    'New Zealand', 'Palau', 'Papua New Guinea', 'Samoa', 'Solomon Islands', 'Tonga',
    'Tuvalu', 'Vanuatu',
]

# Official languages database
official_languages = {
    'Algeria': ['Arabic', 'Tamazight', 'French'],
    'Angola': ['Portuguese'],
    'Benin': ['French'],
    'Botswana': ['English', 'Setswana'],
    'Burkina Faso': ['French'],
    'Burundi': ['Kirundi', 'French', 'English'],
    'Cabo Verde': ['Portuguese'],
    'Cameroon': ['French', 'English'],
    'Central African Republic': ['French', 'Sango'],
    'Chad': ['French', 'Arabic'],
    'Comoros': ['Comorian', 'Arabic', 'French'],
    'Congo': ['French'],
    'DR Congo': ['French'],
    "Cote d'Ivoire": ['French'],
    'Djibouti': ['French', 'Arabic'],
    'Egypt': ['Arabic'],
    'Equatorial Guinea': ['Spanish', 'French', 'Portuguese'],
    'Eritrea': ['Tigrinya', 'Arabic', 'English'],
    'Eswatini': ['English', 'Swazi'],
    'Ethiopia': ['Amharic'],
    'Gabon': ['French'],
    'Gambia': ['English'],
    'Ghana': ['English'],
    'Guinea': ['French'],
    'Guinea-Bissau': ['Portuguese'],
    'Kenya': ['English', 'Swahili'],
    'Lesotho': ['Sesotho', 'English'],
    'Liberia': ['English'],
    'Libya': ['Arabic'],
    'Madagascar': ['Malagasy', 'French'],
    'Malawi': ['English', 'Chichewa'],
    'Mali': ['French'],
    'Mauritania': ['Arabic', 'French'],
    'Mauritius': ['English', 'French'],
    'Morocco': ['Arabic', 'Tamazight', 'French'],
    'Mozambique': ['Portuguese'],
    'Namibia': ['English'],
    'Niger': ['French'],
    'Nigeria': ['English'],
    'Rwanda': ['Kinyarwanda', 'French', 'English'],
    'Sao Tome and Principe': ['Portuguese'],
    'Senegal': ['French'],
    'Seychelles': ['Seychellois Creole', 'English', 'French'],
    'Sierra Leone': ['English'],
    'Somalia': ['Somali', 'Arabic'],
    'South Africa': ['English', 'Afrikaans', 'Zulu'],
    'South Sudan': ['English'],
    'Sudan': ['Arabic', 'English'],
    'Tanzania': ['Swahili', 'English'],
    'Togo': ['French'],
    'Tunisia': ['Arabic', 'French'],
    'Uganda': ['English', 'Swahili'],
    'Zambia': ['English'],
    'Zimbabwe': ['English', 'Shona', 'Ndebele'],
    'Afghanistan': ['Dari', 'Pashto'],
    'Armenia': ['Armenian'],
    'Azerbaijan': ['Azerbaijani'],
    'Bahrain': ['Arabic'],
    'Bangladesh': ['Bengali'],
    'Bhutan': ['Dzongkha'],
    'Brunei': ['Malay'],
    'Cambodia': ['Khmer'],
    'China': ['Mandarin Chinese'],
    'Georgia': ['Georgian'],
    'India': ['Hindi', 'English'],
    'Indonesia': ['Indonesian'],
    'Iran': ['Persian'],
    'Iraq': ['Arabic', 'Kurdish'],
    'Israel': ['Hebrew', 'Arabic'],
    'Japan': ['Japanese'],
    'Jordan': ['Arabic'],
    'Kazakhstan': ['Kazakh', 'Russian'],
    'Kuwait': ['Arabic'],
    'Kyrgyzstan': ['Kyrgyz', 'Russian'],
    'Laos': ['Lao'],
    'Lebanon': ['Arabic'],
    'Malaysia': ['Malay'],
    'Maldives': ['Dhivehi'],
    'Mongolia': ['Mongolian'],
    'Myanmar': ['Burmese'],
    'Nepal': ['Nepali'],
    'North Korea': ['Korean'],
    'Oman': ['Arabic'],
    'Pakistan': ['Urdu', 'English'],
    'Palestine': ['Arabic'],
    'Philippines': ['Filipino', 'English'],
    'Qatar': ['Arabic'],
    'Saudi Arabia': ['Arabic'],
    'Singapore': ['English', 'Malay', 'Mandarin Chinese', 'Tamil'],
    'South Korea': ['Korean'],
    'Sri Lanka': ['Sinhala', 'Tamil'],
    'Syria': ['Arabic'],
    'Taiwan': ['Mandarin Chinese'],
    'Tajikistan': ['Tajik', 'Russian'],
    'Thailand': ['Thai'],
    'Timor-Leste': ['Tetum', 'Portuguese'],
    'Turkey': ['Turkish'],
    'Turkmenistan': ['Turkmen'],
    'UAE': ['Arabic'],
    'Uzbekistan': ['Uzbek'],
    'Vietnam': ['Vietnamese'],
    'Yemen': ['Arabic'],
    'Hong Kong SAR': ['Chinese', 'English'],
    'Albania': ['Albanian'],
    'Andorra': ['Catalan'],
    'Austria': ['German'],
    'Belarus': ['Belarusian', 'Russian'],
    'Belgium': ['Dutch', 'French', 'German'],
    'Bosnia and Herzegovina': ['Bosnian', 'Croatian', 'Serbian'],
    'Bulgaria': ['Bulgarian'],
    'Croatia': ['Croatian'],
    'Cyprus': ['Greek', 'Turkish'],
    'Czech Republic': ['Czech'],
    'Denmark': ['Danish'],
    'Estonia': ['Estonian'],
    'Finland': ['Finnish', 'Swedish'],
    'France': ['French'],
    'Germany': ['German'],
    'Greece': ['Greek'],
    'Hungary': ['Hungarian'],
    'Iceland': ['Icelandic'],
    'Ireland': ['Irish', 'English'],
    'Italy': ['Italian'],
    'Kosovo': ['Albanian', 'Serbian'],
    'Latvia': ['Latvian'],
    'Liechtenstein': ['German'],
    'Lithuania': ['Lithuanian'],
    'Luxembourg': ['Luxembourgish', 'French', 'German'],
    'Malta': ['Maltese', 'English'],
    'Moldova': ['Romanian'],
    'Monaco': ['French'],
    'Montenegro': ['Montenegrin'],
    'Netherlands': ['Dutch'],
    'North Macedonia': ['Macedonian', 'Albanian'],
    'Norway': ['Norwegian'],
    'Poland': ['Polish'],
    'Portugal': ['Portuguese'],
    'Romania': ['Romanian'],
    'Russia': ['Russian'],
    'San Marino': ['Italian'],
    'Serbia': ['Serbian'],
    'Slovakia': ['Slovak'],
    'Slovenia': ['Slovenian'],
    'Spain': ['Spanish'],
    'Sweden': ['Swedish'],
    'Switzerland': ['German', 'French', 'Italian', 'Romansh'],
    'Ukraine': ['Ukrainian'],
    'United Kingdom': ['English'],
    'Vatican City': ['Italian', 'Latin'],
    'Antigua and Barbuda': ['English'],
    'Argentina': ['Spanish'],
    'Bahamas': ['English'],
    'Barbados': ['English'],
    'Belize': ['English'],
    'Bolivia': ['Spanish', 'Quechua', 'Aymara'],
    'Brazil': ['Portuguese'],
    'Canada': ['English', 'French'],
    'Chile': ['Spanish'],
    'Colombia': ['Spanish'],
    'Costa Rica': ['Spanish'],
    'Cuba': ['Spanish'],
    'Dominica': ['English'],
    'Dominican Republic': ['Spanish'],
    'Ecuador': ['Spanish'],
    'El Salvador': ['Spanish'],
    'Grenada': ['English'],
    'Guatemala': ['Spanish'],
    'Guyana': ['English'],
    'Haiti': ['Haitian Creole', 'French'],
    'Honduras': ['Spanish'],
    'Jamaica': ['English'],
    'Mexico': ['Spanish'],
    'Nicaragua': ['Spanish'],
    'Panama': ['Spanish'],
    'Paraguay': ['Spanish', 'Guarani'],
    'Peru': ['Spanish', 'Quechua', 'Aymara'],
    'Saint Kitts and Nevis': ['English'],
    'Saint Lucia': ['English'],
    'Saint Vincent and the Grenadines': ['English'],
    'Suriname': ['Dutch'],
    'Trinidad and Tobago': ['English'],
    'United States': ['English'],
    'US Virgin Islands': ['English'],
    'Uruguay': ['Spanish'],
    'Venezuela': ['Spanish'],
    'Australia': ['English'],
    'Fiji': ['English', 'Fijian', 'Hindi'],
    'Kiribati': ['English', 'Gilbertese'],
    'Marshall Islands': ['Marshallese', 'English'],
    'Micronesia': ['English'],
    'Nauru': ['Nauruan', 'English'],
    'New Zealand': ['English', 'Maori'],
    'Palau': ['Palauan', 'English'],
    'Papua New Guinea': ['Tok Pisin', 'English', 'Hiri Motu'],
    'Samoa': ['Samoan', 'English'],
    'Solomon Islands': ['English'],
    'Tonga': ['Tongan', 'English'],
    'Tuvalu': ['Tuvaluan', 'English'],
    'Vanuatu': ['Bislama', 'English', 'French'],
}

# Deduplicate
all_countries_set = set(all_197)
print(f"Total unique countries/territories: {len(all_countries_set)}")

categorized = set(category_a) | set(category_b_confirmed) | set(category_d)
category_c = sorted(all_countries_set - categorized)

print(f"\nCategory A (Hansen's named): {len(category_a)}")
print(f"Category B (confirmed no Hansen's): {len(category_b_confirmed)}")
print(f"Category C (unconfirmed diseases): {len(category_c)}")
print(f"Category D (no disease-specific exam): {len(category_d)}")
print(f"Total categorized: {len(category_a) + len(category_b_confirmed) + len(category_c) + len(category_d)}")

# Analyze by language
print("\n" + "="*60)
print("Category C Countries by Primary Official Language")
print("="*60)

lang_groups = {}
for country in category_c:
    langs = official_languages.get(country, ['Unknown'])
    primary = langs[0]
    if primary not in lang_groups:
        lang_groups[primary] = []
    lang_groups[primary].append(country)

for lang in sorted(lang_groups.keys(), key=lambda x: len(lang_groups[x]), reverse=True):
    countries = lang_groups[lang]
    print(f"\n{lang} ({len(countries)} countries):")
    for c in sorted(countries):
        all_langs = official_languages.get(c, ['Unknown'])
        print(f"  - {c} ({', '.join(all_langs)})")

# Summary statistics
print("\n" + "="*60)
print("Language Distribution Summary (Category C)")
print("="*60)
for lang in sorted(lang_groups.keys(), key=lambda x: len(lang_groups[x]), reverse=True):
    n = len(lang_groups[lang])
    pct = n / len(category_c) * 100
    print(f"  {lang}: {n} countries ({pct:.1f}%)")

# Count how many have English as ANY official language
english_accessible = [c for c in category_c if 'English' in official_languages.get(c, [])]
non_english = [c for c in category_c if 'English' not in official_languages.get(c, [])]
print(f"\nEnglish as official/co-official: {len(english_accessible)} ({len(english_accessible)/len(category_c)*100:.1f}%)")
print(f"No English as official language: {len(non_english)} ({len(non_english)/len(category_c)*100:.1f}%)")

print("\nNon-English Category C countries:")
for c in sorted(non_english):
    langs = official_languages.get(c, ['Unknown'])
    print(f"  - {c}: {', '.join(langs)}")
