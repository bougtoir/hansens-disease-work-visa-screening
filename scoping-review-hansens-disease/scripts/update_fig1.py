#!/usr/bin/env python3
"""Regenerate Figure 1 with proper world map base layer (白地図)."""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import geopandas as gpd
import numpy as np
import os

OUTPUT_DIR = "/home/ubuntu/scoping_review/figures"
os.makedirs(OUTPUT_DIR, exist_ok=True)

plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 11


def create_world_map_figure():
    """Create world map with proper country outlines and Hansen's disease screening overlay."""

    # Download Natural Earth data
    world = gpd.read_file('https://naciscdn.org/naturalearth/110m/cultural/ne_110m_admin_0_countries.zip')

    fig, ax = plt.subplots(1, 1, figsize=(16, 9))

    # --- Category assignments ---
    # Hansen's disease screening countries (Category A - red)
    hansen_countries = [
        'Saudi Arabia', 'United Arab Emirates', 'Qatar', 'Kuwait', 'Oman', 'Bahrain',
        'China', 'Thailand', 'Malaysia', 'Philippines',
        'South Africa', 'Namibia',
        'Russia', 'Malta',
        'United States of America', 'Barbados',
        'India',
    ]
    # Note: Taiwan, Hong Kong SAR, US Virgin Islands not in Natural Earth as separate

    # Disease-specific screening but NO Hansen's (Category B - blue)
    screening_no_hansen = [
        'Canada', 'Australia', 'United Kingdom', 'New Zealand', 'Japan',
        'South Korea', 'Singapore', 'Israel', 'Kenya', 'Nigeria',
        'Ghana', 'Chile', 'Ethiopia', 'Tanzania', 'Uganda',
        'Zambia', 'Zimbabwe', 'Mozambique', 'Angola', 'Iraq',
        'Jordan', 'Lebanon', 'Egypt', 'Morocco', 'Tunisia',
        'Libya', 'Algeria', 'Pakistan', 'Bangladesh', 'Sri Lanka',
        'Nepal', 'Myanmar', 'Vietnam', 'Cambodia', 'Indonesia',
        'Brunei', 'Papua New Guinea',
    ]

    # No disease-specific exam (Category C - green) - EU/Schengen + others
    no_screening = [
        'France', 'Germany', 'Italy', 'Spain', 'Portugal', 'Netherlands',
        'Belgium', 'Luxembourg', 'Austria', 'Switzerland', 'Sweden',
        'Norway', 'Denmark', 'Finland', 'Iceland', 'Ireland',
        'Greece', 'Poland', 'Czech Republic', 'Slovakia', 'Hungary',
        'Romania', 'Bulgaria', 'Croatia', 'Slovenia', 'Estonia',
        'Latvia', 'Lithuania', 'Cyprus',
        'Brazil', 'Mexico', 'Argentina', 'Colombia', 'Peru',
        'Venezuela', 'Ecuador', 'Bolivia', 'Paraguay', 'Uruguay',
        'Costa Rica', 'Panama', 'Guatemala', 'Honduras', 'El Salvador',
        'Nicaragua', 'Dominican Republic', 'Cuba', 'Jamaica', 'Trinidad and Tobago',
    ]

    # Color mapping
    def get_color(name):
        if name in hansen_countries:
            return '#D32F2F'  # Red
        elif name in screening_no_hansen:
            return '#42A5F5'  # Blue
        elif name in no_screening:
            return '#66BB6A'  # Green
        else:
            return '#E0E0E0'  # Light gray (no data)

    # Match country names (Natural Earth uses different names for some)
    name_map = {
        'United States of America': 'United States of America',
        'S. Korea': 'South Korea',
        'Dem. Rep. Congo': 'Dem. Rep. Congo',
    }

    world['color'] = world['NAME'].apply(
        lambda x: get_color(name_map.get(x, x))
    )

    # Draw base map
    world.plot(ax=ax, color=world['color'], edgecolor='#999999', linewidth=0.3)

    # Add markers for small countries / territories not visible on map
    small_markers = {
        # Hansen's screening (red markers)
        'Taiwan': (121, 23.5, '#D32F2F'),
        'Hong Kong': (114.2, 22.3, '#D32F2F'),
        'Bahrain': (50.5, 26.0, '#D32F2F'),
        'Malta': (14.4, 35.9, '#D32F2F'),
        'Barbados': (-59.6, 13.1, '#D32F2F'),
        'USVI': (-64.9, 18.3, '#D32F2F'),
        # No screening (green markers for small EU)
        'Luxembourg': (6.1, 49.8, '#66BB6A'),
        'Cyprus': (33.4, 35.1, '#66BB6A'),
        'Iceland': (-19.0, 65.0, '#66BB6A'),
    }

    for name, (lon, lat, color) in small_markers.items():
        ax.plot(lon, lat, 'o', color=color, markersize=6, markeredgecolor='black',
                markeredgewidth=0.5, zorder=5)

    # Add region labels with counts
    labels = [
        (45, 20, 'GCC States\n(6 countries)', '#D32F2F', 10),
        (105, 15, 'East/SE Asia\n(5 countries/\nterritories)', '#D32F2F', 9),
        (25, -15, 'Southern Africa\n(2 countries)', '#D32F2F', 9),
        (50, 60, 'Russia', '#D32F2F', 9),
        (-95, 40, 'United States\n& Caribbean', '#D32F2F', 9),
        (80, 25, 'India', '#FF9800', 9),
        (10, 52, 'EU/Schengen\n(no disease-specific\nexam, 26+ countries)', '#2E7D32', 8),
        (135, 38, 'Japan, S. Korea\nSingapore', '#1565C0', 8),
        (-120, 55, 'Canada, Australia\nNZ, UK', '#1565C0', 8),
    ]

    for lon, lat, text, color, fontsize in labels:
        ax.annotate(text, xy=(lon, lat), fontsize=fontsize, fontweight='bold',
                    color=color, ha='center', va='center',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.85,
                              edgecolor=color, linewidth=0.8),
                    zorder=6)

    # Legend
    legend_elements = [
        mpatches.Patch(facecolor='#D32F2F', edgecolor='#999', label="Hansen's disease explicitly named (20 countries)"),
        mpatches.Patch(facecolor='#FF9800', edgecolor='#999', label='Domestic employment laws only (India)'),
        mpatches.Patch(facecolor='#42A5F5', edgecolor='#999', label="Disease-specific screening, NO Hansen's (38 countries)"),
        mpatches.Patch(facecolor='#66BB6A', edgecolor='#999', label='No disease-specific medical exam required (52+ countries)'),
        mpatches.Patch(facecolor='#E0E0E0', edgecolor='#999', label='Information not publicly available / other'),
    ]
    ax.legend(handles=legend_elements, loc='lower left', fontsize=9,
              framealpha=0.95, edgecolor='gray', fancybox=True)

    ax.set_xlim(-180, 180)
    ax.set_ylim(-60, 85)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title(
        "Figure 1: Global Distribution of Hansen's Disease Screening\n"
        "in Work Visa Medical Examinations (2026)",
        fontsize=14, fontweight='bold', pad=15
    )

    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "fig1_world_map.png"), dpi=300,
                bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.savefig(os.path.join(OUTPUT_DIR, "fig1_world_map.tiff"), dpi=300,
                bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()
    print("Saved: fig1_world_map.png and fig1_world_map.tiff")


if __name__ == '__main__':
    create_world_map_figure()
