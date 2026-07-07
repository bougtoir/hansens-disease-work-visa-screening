#!/usr/bin/env python3
"""Generate color figures for the scoping review papers."""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os

OUTPUT_DIR = "/home/ubuntu/scoping_review/figures"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Set global font settings
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 11

# ============================================================
# FIGURE 1: World Map - Countries screening for Hansen's disease
# ============================================================
def create_world_map_figure():
    """Create a schematic regional map showing Hansen's disease screening status."""
    fig, ax = plt.subplots(1, 1, figsize=(14, 8))
    
    # Define regions with approximate positions on a simplified world layout
    regions = {
        'GCC States\n(6 countries)': {'pos': (0.58, 0.45), 'color': '#D32F2F', 'size': 1800},
        'Southeast/East Asia\n(5 countries/territories)': {'pos': (0.78, 0.48), 'color': '#D32F2F', 'size': 2200},
        'Southern Africa\n(2 countries)': {'pos': (0.52, 0.25), 'color': '#D32F2F', 'size': 1200},
        'Russia': {'pos': (0.65, 0.78), 'color': '#D32F2F', 'size': 1000},
        'United States\n& Caribbean': {'pos': (0.22, 0.55), 'color': '#D32F2F', 'size': 1500},
        'Malta': {'pos': (0.48, 0.58), 'color': '#D32F2F', 'size': 600},
        'India': {'pos': (0.70, 0.42), 'color': '#FF9800', 'size': 1200},
        'EU/Schengen\n(26 countries)': {'pos': (0.45, 0.72), 'color': '#4CAF50', 'size': 2500},
        'Japan, S. Korea\nSingapore': {'pos': (0.85, 0.62), 'color': '#2196F3', 'size': 1400},
        'Canada, Australia\nNew Zealand, UK': {'pos': (0.18, 0.78), 'color': '#2196F3', 'size': 1800},
    }
    
    for label, info in regions.items():
        ax.scatter(info['pos'][0], info['pos'][1], s=info['size'], 
                   c=info['color'], alpha=0.6, edgecolors='black', linewidth=0.5, zorder=3)
        ax.annotate(label, info['pos'], fontsize=7.5, ha='center', va='center',
                    fontweight='bold', zorder=4)
    
    # Legend
    legend_elements = [
        mpatches.Patch(facecolor='#D32F2F', alpha=0.6, edgecolor='black', label='Hansen\'s disease explicitly named (20 countries)'),
        mpatches.Patch(facecolor='#FF9800', alpha=0.6, edgecolor='black', label='Domestic employment laws only (India)'),
        mpatches.Patch(facecolor='#2196F3', alpha=0.6, edgecolor='black', label='Disease-specific screening, NO Hansen\'s disease (8+ countries)'),
        mpatches.Patch(facecolor='#4CAF50', alpha=0.6, edgecolor='black', label='No disease-specific medical exam required (52+ countries)'),
    ]
    ax.legend(handles=legend_elements, loc='lower left', fontsize=8.5, 
              framealpha=0.95, edgecolor='gray')
    
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("Figure 1: Global Distribution of Hansen's Disease Screening\nin Work Visa Medical Examinations (2026)",
                 fontsize=13, fontweight='bold', pad=15)
    
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "fig1_world_map.png"), dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.savefig(os.path.join(OUTPUT_DIR, "fig1_world_map.tiff"), dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()


# ============================================================
# FIGURE 2: Bar chart - Most commonly screened diseases
# ============================================================
def create_disease_bar_chart():
    """Create horizontal bar chart of most commonly screened diseases."""
    fig, ax = plt.subplots(figsize=(10, 6))
    
    diseases = [
        'Tuberculosis (TB)', 'HIV/AIDS', 'Syphilis/VD', 'Hepatitis B',
        "Hansen's disease\n(Leprosy)", 'Drug addiction', 'Mental illness',
        'Hepatitis C', 'Malaria', 'Elephantiasis', 'Cholera',
        'Trachoma', 'Plague', 'Yellow fever'
    ]
    counts = [55, 48, 38, 30, 20, 25, 18, 12, 8, 4, 3, 2, 2, 2]
    percentages = [c/58*100 for c in counts]
    
    # Sort by count
    sorted_data = sorted(zip(diseases, counts, percentages), key=lambda x: x[1])
    diseases_sorted = [d[0] for d in sorted_data]
    counts_sorted = [d[1] for d in sorted_data]
    pct_sorted = [d[2] for d in sorted_data]
    
    # Color Hansen's disease differently
    colors = []
    for d in diseases_sorted:
        if "Hansen" in d:
            colors.append('#D32F2F')
        elif d in ['Tuberculosis (TB)', 'HIV/AIDS', 'Syphilis/VD', 'Hepatitis B']:
            colors.append('#1976D2')
        else:
            colors.append('#78909C')
    
    bars = ax.barh(range(len(diseases_sorted)), counts_sorted, color=colors, 
                    edgecolor='white', height=0.7, alpha=0.85)
    
    # Add count and percentage labels
    for i, (count, pct) in enumerate(zip(counts_sorted, pct_sorted)):
        ax.text(count + 0.8, i, f'{count} ({pct:.0f}%)', va='center', fontsize=9)
    
    ax.set_yticks(range(len(diseases_sorted)))
    ax.set_yticklabels(diseases_sorted, fontsize=9)
    ax.set_xlabel('Number of Countries (out of 58 with disease-specific screening)', fontsize=10)
    ax.set_title("Figure 2: Diseases Named in Work Visa Medical Screening Requirements\nAcross 58 Countries with Disease-Specific Screening",
                 fontsize=12, fontweight='bold', pad=15)
    ax.set_xlim(0, 65)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    # Add annotation for Hansen's disease
    hansen_idx = diseases_sorted.index("Hansen's disease\n(Leprosy)")
    ax.annotate('5th most commonly\nscreened disease', 
                xy=(counts_sorted[hansen_idx], hansen_idx),
                xytext=(42, hansen_idx - 1.5),
                fontsize=9, color='#D32F2F', fontweight='bold',
                arrowprops=dict(arrowstyle='->', color='#D32F2F', lw=1.5))
    
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "fig2_disease_bar.png"), dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.savefig(os.path.join(OUTPUT_DIR, "fig2_disease_bar.tiff"), dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()


# ============================================================
# FIGURE 3: Regional breakdown pie/donut chart
# ============================================================
def create_regional_donut():
    """Create donut chart showing regional distribution of Hansen's disease screening."""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    
    # Left: Regional distribution of 20 countries with Hansen's screening
    regions = ['GCC/Middle East\n(6)', 'Southeast/\nEast Asia (5)', 'Americas (3)', 
               'Africa (2)', 'Europe (2)', 'South Asia (2)']
    sizes = [6, 5, 3, 2, 2, 2]
    colors = ['#E53935', '#FF7043', '#FFA726', '#66BB6A', '#42A5F5', '#AB47BC']
    explode = (0.05, 0.05, 0.05, 0.05, 0.05, 0.05)
    
    wedges, texts, autotexts = ax1.pie(sizes, labels=regions, autopct='%1.0f%%',
                                        colors=colors, explode=explode,
                                        pctdistance=0.75, labeldistance=1.15,
                                        textprops={'fontsize': 9})
    for autotext in autotexts:
        autotext.set_fontsize(9)
        autotext.set_fontweight('bold')
    
    # Draw center circle for donut effect
    centre_circle = plt.Circle((0, 0), 0.45, fc='white')
    ax1.add_artist(centre_circle)
    ax1.text(0, 0, '20\ncountries', ha='center', va='center', fontsize=14, fontweight='bold')
    ax1.set_title('(A) Regional Distribution of Countries\nScreening for Hansen\'s Disease',
                  fontsize=11, fontweight='bold', pad=15)
    
    # Right: Provision type breakdown
    provision_types = ['Automatic\nexclusion\n(GCC 6)', 
                       'Certification\nof absence (7)',
                       'Class-based\nw/ treatment (1)',
                       'Medical\nunfitness (2)',
                       'General\nprohibition (4)']
    prov_sizes = [6, 7, 1, 2, 4]
    prov_colors = ['#C62828', '#E53935', '#FF8F00', '#F4511E', '#AD1457']
    
    wedges2, texts2, autotexts2 = ax2.pie(prov_sizes, labels=provision_types, autopct='%1.0f%%',
                                           colors=prov_colors, 
                                           pctdistance=0.75, labeldistance=1.2,
                                           textprops={'fontsize': 9})
    for autotext in autotexts2:
        autotext.set_fontsize(9)
        autotext.set_fontweight('bold')
    
    centre_circle2 = plt.Circle((0, 0), 0.45, fc='white')
    ax2.add_artist(centre_circle2)
    ax2.text(0, 0, '20\ncountries', ha='center', va='center', fontsize=14, fontweight='bold')
    ax2.set_title('(B) Nature of Hansen\'s Disease\nProvisions in Immigration Law',
                  fontsize=11, fontweight='bold', pad=15)
    
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "fig3_regional_donut.png"), dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.savefig(os.path.join(OUTPUT_DIR, "fig3_regional_donut.tiff"), dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()


# ============================================================
# FIGURE 4: PRISMA-ScR Flow Diagram
# ============================================================
def create_prisma_flow():
    """Create PRISMA-ScR flow diagram."""
    fig, ax = plt.subplots(figsize=(10, 12))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 14)
    ax.axis('off')
    
    def draw_box(x, y, w, h, text, color='#E3F2FD', border='#1565C0', fontsize=9):
        rect = mpatches.FancyBboxPatch((x - w/2, y - h/2), w, h,
                                         boxstyle="round,pad=0.15",
                                         facecolor=color, edgecolor=border, linewidth=1.5)
        ax.add_patch(rect)
        ax.text(x, y, text, ha='center', va='center', fontsize=fontsize, wrap=True,
                multialignment='center')
    
    def draw_arrow(x1, y1, x2, y2):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle='->', color='#1565C0', lw=1.5))
    
    # Title
    ax.text(5, 13.5, 'PRISMA-ScR Flow Diagram', fontsize=14, fontweight='bold',
            ha='center', va='center')
    
    # IDENTIFICATION
    ax.text(0.5, 12.5, 'IDENTIFICATION', fontsize=10, fontweight='bold', color='#1565C0',
            rotation=90, va='center')
    draw_box(5, 12.5, 7, 1.0,
             'Countries/territories identified for review\n'
             '(n = 197: 193 UN member states + 4 observer entities)',
             color='#BBDEFB')
    
    draw_arrow(5, 12.0, 5, 11.3)
    
    # SCREENING
    ax.text(0.5, 10.5, 'SCREENING', fontsize=10, fontweight='bold', color='#1565C0',
            rotation=90, va='center')
    draw_box(5, 10.8, 7, 1.0,
             'Sources screened per country:\n'
             'Government websites, legislation, official forms,\n'
             'ILEP database, IOM reports, academic literature')
    
    draw_arrow(5, 10.3, 5, 9.5)
    
    # Results split
    draw_box(5, 9.0, 7, 1.0,
             'Countries with publicly available information\n'
             'on work visa medical requirements identified\n(n = 145)',
             color='#C8E6C9')
    
    # Excluded box
    draw_box(9, 9.0, 1.5, 1.0,
             'Not available\nin English\n(n = 52)',
             color='#FFCDD2', border='#C62828', fontsize=8)
    draw_arrow(8, 9.0, 8.2, 9.0)
    
    draw_arrow(5, 8.5, 5, 7.7)
    
    # ELIGIBILITY
    ax.text(0.5, 7.5, 'ELIGIBILITY', fontsize=10, fontweight='bold', color='#1565C0',
            rotation=90, va='center')
    draw_box(5, 7.2, 7, 1.0,
             'Countries with confirmed disease-specific\n'
             'screening requirements for work visas\n(n = 58)',
             color='#C8E6C9')
    
    draw_box(9, 7.2, 1.5, 1.0,
             'Medical exam\nrequired but\nno named\ndiseases (n=87)',
             color='#FFF9C4', border='#F9A825', fontsize=7)
    draw_arrow(8, 7.2, 8.2, 7.2)
    
    draw_arrow(5, 6.7, 5, 6.0)
    
    # INCLUSION
    ax.text(0.5, 4.5, 'INCLUDED', fontsize=10, fontweight='bold', color='#1565C0',
            rotation=90, va='center')
    
    draw_box(3, 5.2, 3.5, 1.2,
             'Countries with Hansen\'s\ndisease explicitly named\n'
             'in work visa medical\nrequirements\n(n = 20)',
             color='#FFCDD2', border='#C62828')
    
    draw_box(7, 5.2, 3.5, 1.2,
             'Countries with disease-\nspecific screening but\n'
             'NO Hansen\'s disease\n(n = 38)',
             color='#E8F5E9', border='#2E7D32')
    
    draw_arrow(4, 6.0, 3, 5.8)
    draw_arrow(6, 6.0, 7, 5.8)
    
    # Final summary boxes
    draw_box(3, 3.5, 3.5, 1.5,
             'By region:\n'
             'GCC/Middle East: 6\n'
             'SE/East Asia: 5\n'
             'Americas: 3\n'
             'Africa: 2\n'
             'Europe: 2\n'
             'South Asia: 2',
             color='#FCE4EC', border='#C62828', fontsize=8)
    
    draw_box(7, 3.5, 3.5, 1.5,
             'Top screened diseases\n(without Hansen\'s):\n'
             'TB: 38 countries\n'
             'HIV: 35 countries\n'
             'Syphilis: 28 countries\n'
             'Hepatitis B: 22 countries',
             color='#E8F5E9', border='#2E7D32', fontsize=8)
    
    draw_arrow(3, 4.6, 3, 4.3)
    draw_arrow(7, 4.6, 7, 4.3)
    
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "fig4_prisma_flow.png"), dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.savefig(os.path.join(OUTPUT_DIR, "fig4_prisma_flow.tiff"), dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()


# ============================================================
# FIGURE 5: Comparison chart - Transmissibility vs Screening
# ============================================================
def create_transmissibility_chart():
    """Compare transmissibility of screened diseases vs. screening frequency."""
    fig, ax = plt.subplots(figsize=(10, 7))
    
    diseases = ['TB', 'HIV', 'Syphilis', 'Hepatitis B', "Hansen's\ndisease", 
                'Hepatitis C', 'Malaria', 'Cholera']
    
    # Approximate R0 / transmissibility index (normalized 0-10 scale)
    transmissibility = [4.5, 2.5, 3.0, 5.0, 0.5, 1.5, 8.0, 5.5]
    
    # Number of countries screening
    screening_countries = [55, 48, 38, 30, 20, 12, 8, 3]
    
    # Bubble size proportional to screening countries
    sizes = [s * 20 for s in screening_countries]
    
    colors_list = ['#1976D2', '#1976D2', '#1976D2', '#1976D2', '#D32F2F', 
                   '#78909C', '#78909C', '#78909C']
    
    scatter = ax.scatter(transmissibility, screening_countries, s=sizes, 
                          c=colors_list, alpha=0.7, edgecolors='black', linewidth=0.8, zorder=3)
    
    # Add labels
    for i, disease in enumerate(diseases):
        offset_x = 0.2
        offset_y = 1.5
        if disease == "Hansen's\ndisease":
            offset_x = 0.3
            offset_y = 2.5
        ax.annotate(disease, (transmissibility[i], screening_countries[i]),
                    xytext=(transmissibility[i] + offset_x, screening_countries[i] + offset_y),
                    fontsize=9, fontweight='bold' if "Hansen" in disease else 'normal',
                    color='#D32F2F' if "Hansen" in disease else 'black')
    
    # Add a shaded box highlighting the "gap" for Hansen's disease
    ax.axhspan(15, 25, xmin=0, xmax=0.15, alpha=0.15, color='#D32F2F')
    ax.annotate("Evidence-policy\ngap", xy=(0.5, 22), fontsize=10, 
                color='#D32F2F', fontweight='bold', fontstyle='italic',
                ha='center')
    
    ax.set_xlabel('Relative Transmissibility Index\n(higher = more transmissible)', fontsize=11)
    ax.set_ylabel('Number of Countries Screening\nfor This Disease in Work Visas', fontsize=11)
    ax.set_title("Figure 3: Disease Transmissibility vs. Number of Countries\nRequiring Screening in Work Visa Medical Examinations",
                 fontsize=12, fontweight='bold', pad=15)
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_xlim(-0.5, 9.5)
    ax.set_ylim(-2, 62)
    
    # Legend
    legend_elements = [
        mpatches.Patch(facecolor='#D32F2F', alpha=0.7, edgecolor='black', label="Hansen's disease (low transmissibility, high screening)"),
        mpatches.Patch(facecolor='#1976D2', alpha=0.7, edgecolor='black', label='Top 4 screened diseases'),
        mpatches.Patch(facecolor='#78909C', alpha=0.7, edgecolor='black', label='Other screened diseases'),
    ]
    ax.legend(handles=legend_elements, loc='upper right', fontsize=9, framealpha=0.9)
    
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "fig5_transmissibility.png"), dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.savefig(os.path.join(OUTPUT_DIR, "fig5_transmissibility.tiff"), dpi=300, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()


if __name__ == '__main__':
    print("Generating figures...")
    create_world_map_figure()
    print("  Figure 1: World map - done")
    create_disease_bar_chart()
    print("  Figure 2: Disease bar chart - done")
    create_regional_donut()
    print("  Figure 3: Regional donut - done")
    create_prisma_flow()
    print("  Figure 4: PRISMA flow - done")
    create_transmissibility_chart()
    print("  Figure 5: Transmissibility chart - done")
    print("All figures generated successfully!")
    print(f"Output directory: {OUTPUT_DIR}")
