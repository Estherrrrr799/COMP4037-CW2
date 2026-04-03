"""
NHS Hospital Admissions Visualisation Code

Output files:
  nhs_main_combined.png  <- Main figure for submission (two heatmaps merged vertically)
  nhs_trend.png          <- Supplementary figure 1: 26-year emergency trend
  nhs_covid_change.png   <- Supplementary figure 2: COVID-19 impact by category
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec, GridSpecFromSubplotSpec
import re
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# Set the path to the folder containing your CSV files
# ============================================================
DATA_FOLDER = r"./"

# ============================================================
# Load data
# ============================================================
heatmap_emg = pd.read_csv(f"{DATA_FOLDER}nhs_heatmap_emergency.csv", index_col=0)
heatmap_adm = pd.read_csv(f"{DATA_FOLDER}nhs_heatmap_admissions.csv", index_col=0)
by_chapter  = pd.read_csv(f"{DATA_FOLDER}nhs_by_chapter_year.csv")
covid       = pd.read_csv(f"{DATA_FOLDER}nhs_covid_impact.csv")

# ============================================================
# Global data cleaning
# ============================================================
# Remove the residual "Other" category and the incomplete 2012-13 year
for df in [heatmap_emg, heatmap_adm]:
    df.drop(index='Other', errors='ignore', inplace=True)
    df.drop(columns='2012-13', errors='ignore', inplace=True)

by_chapter = by_chapter[by_chapter['icd_chapter'] != 'Other'].copy()
covid      = covid[covid['icd_chapter'] != 'Other'].copy()

# Fix Nervous system 2008-09 anomalous value (~25k vs expected ~110k)
# Replaced with the mean of adjacent years
for df in [heatmap_emg, heatmap_adm]:
    nerv = [i for i in df.index if 'Nervous' in str(i)]
    if nerv and '2008-09' in df.columns:
        ni = nerv[0]
        df.loc[ni, '2008-09'] = (df.loc[ni, '2007-08'] + df.loc[ni, '2009-10']) / 2

# ============================================================
# Utility functions
# ============================================================
def shorten_label(label):
    """Strip chapter prefix (e.g. 'Ch.XIV  ') and return the disease description only."""
    label = str(label).strip()
    label = re.sub(r'^Ch\.[IVX]+/[IVX]+\s+', '', label)
    label = re.sub(r'^Ch\.[IVX]+\s+', '', label)
    return label.strip()


def add_covid_line(ax, years, n_cats, color='#D32F2F', top_label=True):
    """Draw a dashed red rectangle marking the COVID-19 lockdown year (2020-21)."""
    if '2020-21' not in years:
        return
    xi = years.index('2020-21')
    ax.add_patch(plt.Rectangle(
        (xi, -0.05), 1, n_cats + 0.05,
        fill=False, edgecolor=color, linewidth=2.2, linestyle='--', zorder=5))
    if top_label:
        ax.text(xi + 0.5, n_cats + 0.25, 'COVID-19\nLockdown',
                ha='center', va='bottom', fontsize=7.5,
                color=color, fontweight='bold')


def draw_cells(ax, data, data_norm, years, cats, cmap,
               show_val=True, threshold=0.0, fmt_fn=None, bg='#F5F7FA'):
    """Draw all heatmap cells with colour and optional text labels."""
    for yi, year in enumerate(years):
        for ci, cat in enumerate(cats):
            val      = data.loc[cat, year]
            val_norm = data_norm.loc[cat, year]
            if pd.isna(val) or pd.isna(val_norm):
                ax.add_patch(plt.Rectangle((yi, ci), 1, 1,
                             color='#EBEBEB', linewidth=0.3, edgecolor=bg))
                continue
            color = cmap(float(val_norm))
            ax.add_patch(plt.Rectangle((yi, ci), 1, 1,
                         color=color, linewidth=0.3, edgecolor=bg))
            if show_val and val_norm >= threshold:
                lum = 0.299*color[0] + 0.587*color[1] + 0.114*color[2]
                tc  = '#111111' if lum > 0.45 else '#FFFFFF'
                txt = fmt_fn(val) if fmt_fn else f'{val:.0f}'
                ax.text(yi + 0.5, ci + 0.5, txt,
                        ha='center', va='center',
                        fontsize=6.0, color=tc, fontweight='600')


def style_heatmap_ax(ax, years, cats, fontsize_x=8, fontsize_y=9.5,
                     muted='#606070', dark='#1A1A2E'):
    """Apply standard axis styling to a heatmap panel."""
    ax.set_xlim(0, len(years))
    ax.set_ylim(0, len(cats))
    ax.set_xticks([i + 0.5 for i in range(len(years))])
    ax.set_xticklabels(years, rotation=45, ha='right',
                       fontsize=fontsize_x, color=muted)
    ax.set_yticks([i + 0.5 for i in range(len(cats))])
    ax.set_yticklabels(cats, fontsize=fontsize_y, color=dark, fontweight='500')
    ax.tick_params(length=0)
    for sp in ax.spines.values():
        sp.set_visible(False)


# ============================================================
# Colour schemes
# ============================================================
# Figure 1 (emergency volume) - yellow -> orange -> red -> deep purple (row-normalised)
CMAP_VOL = mcolors.LinearSegmentedColormap.from_list('vol', [
    '#FFFDE7', '#FFF176', '#FFB300', '#E65100', '#B71C1C', '#4A148C'
], N=256)

# Figure 2 (emergency ratio) - white -> light blue -> deep blue (absolute values)
CMAP_RATIO = mcolors.LinearSegmentedColormap.from_list('ratio', [
    '#F7FBFF', '#C6DBEF', '#6BAED6', '#2171B5', '#084594', '#08306B'
], N=256)

BG      = '#F5F7FA'
DARK    = '#1A1A2E'
MUTED   = '#606070'
ACCENT  = '#D32F2F'
BLUE    = '#185FA5'
BLUE_D  = '#0C447C'


# ============================================================
# Main figure: two heatmaps merged into one combined figure
# ============================================================
def plot_combined_main():
    print("Generating main figure: combined heatmap...")

    # ── Prepare Figure 1 data (emergency volume, row-normalised) ──
    data_vol = heatmap_emg.copy()
    data_vol.index = [shorten_label(l) for l in data_vol.index]

    sort_col = '2019-20' if '2019-20' in data_vol.columns else data_vol.columns[-3]
    data_vol = data_vol.sort_values(sort_col, ascending=True)
    cats_vol = list(data_vol.index)
    years    = list(data_vol.columns)

    data_vol_norm = data_vol.div(data_vol.max(axis=1), axis=0)

    # Emergency ratio pivot (used for the right-side bar panel)
    ratio_piv = (by_chapter.pivot_table(
        index='icd_chapter', columns='year',
        values='emergency_ratio', aggfunc='mean')
        .drop(columns='2012-13', errors='ignore'))
    ratio_piv.index = [shorten_label(l) for l in ratio_piv.index]
    ratio_piv = ratio_piv.reindex(cats_vol)
    ratio_bar = ratio_piv[sort_col] if sort_col in ratio_piv.columns else ratio_piv.iloc[:, -3]

    # ── Prepare Figure 2 data (emergency ratio, absolute values) ──
    ratio_heat = ratio_piv.copy()
    # Fix Nervous system 2008-09 anomalous value
    if 'Nervous system' in ratio_heat.index and '2008-09' in ratio_heat.columns:
        ratio_heat.loc['Nervous system', '2008-09'] = (
            ratio_heat.loc['Nervous system', '2007-08'] +
            ratio_heat.loc['Nervous system', '2009-10']) / 2

    # Sort by 2019-20 emergency ratio
    ratio_heat = ratio_heat.sort_values(sort_col, ascending=True)
    cats_ratio = list(ratio_heat.index)

    # ── Canvas layout ──────────────────────────────────────────────
    # Three rows: title strip (top) + Figure 1 panel + Figure 2 panel
    fig = plt.figure(figsize=(24, 22))
    fig.patch.set_facecolor(BG)

    outer_gs = GridSpec(3, 1, figure=fig,
                        height_ratios=[0.04, 1, 1],
                        hspace=0.18)

    # ── Overall title strip ────────────────────────────────────────
    ax_title = fig.add_subplot(outer_gs[0])
    ax_title.set_facecolor(BG)
    ax_title.axis('off')
    ax_title.text(0.0, 0.9,
        'Emergency Hospital Admissions by Diagnostic Category, England (1998–2024)',
        transform=ax_title.transAxes, fontsize=15, fontweight='bold',
        color=DARK, va='top')
    ax_title.text(0.0, 0.0,
        'Data: NHS England HES  |  Tool: Python / Matplotlib  |  All ages · All genders  |  2012-13 excluded (partial year)',
        transform=ax_title.transAxes, fontsize=9, color=MUTED, va='bottom')
    # Blue divider line
    ax_title.axhline(0, color=BLUE, linewidth=2, xmin=0, xmax=1)

    # ── Figure 1 panel: volume heatmap + right-side bar ────────────
    gs_top = GridSpecFromSubplotSpec(
        1, 2, subplot_spec=outer_gs[1],
        width_ratios=[17, 3], wspace=0.02)

    ax_vol  = fig.add_subplot(gs_top[0])
    ax_bar  = fig.add_subplot(gs_top[1])
    ax_vol.set_facecolor(BG)
    ax_bar.set_facecolor(BG)

    n_years  = len(years)
    n_cats_v = len(cats_vol)

    draw_cells(ax_vol, data_vol, data_vol_norm, years, cats_vol,
               CMAP_VOL, show_val=True, threshold=0.45,
               fmt_fn=lambda v: f'{v/1e6:.1f}M')
    add_covid_line(ax_vol, years, n_cats_v, top_label=True)
    style_heatmap_ax(ax_vol, years, cats_vol)

    # Figure 1 title and footnote
    ax_vol.set_title(
        'Figure 1  |  Emergency Admissions Volume (row-normalised)',
        fontsize=11, fontweight='bold', color=BLUE_D, pad=12, loc='left')
    ax_vol.text(0, -0.06,
        "Colour = admissions relative to each category's own peak year. "
        "Values shown where \u2265 45% of peak.",
        transform=ax_vol.transAxes, fontsize=7.5, color=MUTED, va='top')

    # Colourbar for Figure 1
    sm1 = plt.cm.ScalarMappable(cmap=CMAP_VOL, norm=plt.Normalize(0, 1))
    sm1.set_array([])
    cb1 = plt.colorbar(sm1, ax=ax_vol, fraction=0.012, pad=0.01, shrink=0.55)
    cb1.set_ticks([0, 0.5, 1.0])
    cb1.set_ticklabels(['Low', '50%', 'Peak'])
    cb1.ax.tick_params(labelsize=8, colors=MUTED)
    cb1.set_label('Relative to category peak', fontsize=8, color=MUTED)
    cb1.outline.set_visible(False)

    # Right-side emergency share bar panel
    bar_colors = [CMAP_RATIO(float(v)) if not pd.isna(v) else '#CCCCCC'
                  for v in ratio_bar.values]
    y_pos = [i + 0.5 for i in range(n_cats_v)]
    ax_bar.barh(y_pos, ratio_bar.values, height=0.70,
                color=bar_colors, zorder=3)
    for y, v in zip(y_pos, ratio_bar.values):
        if not pd.isna(v):
            ax_bar.text(min(v + 0.01, 1.0), y, f'{v:.0%}',
                        va='center', fontsize=7.5, color=DARK)
    ax_bar.set_xlim(0, 1.18)
    ax_bar.set_ylim(0, n_cats_v)
    ax_bar.set_yticks([])
    ax_bar.set_xticks([0, 0.5, 1.0])
    ax_bar.set_xticklabels(['0%', '50%', '100%'], fontsize=8, color=MUTED)
    ax_bar.axvline(0.5, color='#CCCCCC', linewidth=0.8, linestyle='--')
    ax_bar.set_xlabel('Emergency\nshare (2019-20)', fontsize=8.5, color=MUTED)
    ax_bar.set_title('Emergency\nshare', fontsize=8.5, color=MUTED, pad=4)
    for sp in ax_bar.spines.values():
        sp.set_visible(False)
    ax_bar.tick_params(length=0)
    ax_bar.set_facecolor(BG)

    # ── Figure 2 panel: emergency ratio heatmap ────────────────────
    gs_bot = GridSpecFromSubplotSpec(
        1, 2, subplot_spec=outer_gs[2],
        width_ratios=[17, 3], wspace=0.02)

    ax_ratio = fig.add_subplot(gs_bot[0])
    ax_blank = fig.add_subplot(gs_bot[1])   # Spacer to align width with Figure 1
    ax_ratio.set_facecolor(BG)
    ax_blank.set_facecolor(BG)
    ax_blank.axis('off')

    n_cats_r = len(cats_ratio)

    # Draw ratio cells (absolute values 0–1 mapped directly to colour)
    for yi, year in enumerate(years):
        for ci, cat in enumerate(cats_ratio):
            val = ratio_heat.loc[cat, year]
            if pd.isna(val):
                ax_ratio.add_patch(plt.Rectangle((yi, ci), 1, 1,
                    color='#EBEBEB', linewidth=0.3, edgecolor=BG))
                continue
            color = CMAP_RATIO(float(val))
            ax_ratio.add_patch(plt.Rectangle((yi, ci), 1, 1,
                color=color, linewidth=0.3, edgecolor=BG))
            lum = 0.299*color[0] + 0.587*color[1] + 0.114*color[2]
            tc  = '#111111' if lum > 0.45 else '#FFFFFF'
            ax_ratio.text(yi + 0.5, ci + 0.5, f'{val:.0%}',
                ha='center', va='center',
                fontsize=5.8, color=tc, fontweight='500')

    add_covid_line(ax_ratio, years, n_cats_r, top_label=True)
    style_heatmap_ax(ax_ratio, years, cats_ratio)

    # Annotate key categories with arrow callouts on the right
    annotations = {
        'Infectious & parasitic': ('> 80% emergency throughout',  '#084594'),
        'Mental & behavioural':   ('+22pp rise 1998\u21922024',   '#1565C0'),
        'Neoplasms':              ('< 15% \u2014 mostly planned', '#880000'),
    }
    for cat_name, (note, acolor) in annotations.items():
        matches = [i for i, c in enumerate(cats_ratio) if cat_name in c]
        if matches:
            ci = matches[0]
            ax_ratio.annotate(note,
                xy=(n_years + 0.2, ci + 0.5),
                xytext=(n_years + 0.4, ci + 0.5),
                fontsize=7.5, color=acolor, va='center',
                arrowprops=dict(arrowstyle='->', color=acolor, lw=1.2))

    ax_ratio.set_xlim(0, n_years + 3.8)
    ax_ratio.set_title(
        'Figure 2  |  Emergency Admission Ratio (absolute values)',
        fontsize=11, fontweight='bold', color=BLUE_D, pad=12, loc='left')
    ax_ratio.text(0, -0.06,
        'Colour depth = proportion of total admissions classified as emergency. '
        'Darker blue = higher emergency burden.',
        transform=ax_ratio.transAxes, fontsize=7.5, color=MUTED, va='top')

    # Colourbar for Figure 2
    sm2 = plt.cm.ScalarMappable(cmap=CMAP_RATIO, norm=plt.Normalize(0, 1))
    sm2.set_array([])
    cb2 = plt.colorbar(sm2, ax=ax_ratio, fraction=0.012, pad=0.01, shrink=0.55)
    cb2.set_ticks([0, 0.25, 0.5, 0.75, 1.0])
    cb2.set_ticklabels(['0%', '25%', '50%', '75%', '100%'])
    cb2.ax.tick_params(labelsize=8, colors=MUTED)
    cb2.set_label('Emergency / total admissions', fontsize=8, color=MUTED)
    cb2.outline.set_visible(False)

    # ── Save ───────────────────────────────────────────────────────
    out = f"{DATA_FOLDER}nhs_main_combined.png"
    plt.savefig(out, dpi=180, bbox_inches='tight',
                facecolor=BG, edgecolor='none')
    plt.close()
    print(f"  Saved: {out}")


# ============================================================
# Supplementary figure 1: 26-year emergency admissions trend
# ============================================================
def plot_trend():
    print("Generating supplementary figure 1: long-term trend...")

    yr = (by_chapter.groupby(['year_start', 'year'])['emergency']
          .sum().reset_index().sort_values('year_start'))
    yr = yr[yr['year'] != '2012-13']

    fig, ax = plt.subplots(figsize=(14, 5))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)

    x    = range(len(yr))
    vals = yr['emergency'].values / 1e6
    yrs  = yr['year'].values

    ax.fill_between(x, vals, alpha=0.15, color='#1D75B3')
    ax.plot(x, vals, color='#1D75B3', linewidth=2.5, zorder=3)

    # Highlight the COVID-19 lockdown year
    if '2020-21' in list(yrs):
        ci  = list(yrs).index('2020-21')
        pct = (vals[ci] - vals[ci-1]) / vals[ci-1] * 100
        ax.scatter([ci], [vals[ci]], color=ACCENT, s=80, zorder=5)
        ax.annotate(f'COVID-19 lockdown\n{vals[ci]:.1f}M  ({pct:+.1f}%)',
                    xy=(ci, vals[ci]), xytext=(ci - 2.5, vals[ci] - 1.6),
                    fontsize=8.5, color=ACCENT,
                    arrowprops=dict(arrowstyle='->', color=ACCENT, lw=1.5))

    ax.set_xticks(list(x))
    ax.set_xticklabels(yrs, rotation=45, ha='right', fontsize=8, color=MUTED)
    ax.set_ylabel('Emergency Admissions (Millions)', fontsize=9, color=MUTED)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'{v:.0f}M'))
    ax.set_title('Total Emergency Hospital Admissions, England (1998–2024)',
                 fontsize=12, fontweight='bold', color=DARK, loc='left')
    ax.grid(axis='y', color='#D0D8E0', linewidth=0.7, linestyle='--', alpha=0.7)
    ax.set_axisbelow(True)
    for sp in ['top', 'right']:
        ax.spines[sp].set_visible(False)
    ax.spines['left'].set_color('#D0D8E0')
    ax.spines['bottom'].set_color('#D0D8E0')
    ax.tick_params(colors=MUTED, length=0)

    plt.tight_layout()
    out = f"{DATA_FOLDER}nhs_trend.png"
    plt.savefig(out, dpi=150, bbox_inches='tight',
                facecolor=BG, edgecolor='none')
    plt.close()
    print(f"  Saved: {out}")


# ============================================================
# Supplementary figure 2: COVID-19 impact by category
# ============================================================
def plot_covid_change():
    print("Generating supplementary figure 2: COVID-19 impact bar chart...")

    pre    = covid[covid['year'] == '2019-20'].set_index('icd_chapter')['emergency']
    post   = covid[covid['year'] == '2020-21'].set_index('icd_chapter')['emergency']
    change = ((post - pre) / pre * 100).dropna().sort_values()
    change.index = [shorten_label(l) for l in change.index]

    fig, ax = plt.subplots(figsize=(10, 7))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)

    bar_colors = [ACCENT if v < -20 else '#E67E22' if v < 0 else '#27AE60'
                  for v in change.values]
    ax.barh(range(len(change)), change.values,
            color=bar_colors, height=0.65, zorder=3)

    for i, v in enumerate(change.values):
        xpos = v - 1.2 if v < 0 else v + 0.3
        ax.text(xpos, i, f'{v:+.1f}%',
                ha='right' if v < 0 else 'left', va='center',
                fontsize=8.5, color=DARK, fontweight='500')

    ax.axvline(0, color=DARK, linewidth=1, zorder=4)
    ax.set_yticks(range(len(change)))
    ax.set_yticklabels(change.index, fontsize=9.5, color=DARK)
    ax.set_xlabel('Change in Emergency Admissions (%)', fontsize=9, color=MUTED)
    ax.set_title(
        'Change in Emergency Admissions: 2020-21 vs 2019-20\n(Impact of COVID-19 Lockdown)',
        fontsize=12, fontweight='bold', color=DARK, loc='left')
    ax.legend(
        handles=[mpatches.Patch(color=ACCENT,    label='Severe drop (> 20%)'),
                 mpatches.Patch(color='#E67E22', label='Moderate drop')],
        fontsize=8.5, loc='lower right', framealpha=0.9, edgecolor='#D0D8E0')
    ax.grid(axis='x', color='#D0D8E0', linewidth=0.7, linestyle='--', alpha=0.7)
    ax.set_axisbelow(True)
    for sp in ['top', 'right']:
        ax.spines[sp].set_visible(False)
    ax.spines['left'].set_color('#D0D8E0')
    ax.spines['bottom'].set_color('#D0D8E0')
    ax.tick_params(colors=MUTED, length=0)

    plt.tight_layout()
    out = f"{DATA_FOLDER}nhs_covid_change.png"
    plt.savefig(out, dpi=150, bbox_inches='tight',
                facecolor=BG, edgecolor='none')
    plt.close()
    print(f"  Saved: {out}")


# ============================================================
# Entry point
# ============================================================
if __name__ == '__main__':
    print("=" * 60)
    print("NHS Data Visualisation v4")
    print("=" * 60)

    plot_combined_main()   # Main figure (for submission)
    plot_trend()           # Supplementary figure 1
    plot_covid_change()    # Supplementary figure 2

    print()
    print("=" * 60)
    print("All done! The following files have been generated:")
    print()
    print("  nhs_main_combined.png   <- Main figure (submit this)")
    print("    Figure 1: Emergency volume heatmap (row-normalised) + share bar")
    print("    Figure 2: Emergency ratio heatmap (absolute values)")
    print()
    print("  nhs_trend.png           <- Supplementary figure 1: 26-year trend")
    print("  nhs_covid_change.png    <- Supplementary figure 2: COVID impact")
    print("=" * 60)
