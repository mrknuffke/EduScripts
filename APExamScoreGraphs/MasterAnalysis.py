import sys
import os
from datetime import date

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

import seaborn as sns
from matplotlib.colors import Normalize
from matplotlib.backends.backend_pdf import PdfPages
from scipy import stats
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score

# ──────────────────────────────────────────────────────────────────────────────
# Tufte-Inspired Style — maximize data-ink, minimize chartjunk
# ──────────────────────────────────────────────────────────────────────────────
STYLE = {
    'figure.figsize': (10, 6),
    'figure.facecolor': 'white',
    'axes.facecolor': 'white',
    'axes.titlesize': 12,
    'axes.titleweight': 'normal',
    'axes.labelsize': 10,
    'axes.labelcolor': '#444444',
    'xtick.labelsize': 9,
    'ytick.labelsize': 9,
    'xtick.color': '#666666',
    'ytick.color': '#666666',
    'axes.grid': False,
    'axes.spines.top': False,
    'axes.spines.right': False,
    'axes.spines.left': True,
    'axes.spines.bottom': True,
    'axes.edgecolor': '#999999',
    'axes.linewidth': 0.6,
    'font.family': 'sans-serif',
    'text.color': '#333333',
}
plt.rcParams.update(STYLE)

# Muted, restrained palette — warm greys with consistent semantic colours
PALETTE_ACCENT = '#2c7bb6'       # blue — test averages, primary data, "passing"
PALETTE_SECONDARY = '#c0392b'    # red  — AP scores, secondary series, "below threshold"
PALETTE_NEUTRAL = '#888888'      # grey — regression lines, reference marks
PALETTE_MUTED = ['#6e8898', '#8aaa9e', '#a3b5a5', '#c4cdb4', '#ddd5c7',
                 '#bfb1a4', '#9e8e82', '#7d6d62']  # warm muted tones
N_FONT = {'fontsize': 8, 'color': '#777777'}
CAPTION_FONT = {'fontsize': 9, 'style': 'italic', 'color': '#555555',
                'wrap': True, 'ha': 'left', 'va': 'top'}

# ──────────────────────────────────────────────────────────────────────────────
# Grade Scale (10-point banding)
# ──────────────────────────────────────────────────────────────────────────────
GRADE_THRESHOLDS = [
    ('A+', 90),
    ('A',  80),
    ('B+', 70),
    ('B',  60),
    ('C+', 50),
    ('C',  40),
]
GRADE_ORDER = ['D', 'D+', 'C', 'C+', 'B', 'B+', 'A', 'A+']


def expected_grade(test_avg):
    """Returns the expected letter grade for a given test average.

    Boundaries: >90 = A+, >80 = A, >70 = B+, >60 = B, >50 = C+, >40 = C.
    A score of exactly 90.0 is an A (not A+), exactly 80.0 is a B+ (not A), etc.
    """
    for grade, threshold in GRADE_THRESHOLDS:
        if test_avg > threshold:
            return grade
    return 'Below C'


# ──────────────────────────────────────────────────────────────────────────────
# Interactive Validation
# ──────────────────────────────────────────────────────────────────────────────
def validate_grade_assignments(df, script_dir, csv_path, raw_df):
    """Flags grade/test-average mismatches and lets the user decide how to proceed."""
    print("── Validating Grade Assignments ──")
    flagged = []
    for idx, row in df.iterrows():
        exp = expected_grade(row['Test Average'])
        if row['Test Average Grade'] != exp:
            flagged.append({
                'Index': idx,
                'Student': row['Student'],
                'Year': row['Schoolyear'],
                'Block': row['Block'],
                'Test Avg': row['Test Average'],
                'Current Grade': row['Test Average Grade'],
                'Expected Grade': exp,
            })

    if not flagged:
        print("   No grade mismatches found.\n")
        return df

    # Save flagged rows
    flagged_df = pd.DataFrame(flagged)
    flagged_path = os.path.join(script_dir, 'flagged_for_review.csv')
    flagged_df.to_csv(flagged_path, index=False)

    print(f"\n   Found {len(flagged)} grade mismatch(es).")
    print("   Options:")
    print("     [review] - Review each mismatch one at a time")
    print("     [fix]    - Auto-correct all grades to match the 10-point scale")
    print("     [Enter]  - Continue with data as-is")
    print(f"     [quit]   - Exit so you can manually edit {os.path.basename(csv_path)}")
    try:
        if not sys.stdin.isatty():
            choice = ''
        else:
            choice = input("\n   > ").strip().lower()
    except EOFError:
        choice = ''

    if choice == 'quit':
        print(f"   Exiting. Edit {os.path.basename(csv_path)} and re-run.")
        sys.exit(0)
    # Track all edits: {original_index: new_grade_value}
    grade_edits = {}
    drop_indices = []
    changed = False

    if choice == 'fix':
        for f in flagged:
            grade_edits[f['Index']] = f['Expected Grade']
        print(f"   Auto-corrected {len(flagged)} grade(s).")
        changed = True
    elif choice == 'review':
        corrected = 0
        skipped = 0
        removed = 0
        for i, f in enumerate(flagged, 1):
            print(f"\n   [{i}/{len(flagged)}] {f['Student']} ({f['Year']}, {f['Block']})")
            print(f"           Test Average: {f['Test Avg']}")
            print(f"           Current Grade: {f['Current Grade']}  →  Expected: {f['Expected Grade']}")
            print(f"             [Enter] Accept correction to '{f['Expected Grade']}'")
            print(f"             [s]     Skip (keep '{f['Current Grade']}')")
            print(f"             [d]     Remove this student from the dataset")
            print(f"             [TYPE]  Enter a custom grade (e.g. B+)")
            try:
                resp = input("           > ").strip()
            except EOFError:
                resp = 's'
            if resp == '':
                grade_edits[f['Index']] = f['Expected Grade']
                print(f"           → Changed to '{f['Expected Grade']}'")
                corrected += 1
            elif resp.lower() == 's':
                print(f"           → Kept '{f['Current Grade']}'")
                skipped += 1
            elif resp.lower() == 'd':
                drop_indices.append(f['Index'])
                print(f"           → Marked for removal")
                removed += 1
            else:
                valid_grades = [g for g, _ in GRADE_THRESHOLDS] + ['Below C']
                if resp in valid_grades:
                    grade_edits[f['Index']] = resp
                    print(f"           → Changed to '{resp}'")
                    corrected += 1
                else:
                    print(f"           → '{resp}' not recognized, keeping '{f['Current Grade']}'")
                    skipped += 1
        parts = []
        if corrected: parts.append(f"{corrected} corrected")
        if skipped: parts.append(f"{skipped} kept as-is")
        if removed: parts.append(f"{removed} removed")
        print(f"\n   Review complete: {', '.join(parts)}.")
        changed = (corrected > 0 or removed > 0)
    else:
        print("   Continuing with original data.\n")

    # Apply edits to the working dataframe
    for idx, new_grade in grade_edits.items():
        df.at[idx, 'Test Average Grade'] = new_grade
    if drop_indices:
        df = df.drop(index=drop_indices).reset_index(drop=True)

    # Offer to save changes back to the CSV
    if changed:
        try:
            save = input(f"\n   Save changes to {os.path.basename(csv_path)}? [y/N] > ").strip().lower()
        except EOFError:
            save = 'n'
        if save == 'y':
            updated = raw_df.copy()
            for idx, new_grade in grade_edits.items():
                updated.at[idx, 'Test Average Grade'] = new_grade
            if drop_indices:
                updated = updated.drop(index=drop_indices)
            updated.to_csv(csv_path, index=False)
            print(f"   → Saved to {csv_path}\n")
        else:
            print("   → Changes applied for this run only (CSV not modified).\n")

    return df


# ──────────────────────────────────────────────────────────────────────────────
# Figure 1 — Overall Correlation
# ──────────────────────────────────────────────────────────────────────────────
def fig_correlation(df):
    freq = df.groupby(['Test Average', 'AP Exam Score']).size().reset_index(name='Freq')
    corr = df[['Test Average', 'AP Exam Score']].corr()
    r = corr.iloc[0, 1]
    r2 = r ** 2

    fig, ax = plt.subplots(figsize=(10, 6))
    
    # Scatter points — translucent blue bubbles scaled by frequency
    sns.scatterplot(data=freq, x='Test Average', y='AP Exam Score',
                    size='Freq', sizes=(40, 350), color='#4682B4',
                    alpha=0.55, legend=False, ax=ax, zorder=2)
    
    # Regression line with 95% Confidence Interval band
    sns.regplot(data=df, x='Test Average', y='AP Exam Score', scatter=False,
                color='red', ci=95,
                line_kws={'linewidth': 1.8, 'zorder': 3},
                scatter_kws={'alpha': 0.15},
                ax=ax)

    ax.set_title('Test Average - AP Exam Correlation', fontsize=12, pad=12)
    ax.set_xlabel('In-Class Assessment Average', fontsize=10, labelpad=8)
    ax.set_ylabel('AP Exam Score', fontsize=10, labelpad=8)
    ax.set_xlim(38, 98)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_ylim(0.8, 5.3)

    # Rounded box for Pearson's r and R-squared
    bbox_props = dict(boxstyle='round,pad=0.6', facecolor='#FDF6E3', edgecolor='#B58900', alpha=0.85, linewidth=0.8)
    ax.text(0.04, 0.94, f"Pearson's r = {r:.3f}\nR-squared = {r2:.3f}",
            transform=ax.transAxes, fontsize=11, va='top', color='#333333', bbox=bbox_props, zorder=5)

    ax.grid(True, linestyle='--', color='#d3d3d3', alpha=0.7, zorder=1)

    for spine in ax.spines.values():
        spine.set_linewidth(0.6)
        spine.set_color('#333333')

    fig.tight_layout()
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# Figure 2 — Grade vs AP Score Frequency Bubble Chart
# ──────────────────────────────────────────────────────────────────────────────
def fig_heatmap(df):
    df = df.copy()
    
    grade_order = ['D', 'D+', 'C', 'C+', 'B', 'B+', 'A', 'A+']
    df['Test Average Grade'] = pd.Categorical(df['Test Average Grade'],
                                               categories=grade_order, ordered=True)
    
    # Frequency per (Grade, AP Exam Score) cell
    freq = df.groupby(['Test Average Grade', 'AP Exam Score'], observed=False).size().reset_index(name='Count')
    freq_active = freq[freq['Count'] > 0].copy()
    
    grade_map = {g: i for i, g in enumerate(grade_order)}
    freq_active['x_pos'] = freq_active['Test Average Grade'].map(grade_map)
    
    fig, ax = plt.subplots(figsize=(10, 6))
    
    # YlGnBu colormap matching original bubble chart
    cmap = plt.cm.YlGnBu
    max_count = freq_active['Count'].max()
    norm = Normalize(vmin=1, vmax=max_count)
    
    for _, row in freq_active.iterrows():
        x = row['x_pos']
        y = row['AP Exam Score']
        cnt = row['Count']
        
        # Scaling bubble size (area in pt^2)
        size = 200 + (cnt / max_count) * 2600
        color = cmap(norm(cnt))
        
        # Draw bubble
        ax.scatter(x, y, s=size, color=color, edgecolors='black', linewidth=1.1, zorder=3)
        
        # Text color based on background darkness
        text_color = 'white' if norm(cnt) > 0.55 else 'black'
        
        # Exact count label inside bubble
        ax.annotate(str(cnt), (x, y), ha='center', va='center',
                    fontsize=10, fontweight='bold', color=text_color, zorder=4)
        
    ax.set_xticks(range(len(grade_order)))
    ax.set_xticklabels(grade_order, fontsize=9)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels([1, 2, 3, 4, 5], fontsize=9)
    ax.set_xlim(-0.5, len(grade_order) - 0.5)
    ax.set_ylim(0.5, 5.5)
    
    ax.set_title('AP Exam Score Frequency by Average In-Class Assessment Letter Grade', fontsize=12, pad=12)
    ax.set_xlabel('Average In-Class Assessment Letter Grade', fontsize=10, labelpad=8)
    ax.set_ylabel('AP Exam Score', fontsize=10, labelpad=8)
    
    ax.grid(True, linestyle=':', alpha=0.5, color='#cccccc', zorder=1)
    
    for spine in ax.spines.values():
        spine.set_linewidth(0.6)
        spine.set_color('#333333')
        
    fig.tight_layout()
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# Figure 3 — Descriptive Stats Table
# ──────────────────────────────────────────────────────────────────────────────
def fig_stats_table(df):
    tbl = df.groupby('AP Exam Score')['Test Average'].agg(
        N='size', Mean='mean', Median='median',
        Range=lambda x: np.ptp(x) if len(x) > 0 else 0,
        SD='std'
    ).fillna(0).round(2)
    tbl['N'] = tbl['N'].astype(int)

    # Format for display so N renders as integer strings
    display = tbl.copy()
    display['N'] = display['N'].astype(str)

    fig, ax = plt.subplots(figsize=(8, 3))
    ax.axis('off')
    table = ax.table(cellText=display.values, colLabels=display.columns,
                     rowLabels=tbl.index, cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.4)
    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor('#dddddd')
        cell.set_linewidth(0.5)
        if r == 0:
            cell.set_facecolor('white')
            cell.set_text_props(color='#333333', fontweight='bold')
        else:
            cell.set_facecolor('white')
        if c == -1:
            cell.set_text_props(fontweight='bold', color='#444444')
    fig.suptitle('Descriptive statistics of test averages by AP score',
                 fontsize=12, color='#333333')
    fig.tight_layout(rect=[0, 0, 1, 0.92])
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# Figure 4 — Probability Heatmap Table
# ──────────────────────────────────────────────────────────────────────────────
def fig_probability_table(df):
    df = df.copy()
    bins = [0, 40, 50, 60, 70, 80, 90, 101]
    labels = ['Below 40', '40–49 (C)', '50–59 (C+)', '60–69 (B)', '70–79 (B+)', '80–89 (A)', '90–100 (A+)']
    df['Range'] = pd.cut(df['Test Average'], bins=bins, labels=labels,
                         right=False, include_lowest=True)

    ct = pd.crosstab(df['Range'], df['AP Exam Score'])
    probs = ct.div(ct.sum(axis=1), axis=0).fillna(0)
    summary = pd.concat([ct.sum(axis=1), probs], axis=1)
    summary.columns = ['N'] + [f'P(Score={i})' for i in probs.columns]
    for s in range(1, 6):
        col = f'P(Score={s})'
        if col not in summary.columns:
            summary[col] = 0.0
    summary = summary[['N'] + [f'P(Score={i})' for i in range(1, 6)]]

    display = summary.copy()
    for col in display.columns[1:]:
        display[col] = display[col].apply(lambda x: f'{x:.0%}')
    display['N'] = display['N'].astype(int)

    color_vals = summary.iloc[:, 1:].values
    cell_colors = plt.get_cmap('Blues')(Normalize(vmin=0, vmax=1)(color_vals))
    n_col = np.full((cell_colors.shape[0], 1, 4), [1, 1, 1, 1])
    all_colors = np.hstack([n_col, cell_colors])

    fig, ax = plt.subplots(figsize=(10, 4))
    ax.axis('off')
    table = ax.table(cellText=display.values, colLabels=summary.columns,
                     rowLabels=summary.index, cellColours=all_colors,
                     cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.5)
    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor('#dddddd')
        cell.set_linewidth(0.5)
        if r == 0:
            cell.set_facecolor('white')
            cell.set_text_props(color='#333333', fontweight='bold')
        elif c > 0:
            # Check probability value to adjust text color for contrast
            # summary index is r-1 (because r=0 is header)
            # summary column is c (0 is N, 1-5 are probs)
            try:
                # Get the raw probability value from the summary dataframe
                val = summary.iloc[r-1, c]
                if val > 0.7:  # Dark blue background
                    cell.get_text().set_color('white')
                else:
                    cell.get_text().set_color('#333333')
            except (IndexError, KeyError):
                pass
        
        if c == -1:
            cell.set_text_props(fontweight='bold', fontsize=8, color='#444444')
    fig.suptitle('Probability of AP score by test average range',
                 fontsize=12, color='#333333')
    fig.tight_layout(rect=[0, 0, 1, 0.92])
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# Figure 5 — Yearly Trends with 95% CI
# ──────────────────────────────────────────────────────────────────────────────
def fig_yearly_trends(df):
    ys = df.groupby('Schoolyear').agg(
        N=('Test Average', 'size'),
        MTA=('Test Average', 'mean'), SEM_TA=('Test Average', 'sem'),
        MAP=('AP Exam Score', 'mean'), SEM_AP=('AP Exam Score', 'sem'),
    ).reset_index()
    t_crit = stats.t.ppf(0.975, ys['N'] - 1)
    ys['CI_TA'] = t_crit * ys['SEM_TA']
    ys['CI_AP'] = t_crit * ys['SEM_AP']

    # ANOVA
    groups = [g['Test Average'].values for _, g in df.groupby('Schoolyear')]
    f_stat, p_val = stats.f_oneway(*groups)
    anova_text = f"ANOVA: F = {f_stat:.2f}, p = {p_val:.4f}"
    if p_val < 0.05:
        anova_text += " (significant)"
    else:
        anova_text += " (not significant)"

    # Stacked panels sharing x-axis (avoids dual-axis distortion)
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 7), sharex=True,
                                    gridspec_kw={'height_ratios': [1, 1],
                                                 'hspace': 0.08})
    x_idxs = np.arange(len(ys['Schoolyear']))

    # Top panel — Mean Test Average
    ax1.errorbar(x_idxs, ys['MTA'], yerr=ys['CI_TA'],
                 color=PALETTE_ACCENT, marker='o', markersize=5,
                 capsize=3, linewidth=1.5, alpha=0.9)
    ax1.set_ylabel('Mean test average')
    # Tight y-limits around the data range
    ta_min = (ys['MTA'] - ys['CI_TA']).min()
    ta_max = (ys['MTA'] + ys['CI_TA']).max()
    ta_pad = (ta_max - ta_min) * 0.3
    ax1.set_ylim(ta_min - ta_pad, ta_max + ta_pad)
    for i, row in ys.iterrows():
        ax1.annotate(f"n={row['N']}", (i, row['MTA']),
                     textcoords='offset points', xytext=(0, 12), ha='center', **N_FONT)
    ax1.set_title('Year-over-year performance trends (95% CI)')
    ax1.spines['bottom'].set_visible(False)
    ax1.tick_params(axis='x', length=0)

    # Bottom panel — Mean AP Score
    ax2.errorbar(x_idxs, ys['MAP'], yerr=ys['CI_AP'],
                 color=PALETTE_SECONDARY, marker='s', markersize=5,
                 capsize=3, linewidth=1.5, alpha=0.9)
    ax2.set_ylabel('Mean AP score')
    ap_min = (ys['MAP'] - ys['CI_AP']).min()
    ap_max = (ys['MAP'] + ys['CI_AP']).max()
    ap_pad = (ap_max - ap_min) * 0.3
    ax2.set_ylim(ap_min - ap_pad, ap_max + ap_pad)
    ax2.set_xticks(x_idxs)
    ax2.set_xticklabels(ys['Schoolyear'], rotation=45)

    fig.subplots_adjust(bottom=0.18)
    fig.text(0.5, 0.02, anova_text, ha='center', fontsize=9,
             style='italic', color='#777777')

    # Compute Tukey HSD if significant
    tukey_rows = []
    print(f"\n   {anova_text}")
    if p_val < 0.05:
        tukey = stats.tukey_hsd(*groups)
        years = sorted(df['Schoolyear'].unique())
        print("   Tukey HSD Post-Hoc:")
        for i in range(len(years)):
            for j in range(i + 1, len(years)):
                p = tukey.pvalue[i][j]
                sig = " *" if p < 0.05 else ""
                tukey_rows.append((years[i], years[j], p, p < 0.05))
                print(f"     {years[i]} vs {years[j]}: p = {p:.4f}{sig}")

    return fig, tukey_rows


# ──────────────────────────────────────────────────────────────────────────────
# Figure 6 — Box Plot
# ──────────────────────────────────────────────────────────────────────────────
def fig_boxplot(df):
    fig, ax = plt.subplots(figsize=(10, 6))
    order = sorted(df['AP Exam Score'].unique())

    # Tufte-style boxes: light fill, thin edges, no heavy colour
    box_props = dict(facecolor='#f0f0f0', edgecolor='#888888', linewidth=0.8)
    median_props = dict(color='#333333', linewidth=1.2)
    whisker_props = dict(color='#888888', linewidth=0.8)
    cap_props = dict(color='#888888', linewidth=0.8)
    flier_props = dict(marker='', linewidth=0)  # hide default outliers; stripplot shows data

    sns.boxplot(x='AP Exam Score', y='Test Average', hue='AP Exam Score',
                data=df, order=order, legend=False, ax=ax, width=0.35,
                boxprops=box_props, medianprops=median_props,
                whiskerprops=whisker_props, capprops=cap_props,
                flierprops=flier_props)

    # jitter=False → dots sit directly at their x-position, aligned with outliers
    sns.stripplot(x='AP Exam Score', y='Test Average', data=df,
                  order=order, color=PALETTE_ACCENT, alpha=0.4, size=3.5,
                  jitter=False, ax=ax)

    counts = df.groupby('AP Exam Score').size()
    ymax = df['Test Average'].max() + 4
    for i, score in enumerate(order):
        if score in counts.index:
            ax.text(i, ymax, f'n={counts[score]}', ha='center', **N_FONT)

    ax.set_title('Distribution of test averages by AP exam score')
    ax.set_xlabel('AP exam score')
    ax.set_ylabel('In-class test average')
    ax.set_ylim(None, ymax + 5)
    fig.tight_layout()
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# Figure 7 — "3 or Above" Rate (side-by-side)
# ──────────────────────────────────────────────────────────────────────────────
def fig_three_plus(df):
    df = df.copy()
    df['Three_Plus'] = (df['AP Exam Score'] >= 3).astype(int)

    by_grade = df.groupby('Test Average Grade').agg(
        Rate=('Three_Plus', 'mean'), N=('Three_Plus', 'size')
    ).reindex(GRADE_ORDER).dropna()
    by_grade['Rate'] *= 100

    by_year = df.groupby('Schoolyear').agg(
        Rate=('Three_Plus', 'mean'), N=('Three_Plus', 'size')
    )
    by_year['Rate'] *= 100

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5.5))

    # By grade — thin horizontal bars
    y_pos_g = range(len(by_grade))
    ax1.barh(list(y_pos_g), by_grade['Rate'], height=0.4,
             color=PALETTE_ACCENT, alpha=0.7, edgecolor='none')
    for i, (idx, row) in enumerate(by_grade.iterrows()):
        ax1.text(row['Rate'] + 2, i, f"{row['Rate']:.0f}%  (n={int(row['N'])})",
                 va='center', fontsize=8, color='#555555')
    ax1.set_yticks(list(y_pos_g))
    ax1.set_yticklabels(by_grade.index)
    ax1.set_title('By letter grade')
    ax1.set_xlabel('% scoring 3 or above')
    ax1.set_xlim(0, 115)
    ax1.invert_yaxis()

    # By year — thin horizontal bars
    y_pos_y = range(len(by_year))
    ax2.barh(list(y_pos_y), by_year['Rate'], height=0.4,
             color=PALETTE_ACCENT, alpha=0.7, edgecolor='none')
    for i, (idx, row) in enumerate(by_year.iterrows()):
        ax2.text(row['Rate'] + 2, i, f"{row['Rate']:.0f}%  (n={int(row['N'])})",
                 va='center', fontsize=8, color='#555555')
    ax2.set_yticks(list(y_pos_y))
    ax2.set_yticklabels(by_year.index)
    ax2.set_title('By school year')
    ax2.set_xlabel('% scoring 3 or above')
    ax2.set_xlim(0, 115)
    ax2.invert_yaxis()

    fig.suptitle('Rate of students scoring 3 or above on AP exam',
                 fontsize=12, color='#333333')
    fig.tight_layout(rect=[0, 0, 1, 0.93])
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# Figure 8 — Multi-Threshold Logistic Regression
# ──────────────────────────────────────────────────────────────────────────────
# Score targets and the probability levels at which we want thresholds
SCORE_TARGETS = [
    ('3+', lambda s: s >= 3),
    ('4+', lambda s: s >= 4),
    ('5',  lambda s: s == 5),
]
PROB_LEVELS = [0.50, 0.75, 0.90]
TARGET_COLORS = {
    '3+': PALETTE_ACCENT,       # blue
    '4+': '#e6a817',            # amber
    '5':  PALETTE_SECONDARY,    # red
}
TARGET_STYLES = {
    '3+': '-',
    '4+': '--',
    '5':  ':',
}


def _find_threshold(xr_vals, probs, level):
    """Return the test average where predicted probability first reaches `level`,
    or None if the curve never reaches it within the data range."""
    idx = np.where(probs >= level)[0]
    if len(idx) == 0:
        return None
    return xr_vals[idx[0]]


def fig_logistic(df):
    df = df.copy()
    X = df[['Test Average']]
    xr = pd.DataFrame(np.linspace(40, 100, 500), columns=['Test Average'])

    results = {}   # {target_label: {prob_level: threshold_value, 'acc': ..., 'probs': ...}}

    fig, ax = plt.subplots(figsize=(10, 6))

    for label, fn in SCORE_TARGETS:
        y = fn(df['AP Exam Score']).astype(int)

        # Skip if fewer than 5 positive cases (model unreliable)
        if y.sum() < 5:
            print(f"   [{label}] Skipped — only {y.sum()} positive cases")
            continue

        # Stratify when the minority class has enough samples for both splits;
        # fall back to unstratified when counts are too low.
        min_class = min(y.sum(), len(y) - y.sum())
        use_stratify = y if min_class >= 3 else None
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=0.2, random_state=42, stratify=use_stratify)
        model = LogisticRegression(random_state=42)
        model.fit(X_train, y_train)

        acc = accuracy_score(y_test, model.predict(X_test))
        probs = model.predict_proba(xr)[:, 1]

        results[label] = {'acc': acc, 'probs': probs}
        print(f"   [{label}] Accuracy: {acc:.1%}")

        # Plot curve
        ax.plot(xr.values, probs, color=TARGET_COLORS[label],
                linewidth=1.5, linestyle=TARGET_STYLES[label], label=f'P(score {label})')

        # Find and mark thresholds
        for level in PROB_LEVELS:
            t = _find_threshold(xr.values.flatten(), probs, level)
            if t is not None:
                results[label][level] = t
                # Small tick mark on the curve
                ax.plot(t, level, marker='|', color=TARGET_COLORS[label],
                        markersize=8, markeredgewidth=1.5, zorder=6)
                print(f"         {level:.0%} threshold: test avg = {t:.1f}")
            else:
                print(f"         {level:.0%} threshold: not reached in data range")

    # Horizontal reference lines at probability levels
    for level in PROB_LEVELS:
        ax.axhline(level, color='#dddddd', linestyle=':', linewidth=0.6)
        ax.text(100.5, level, f'{level:.0%}', fontsize=7, va='center', color='#999999')

    ax.set_title('Predicted probability of scoring at each level')
    ax.set_xlabel('In-class test average')
    ax.set_ylabel('Predicted probability')
    ax.set_xlim(40, 100)
    ax.set_ylim(-0.02, 1.02)
    ax.legend(loc='lower right', fontsize=9, frameon=False)
    fig.tight_layout()
    return fig, results


# ──────────────────────────────────────────────────────────────────────────────
# Logistic Regression — Predicted Probabilities Table
# ──────────────────────────────────────────────────────────────────────────────
def fig_logistic_table(logistic_results):
    """Renders a table of predicted probabilities at representative test averages."""
    sample_avgs = list(range(45, 100, 5))  # 45, 50, 55, ... 95
    xr = np.linspace(40, 100, 500)

    # Build table data: rows = test averages, cols = each target's probability
    targets = [lbl for lbl in logistic_results if 'probs' in logistic_results[lbl]]
    col_labels = ['Test Avg'] + [f'P(score {t})' for t in targets]
    rows = []
    for avg in sample_avgs:
        idx = np.argmin(np.abs(xr - avg))
        row = [str(avg)]
        for t in targets:
            p = logistic_results[t]['probs'][idx]
            row.append(f'{p:.0%}')
        rows.append(row)

    # Color-map the probability cells
    prob_vals = np.array([
        [logistic_results[t]['probs'][np.argmin(np.abs(xr - avg))]
         for t in targets]
        for avg in sample_avgs
    ])
    cell_colors = plt.get_cmap('Blues')(Normalize(vmin=0, vmax=1)(prob_vals))
    avg_col = np.full((cell_colors.shape[0], 1, 4), [1, 1, 1, 1])
    all_colors = np.hstack([avg_col, cell_colors])

    fig, ax = plt.subplots(figsize=(10, max(4, len(sample_avgs) * 0.4 + 2)))
    ax.axis('off')
    table = ax.table(cellText=rows, colLabels=col_labels,
                     cellColours=all_colors, cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.4)
    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor('#dddddd')
        cell.set_linewidth(0.5)
        if r == 0:
            cell.set_facecolor('white')
            cell.set_text_props(color='#333333', fontweight='bold')
        elif c > 0:
            try:
                val = prob_vals[r - 1, c - 1]
                if val > 0.7:
                    cell.get_text().set_color('white')
                else:
                    cell.get_text().set_color('#333333')
            except (IndexError, KeyError):
                pass
    fig.suptitle('Predicted probability of scoring at each level by test average',
                 fontsize=12, color='#333333')
    fig.tight_layout(rect=[0, 0, 1, 0.93])
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# Section Comparison (table page for PDF)
# ──────────────────────────────────────────────────────────────────────────────
def fig_section_table(df):
    df = df.copy()
    df['Three_Plus'] = (df['AP Exam Score'] >= 3).astype(int)
    ss = df.groupby(['Schoolyear', 'Block']).agg(
        N=('Test Average', 'size'),
        Mean_TA=('Test Average', lambda x: round(x.mean(), 1)),
        Mean_AP=('AP Exam Score', lambda x: round(x.mean(), 2)),
        Pct_3Plus=('Three_Plus', lambda x: round(x.mean() * 100, 1)),
    ).reset_index()

    fig, ax = plt.subplots(figsize=(10, max(3.5, len(ss) * 0.45 + 1.5)))
    ax.axis('off')
    cols = ['Year', 'Section', 'N', 'Mean Test Avg', 'Mean AP Score', '% 3 or Above']
    data = ss[['Schoolyear', 'Block', 'N', 'Mean_TA', 'Mean_AP', 'Pct_3Plus']].values.tolist()
    table = ax.table(cellText=data, colLabels=cols, cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1.2, 1.4)
    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor('#dddddd')
        cell.set_linewidth(0.5)
        if r == 0:
            cell.set_facecolor('white')
            cell.set_text_props(color='#333333', fontweight='bold')
        else:
            cell.set_facecolor('white')
    fig.suptitle('Section comparison by school year', fontsize=12, color='#333333')
    fig.tight_layout(rect=[0, 0, 1, 0.93])
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# AP Score Distribution (small histogram for PDF)
# ──────────────────────────────────────────────────────────────────────────────
def fig_ap_distribution(df):
    fig, ax = plt.subplots(figsize=(10, 5))
    counts = df['AP Exam Score'].value_counts().sort_index()

    # Thin bars — minimal ink for 5 discrete values
    ax.bar(counts.index, counts.values, width=0.5, color=PALETTE_ACCENT,
           alpha=0.7, edgecolor='none')
    for score, val in counts.items():
        ax.text(score, val + 1.5, str(val), ha='center', **N_FONT)

    ax.set_title('Overall distribution of AP exam scores')
    ax.set_xlabel('AP exam score')
    ax.set_ylabel('Number of students')
    ax.set_xticks([1, 2, 3, 4, 5])
    ax.set_ylim(0, counts.max() + 8)
    fig.tight_layout()
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# Figure — Small Multiples: AP Score Distribution by Section
# ──────────────────────────────────────────────────────────────────────────────
HIGHLIGHT_BEST = '#4a9e6e'   # muted green
HIGHLIGHT_WORST = '#c0392b'  # muted red


def fig_small_multiples_ap(df):
    sections = (df.groupby(['Schoolyear', 'Block'])
                  .size().reset_index(name='N')
                  .sort_values(['Schoolyear', 'Block']))
    n_sections = len(sections)
    ncols = min(3, n_sections)
    nrows = int(np.ceil(n_sections / ncols))

    # Compute mean AP score per section to identify best/worst
    sec_means = df.groupby(['Schoolyear', 'Block'])['AP Exam Score'].mean()
    best_key = sec_means.idxmax()
    worst_key = sec_means.idxmin()

    fig, axes = plt.subplots(nrows, ncols, figsize=(4 * ncols, 3.2 * nrows),
                             sharex=True, sharey=True)
    axes = np.atleast_2d(axes)
    global_max = 0

    for i, (_, sec) in enumerate(sections.iterrows()):
        r, c = divmod(i, ncols)
        ax = axes[r][c]
        mask = (df['Schoolyear'] == sec['Schoolyear']) & (df['Block'] == sec['Block'])
        sub = df.loc[mask]
        counts = sub['AP Exam Score'].value_counts().reindex([1, 2, 3, 4, 5], fill_value=0)
        global_max = max(global_max, counts.max())

        ax.bar(counts.index, counts.values, width=0.5, color=PALETTE_ACCENT,
               alpha=0.7, edgecolor='none')
        for score, val in counts.items():
            if val > 0:
                ax.text(score, val + 0.4, str(val), ha='center', fontsize=7, color='#777777')

        sec_key = (sec['Schoolyear'], sec['Block'])
        mean_ap = sec_means[sec_key]
        ax.set_title(f"{sec['Schoolyear']}  {sec['Block']}  (n={sec['N']}, μ={mean_ap:.2f})",
                     fontsize=9, color='#444444')
        ax.set_xticks([1, 2, 3, 4, 5])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        # Highlight best/worst with a colored border
        if sec_key == best_key:
            for spine in ax.spines.values():
                spine.set_visible(True)
                spine.set_edgecolor(HIGHLIGHT_BEST)
                spine.set_linewidth(2)
        elif sec_key == worst_key:
            for spine in ax.spines.values():
                spine.set_visible(True)
                spine.set_edgecolor(HIGHLIGHT_WORST)
                spine.set_linewidth(2)

    # Hide unused subplots
    for i in range(n_sections, nrows * ncols):
        r, c = divmod(i, ncols)
        axes[r][c].set_visible(False)

    # Shared labels
    for ax in axes[-1]:
        if ax.get_visible():
            ax.set_xlabel('AP score', fontsize=9)
    for ax in axes[:, 0]:
        ax.set_ylabel('Count', fontsize=9)

    # Set consistent y-limit after computing global max
    for row in axes:
        for ax in row:
            if ax.get_visible():
                ax.set_ylim(0, global_max + 3)

    fig.suptitle('AP score distribution by section', fontsize=12, color='#333333')
    fig.tight_layout(rect=[0, 0, 1, 0.94])
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# Figure — Small Multiples: Test Average Distribution by Section
# ──────────────────────────────────────────────────────────────────────────────
def fig_small_multiples_test_avg(df):
    sections = (df.groupby(['Schoolyear', 'Block'])
                  .size().reset_index(name='N')
                  .sort_values(['Schoolyear', 'Block']))
    n_sections = len(sections)
    ncols = min(3, n_sections)
    nrows = int(np.ceil(n_sections / ncols))

    # Compute median test average per section to identify best/worst
    sec_medians = df.groupby(['Schoolyear', 'Block'])['Test Average'].median()
    best_key = sec_medians.idxmax()
    worst_key = sec_medians.idxmin()

    fig, axes = plt.subplots(nrows, ncols, figsize=(4 * ncols, 3.2 * nrows),
                             sharex=True, sharey=True)
    axes = np.atleast_2d(axes)

    # Consistent x-range across all panels
    x_lo, x_hi = 35, 100

    for i, (_, sec) in enumerate(sections.iterrows()):
        r, c = divmod(i, ncols)
        ax = axes[r][c]
        mask = (df['Schoolyear'] == sec['Schoolyear']) & (df['Block'] == sec['Block'])
        sub = df.loc[mask]

        # Strip plot with jitter
        jitter = np.random.RandomState(42).uniform(-0.15, 0.15, size=len(sub))
        ax.scatter(sub['Test Average'], jitter, color=PALETTE_ACCENT,
                   alpha=0.5, s=18, zorder=5)
        # Median marker
        med = sub['Test Average'].median()
        ax.axvline(med, color='#333333', linewidth=1, linestyle='-', alpha=0.6)
        ax.text(med, 0.35, f'{med:.0f}', ha='center', fontsize=7, color='#333333')

        ax.set_title(f"{sec['Schoolyear']}  {sec['Block']}  (n={sec['N']})",
                     fontsize=9, color='#444444')
        ax.set_xlim(x_lo, x_hi)
        ax.set_ylim(-0.4, 0.5)
        ax.set_yticks([])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)

        # Highlight best/worst with a colored border
        sec_key = (sec['Schoolyear'], sec['Block'])
        if sec_key == best_key:
            for spine in ax.spines.values():
                spine.set_visible(True)
                spine.set_edgecolor(HIGHLIGHT_BEST)
                spine.set_linewidth(2)
        elif sec_key == worst_key:
            for spine in ax.spines.values():
                spine.set_visible(True)
                spine.set_edgecolor(HIGHLIGHT_WORST)
                spine.set_linewidth(2)

    # Hide unused subplots
    for i in range(n_sections, nrows * ncols):
        r, c = divmod(i, ncols)
        axes[r][c].set_visible(False)

    for ax in axes[-1]:
        if ax.get_visible():
            ax.set_xlabel('Test average', fontsize=9)

    fig.suptitle('Test average distribution by section (median marked)',
                 fontsize=12, color='#333333')
    fig.tight_layout(rect=[0, 0, 1, 0.94])
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# PDF Report
# ──────────────────────────────────────────────────────────────────────────────
CAPTIONS = {
    'correlation': (
        "Shows the relationship between in-class test averages and AP exam scores. "
        "Each bubble represents one or more students (scaled by frequency). "
        "The red regression line shows the overall trend along with a shaded 95% confidence interval band."
    ),
    'heatmap': (
        "Displays the frequency of AP exam scores (1–5) within each course summative letter grade band. "
        "Bubble sizes and colors scale with student count, with exact counts displayed inside each bubble."
    ),
    'stats': (
        "Descriptive statistics for in-class test averages, grouped by AP exam score received. "
        "N = number of students, Range = difference between highest and lowest test average in the group."
    ),
    'probability': (
        "For each test average range (aligned to the 10-point grading scale), shows the probability "
        "of receiving each AP score. Color intensity reflects probability — brighter cells are more likely outcomes."
    ),
    'trends': (
        "Mean in-class test average (blue, top panel) and mean AP score (red, bottom panel) by school year, "
        "with 95% confidence intervals. Error bars reflect uncertainty due to sample size. "
        "The ANOVA result at the bottom tests whether year-to-year differences are statistically significant."
    ),
    'boxplot': (
        "Distribution of in-class test averages for students who received each AP score. "
        "Box shows interquartile range; whiskers extend to 1.5x IQR. Individual student data points "
        "are overlaid. N = number of students per score."
    ),
    'three_plus': (
        "Percentage of students scoring 3 or above on the AP exam, broken down by letter grade (left) "
        "and by school year (right). N shown above each bar. Higher grades correlate with higher 3+ rates."
    ),
    'logistic': (
        "Logistic regression curves predicting the probability of scoring 3+, 4+, or 5 based on "
        "in-class test average. Tick marks show where each curve crosses the 50%, 75%, and 90% "
        "probability levels. Curves only shown where sufficient data exists to fit a reliable model."
    ),
    'sections': (
        "Comparison of class sections within each school year. Shows the number of students, "
        "mean test average, mean AP score, and percentage scoring 3 or above for each section."
    ),
    'ap_dist': (
        "Overall distribution of AP exam scores across all years and sections. "
        "Shows how many students received each score from 1 to 5."
    ),
    'sm_ap': (
        "Small multiples showing AP score distribution for each section-year combination. "
        "Shared axes allow direct comparison across sections."
    ),
    'sm_test_avg': (
        "Small multiples showing the spread of test averages for each section. "
        "Each dot is a student; the vertical line marks the section median."
    ),
    'logistic_table': (
        "Predicted probabilities from the logistic regression models, sampled at 5-point "
        "intervals of in-class test average. Darker cells indicate higher probability. "
        "Read across a row to see how likely each outcome is for a given test average."
    ),
}


def add_captioned_page(pdf, fig, caption_key, save_dir=None, filename=None):
    """Adds a figure to the PDF with a caption, and optionally saves as PNG."""
    caption = CAPTIONS.get(caption_key, '')
    if caption:
        fig.text(0.05, -0.02, caption, **CAPTION_FONT)
    fig.savefig(pdf, format='pdf', bbox_inches='tight', dpi=300)
    if save_dir and filename:
        fig.savefig(os.path.join(save_dir, filename), dpi=150, bbox_inches='tight')
    plt.close(fig)


def make_title_page(pdf, n_students, years):
    """Creates a title page for the PDF report."""
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.axis('off')
    ax.text(0.5, 0.65, 'AP Exam Score Analysis Report', transform=ax.transAxes,
            fontsize=20, ha='center', va='center', color='#333333')
    ax.text(0.5, 0.52, date.today().strftime('%B %d, %Y'), transform=ax.transAxes,
            fontsize=11, ha='center', va='center', color='#888888')
    summary = f"{n_students} students  ·  {years[0]} through {years[-1]}  ·  {len(years)} school years"
    ax.text(0.5, 0.42, summary, transform=ax.transAxes,
            fontsize=10, ha='center', va='center', color='#999999')
    fig.tight_layout()
    fig.savefig(pdf, format='pdf')
    plt.close(fig)


def make_summary_page(pdf, anova_text, anova_explanation, logistic_results,
                      tukey_rows):
    """Creates a final summary page with key statistical findings."""
    fig, ax = plt.subplots(figsize=(10, 8))
    ax.axis('off')
    ax.text(0.5, 0.95, 'Key statistical findings', transform=ax.transAxes,
            fontsize=14, ha='center', color='#333333')

    lines = [
        f"ANOVA (test averages across years):  {anova_text}",
        "",
        anova_explanation,
    ]

    # Tukey HSD post-hoc results
    if tukey_rows:
        lines.append("")
        lines.append("Tukey HSD post-hoc pairwise comparisons:")
        lines.append("  " + "─" * 45)
        for y1, y2, p, sig in tukey_rows:
            marker = " *" if sig else ""
            lines.append(f"  {y1} vs {y2}:  p = {p:.4f}{marker}")
        lines.append("  (* = significant at p < 0.05)")

    lines += [
        "",
        "Logistic regression thresholds (test average needed):",
        "  Target    50% likely    75% likely    90% likely    Accuracy",
        "  " + "─" * 60,
    ]

    for label, data in logistic_results.items():
        parts = [f"  {label:8s}"]
        for level in PROB_LEVELS:
            t = data.get(level)
            parts.append(f"{t:>8.1f}" if t is not None else "     n/a")
            parts.append("       ")
        parts.append(f"  {data['acc']:.0%}")
        lines.append(''.join(parts))

    lines.append("")
    # Plain-language summary for the most common question: what does it take to pass?
    t3_50 = logistic_results.get('3+', {}).get(0.50)
    t3_75 = logistic_results.get('3+', {}).get(0.75)
    if t3_50 is not None:
        lines.append(f"A student with a test average above {t3_50:.0f} has a >50% chance of scoring 3+.")
    if t3_75 is not None:
        lines.append(f"A test average above {t3_75:.0f} gives a >75% chance of scoring 3+.")

    text = '\n'.join(lines)
    ax.text(0.05, 0.85, text, transform=ax.transAxes, fontsize=10,
            va='top', family='monospace', color='#555555')
    fig.tight_layout()
    fig.savefig(pdf, format='pdf')
    plt.close(fig)


# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        if len(sys.argv) > 1:
            file_path = os.path.abspath(sys.argv[1])
        elif os.path.exists(os.path.join(script_dir, 'AP Test Score-Exam Score Tracking - Data.csv')):
            file_path = os.path.join(script_dir, 'AP Test Score-Exam Score Tracking - Data.csv')
        else:
            file_path = os.path.join(script_dir, 'your_data.csv')

        print(f"── Reading Data: {os.path.basename(file_path)} ──")
        raw = pd.read_csv(file_path)

        # Clean
        df = raw.dropna(subset=['Test Average', 'AP Exam Score', 'Test Average Grade']).copy()
        for col in ['Test Average', 'AP Exam Score']:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        df.dropna(subset=['Test Average', 'AP Exam Score'], inplace=True)
        df['AP Exam Score'] = df['AP Exam Score'].astype(int)

        # Validate interactively
        df = validate_grade_assignments(df, script_dir, file_path, raw)

        # Generate figures
        print("── Generating Figures ──")
        figures = {}

        print("   [1/8] Correlation plot")
        figures['correlation'] = fig_correlation(df)

        print("   [2/8] Grade heatmap")
        figures['heatmap'] = fig_heatmap(df)

        print("   [3/8] Descriptive stats table")
        figures['stats'] = fig_stats_table(df)

        print("   [4/8] Probability table")
        figures['probability'] = fig_probability_table(df.copy())

        print("   [5/8] Yearly trends with CI")
        trends_fig, tukey_rows = fig_yearly_trends(df)
        figures['trends'] = trends_fig

        print("   [6/8] Box plot")
        figures['boxplot'] = fig_boxplot(df)

        print("   [7/8] 3-or-above rate")
        figures['three_plus'] = fig_three_plus(df)

        print("   [8/8] Logistic regression (multi-threshold)")
        logistic_fig, logistic_results = fig_logistic(df)
        figures['logistic'] = logistic_fig

        # Logistic probabilities table
        print("   [+] Logistic probabilities table")
        figures['logistic_table'] = fig_logistic_table(logistic_results)

        # Extra pages
        figures['sections'] = fig_section_table(df)
        figures['ap_dist'] = fig_ap_distribution(df)

        print("   [+] Small multiples — AP scores by section")
        figures['sm_ap'] = fig_small_multiples_ap(df)
        print("   [+] Small multiples — Test averages by section")
        figures['sm_test_avg'] = fig_small_multiples_test_avg(df)

        # ANOVA text and plain-language explanation for summary page
        groups = [g['Test Average'].values for _, g in df.groupby('Schoolyear')]
        f_stat, p_val = stats.f_oneway(*groups)
        anova_text = f"F = {f_stat:.2f}, p = {p_val:.4f}"
        if p_val < 0.05:
            anova_text += " (statistically significant)"
            anova_explanation = (
                f"The one-way ANOVA tests whether mean test averages differ across\n"
                f"school years. With p = {p_val:.4f} (below the 0.05 threshold),\n"
                f"we can reject the null hypothesis that all years have the same mean.\n"
                f"In plain terms: there IS a statistically meaningful difference in\n"
                f"student test averages between at least two of the school years.\n"
                f"This does not tell us which years differ — the Tukey HSD post-hoc\n"
                f"test (shown below) identifies which specific year-pairs show\n"
                f"significant differences."
            )
        else:
            anova_text += " (not significant)"
            anova_explanation = (
                f"The one-way ANOVA tests whether mean test averages differ across\n"
                f"school years. With p = {p_val:.4f} (above the 0.05 threshold),\n"
                f"we cannot reject the null hypothesis that all years have the same\n"
                f"mean. In plain terms: the differences in test averages between\n"
                f"school years are small enough that they could easily be due to\n"
                f"normal random variation, not a real underlying change."
            )

        # Build PDF
        print("\n── Building PDF Report ──")
        pdf_path = os.path.join(script_dir, 'AP_Exam_Analysis_Report.pdf')
        years = sorted(df['Schoolyear'].unique())

        with PdfPages(pdf_path) as pdf:
            make_title_page(pdf, len(df), years)

            page_order = [
                ('correlation', 'Overall_Correlation_Plot.png'),
                ('heatmap', 'Letter_Grade_Frequency_Heatmap.png'),
                ('stats', 'Descriptive_Stats_Table.png'),
                ('probability', 'Probability_Heatmap_Table.png'),
                ('trends', 'Yearly_Trends_With_CI.png'),
                ('boxplot', 'Test_Average_by_AP_Score_Boxplot.png'),
                ('three_plus', 'Three_Plus_Rate.png'),
                ('logistic', 'Logistic_Regression_Threshold.png'),
                ('logistic_table', 'Logistic_Probabilities_Table.png'),
                ('sections', 'Section_Comparison_Table.png'),
                ('ap_dist', 'AP_Score_Distribution.png'),
                ('sm_ap', 'Small_Multiples_AP_Scores.png'),
                ('sm_test_avg', 'Small_Multiples_Test_Averages.png'),
            ]

            for key, png_name in page_order:
                add_captioned_page(pdf, figures[key], key, script_dir, png_name)

            make_summary_page(pdf, anova_text, anova_explanation, logistic_results,
                              tukey_rows)

        print(f"-> PDF saved to: {pdf_path}")
        print(f"-> PNGs saved to: {script_dir}/")
        print("\n── Done ──")

    except FileNotFoundError:
        print(f"Error: Data file not found at '{file_path}'. Place it in the same folder as this script.")
    except Exception as e:
        import traceback
        print(f"Error: {e}")
        traceback.print_exc()
