import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from student_feedback_analyzer import CLASS_SCORE_COL, OVERALL_SCORE_COL, FEELING_WORD_COL, CLASS_COL, WORD_TO_SENTIMENT

def setup_tufte_axes(ax, ygrid=True):
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_linewidth(0.5)
    ax.spines['bottom'].set_linewidth(0.5)
    ax.spines['left'].set_color('#aaaaaa')
    ax.spines['bottom'].set_color('#aaaaaa')
    ax.tick_params(colors='#555555', width=0.5, length=4)
    if ygrid:
        ax.yaxis.grid(True, color='#f4f4f4', linestyle='-', linewidth=0.8)
        ax.set_axisbelow(True)
    else:
        ax.grid(False)

def create_score_trends_chart(class_df, course, period_order, label_map, pdf):
    print(f"  > Generating Score Trends for '{course}'...")
    trend_data = class_df.groupby('Survey_Period', observed=True)[[CLASS_SCORE_COL, OVERALL_SCORE_COL]].mean().reset_index()
    trend_data['Survey_Period'] = pd.Categorical(trend_data['Survey_Period'], categories=period_order, ordered=True)
    trend_data = trend_data.sort_values('Survey_Period')

    x_positions = range(len(period_order))
    class_means = trend_data.set_index('Survey_Period')[CLASS_SCORE_COL].reindex(period_order).values
    overall_means = trend_data.set_index('Survey_Period')[OVERALL_SCORE_COL].reindex(period_order).values

    fig, ax = plt.subplots(figsize=(10, 5))
    setup_tufte_axes(ax)

    ax.plot(x_positions, class_means, marker='o', linewidth=2.5, markersize=8, label='Class Rating', color='#2a9d8f')
    ax.plot(x_positions, overall_means, marker='s', linewidth=2.5, markersize=8, label='School Rating', color='#e76f51', linestyle='--')

    ax.set_ylim(1, 7.2)
    ax.set_yticks(range(1, 8))
    ax.set_xticks(list(x_positions))
    
    from student_feedback_analyzer import _short_labels_for
    ax.set_xticklabels(_short_labels_for(period_order, label_map), rotation=0, ha='center', fontsize=10, color='#333333')
    
    ax.set_title(f"Average Score Over Time: {course}", fontsize=14, pad=20, color='#222222', loc='left')
    ax.set_ylabel("Score (1-7)", fontsize=11, color='#555555', labelpad=10)
    
    ax.legend(frameon=False, loc='upper right', bbox_to_anchor=(1, 1.05), ncol=2, fontsize=10, labelcolor='#444444')
    plt.tight_layout()

    pdf.savefig(fig)
    plt.close(fig)

def create_sentiment_evolution_chart(class_df, course, period_order, label_map, pdf):
    print(f"  > Generating Sentiment Evolution for '{course}'...")
    class_df = class_df.copy()
    class_df['Sentiment'] = class_df[FEELING_WORD_COL].astype(str).str.strip().str.title().map(WORD_TO_SENTIMENT).fillna('Other')

    sentiment_counts = class_df.groupby(['Survey_Period', 'Sentiment'], observed=True).size().unstack(fill_value=0)
    cols_present = [col for col in ['Positive', 'Neutral', 'Negative', 'Other'] if col in sentiment_counts.columns]
    sentiment_counts = sentiment_counts[cols_present]

    sentiment_pct = sentiment_counts.div(sentiment_counts.sum(axis=1), axis=0) * 100
    sentiment_pct = sentiment_pct.reindex(period_order).dropna()
    from student_feedback_analyzer import _short_labels_for
    sentiment_pct.index = _short_labels_for(sentiment_pct.index.tolist(), label_map)

    tufte_colors = {'Positive': '#74a089', 'Neutral': '#c2c2c2', 'Negative': '#cb7a77', 'Other': '#e0e0e0'}
    colors = [tufte_colors.get(c, '#cccccc') for c in cols_present]
    
    fig, ax = plt.subplots(figsize=(10, 5))
    sentiment_pct.plot(kind='bar', stacked=True, color=colors, width=0.6, ax=ax, edgecolor='white', linewidth=0.5)

    setup_tufte_axes(ax, ygrid=False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.tick_params(axis='x', length=0)
    ax.tick_params(axis='y', length=0)

    ax.set_title(f"Sentiment Distribution Over Time: {course}", fontsize=14, pad=20, color='#222222', loc='left')
    ax.set_xlabel(None)
    ax.set_ylim(0, 100)
    ax.set_yticklabels([f"{int(y)}%" for y in ax.get_yticks()], color='#555555')
    plt.xticks(rotation=0, ha='center', fontsize=10, color='#333333')
    
    ax.legend(title='', bbox_to_anchor=(1.02, 1), loc='upper left', frameon=False, labelcolor='#444444')
    plt.tight_layout()

    pdf.savefig(fig)
    plt.close(fig)

def create_small_multiples_score_trends(df, classes, period_order, label_map, pdf):
    print("  > Generating Score Trends by Course...")
    from student_feedback_analyzer import _short_labels_for
    short_labels = _short_labels_for(period_order, label_map)

    n = len(classes)
    if n == 0:
        return
    ncols = min(3, n)
    nrows = int(np.ceil(n / ncols))

    fig, axes = plt.subplots(nrows, ncols, figsize=(4.5 * ncols, 3.5 * nrows), sharey=True, sharex=True, squeeze=False)

    overall_means = df.groupby('Survey_Period', observed=True)[CLASS_SCORE_COL].mean().reindex(period_order)

    for idx, course in enumerate(classes):
        r, c = idx // ncols, idx % ncols
        ax = axes[r][c]
        class_df = df[df[CLASS_COL] == course]
        means = class_df.groupby('Survey_Period', observed=True)[CLASS_SCORE_COL].mean().reindex(period_order)

        setup_tufte_axes(ax, ygrid=True)
        ax.plot(range(len(period_order)), overall_means.values, color='#d9d9d9', linewidth=1.5, label='All classes')
        ax.plot(range(len(period_order)), means.values, marker='o', linewidth=2, color='#2a9d8f', markersize=6, label=course)

        ax.set_title(course, fontsize=11, color='#333333', loc='left')
        ax.set_ylim(1, 7)
        ax.set_yticks([1, 3, 5, 7])
        
        if r == nrows - 1 or (r == nrows - 2 and idx + ncols >= n):
            ax.set_xticks(range(len(period_order)))
            ax.set_xticklabels(short_labels, rotation=45, ha='right', fontsize=9, color='#555555')
        else:
            ax.tick_params(labelbottom=False)

    for idx in range(n, nrows * ncols):
        axes[idx // ncols][idx % ncols].axis('off')

    fig.suptitle("In-Class Score Trends by Course", fontsize=14, y=1.02, color='#222222')
    fig.text(-0.02, 0.5, 'Average Score (1-7)', va='center', rotation='vertical', fontsize=11, color='#555555')
    plt.tight_layout()

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)

def create_feeling_words_heatmap(class_df, course, period_order, label_map, pdf, top_n=15):
    print(f"  > Generating Feeling-Words Heatmap for '{course}'...")
    from student_feedback_analyzer import _short_labels_for
    short_labels = _short_labels_for(period_order, label_map)

    word_series = class_df[FEELING_WORD_COL].dropna()
    word_series = word_series[word_series.astype(str).str.strip() != '']
    if word_series.empty:
        return

    top_words = word_series.value_counts().head(top_n).index.tolist()
    sub = class_df[class_df[FEELING_WORD_COL].isin(top_words)]
    counts = sub.groupby(['Survey_Period', FEELING_WORD_COL], observed=True).size().unstack(fill_value=0)
    counts = counts.reindex(index=period_order, columns=top_words, fill_value=0)

    def word_sort_key(w):
        sent = WORD_TO_SENTIMENT.get(w, 'Other')
        order = {'Positive': 0, 'Neutral': 1, 'Negative': 2, 'Other': 3}
        return (order.get(sent, 3), -counts[w].sum())
    ordered_words = sorted(top_words, key=word_sort_key)
    counts = counts[ordered_words]

    fig, ax = plt.subplots(figsize=(max(8, len(period_order) * 1.5), max(5, len(ordered_words) * 0.45)))
    
    sns.heatmap(counts.T, annot=True, fmt='d', cmap='Greys', cbar=False,
                linewidths=1.5, linecolor='white', ax=ax, annot_kws={'fontsize': 10})

    ax.tick_params(axis='both', length=0)
    
    tufte_text_colors = {'Positive': '#4a7c59', 'Neutral': '#777777', 'Negative': '#9e504e', 'Other': '#888888'}
    for tick_label in ax.get_yticklabels():
        w = tick_label.get_text()
        sent = WORD_TO_SENTIMENT.get(w, 'Other')
        tick_label.set_color(tufte_text_colors.get(sent, '#333333'))
        tick_label.set_fontweight('normal')

    ax.set_xticklabels(short_labels, rotation=0, ha='center', fontsize=10, color='#333333')
    ax.set_title(f"Feeling-Word Frequency: {course}", fontsize=14, pad=20, color='#222222', loc='left')
    ax.set_xlabel(None)
    ax.set_ylabel(None)
    plt.tight_layout()

    pdf.savefig(fig)
    plt.close(fig)

def _stagger_label_positions(values, min_gap):
    order = sorted(range(len(values)), key=lambda i: values[i])
    sorted_vals = [values[i] for i in order]
    adjusted = list(sorted_vals)
    for i in range(1, len(adjusted)):
        if adjusted[i] - adjusted[i - 1] < min_gap:
            adjusted[i] = adjusted[i - 1] + min_gap
    out = [0.0] * len(values)
    for new_pos, orig_idx in zip(adjusted, order):
        out[orig_idx] = new_pos
    return out

def create_sentiment_slopegraph(df, classes, period_order, label_map, pdf):
    print("  > Generating Sentiment Slopegraph (first vs last period)...")
    if len(period_order) < 2:
        return

    first_period, last_period = period_order[0], period_order[-1]
    df_local = df.copy()
    df_local['Sentiment'] = df_local[FEELING_WORD_COL].astype(str).str.strip().str.title().map(WORD_TO_SENTIMENT).fillna('Other')

    rows = []
    for course in classes:
        class_df = df_local[df_local[CLASS_COL] == course]
        for period in (first_period, last_period):
            period_df = class_df[class_df['Survey_Period'] == period]
            if len(period_df) == 0:
                rows.append({'course': course, 'period': period, 'pct_positive': np.nan, 'n': 0})
            else:
                pct = (period_df['Sentiment'] == 'Positive').mean() * 100
                rows.append({'course': course, 'period': period, 'pct_positive': pct, 'n': len(period_df)})

    slope_df = pd.DataFrame(rows)
    valid_courses = [c for c in classes
                     if slope_df[(slope_df['course'] == c) & (slope_df['period'] == first_period)]['n'].iloc[0] > 0
                     and slope_df[(slope_df['course'] == c) & (slope_df['period'] == last_period)]['n'].iloc[0] > 0]

    if not valid_courses:
        return

    starts = [slope_df[(slope_df['course'] == c) & (slope_df['period'] == first_period)]['pct_positive'].iloc[0] for c in valid_courses]
    ends = [slope_df[(slope_df['course'] == c) & (slope_df['period'] == last_period)]['pct_positive'].iloc[0] for c in valid_courses]

    fig_height = max(5, len(valid_courses) * 0.5)
    min_gap = 100 / (fig_height * 3.5)
    start_label_y = _stagger_label_positions(starts, min_gap)
    end_label_y = _stagger_label_positions(ends, min_gap)

    fig, ax = plt.subplots(figsize=(8, fig_height))
    setup_tufte_axes(ax, ygrid=False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.tick_params(axis='both', length=0)
    ax.set_yticks([])

    for course, start, end, sy, ey in zip(valid_courses, starts, ends, start_label_y, end_label_y):
        color = '#4a7c59' if end >= start else '#9e504e'
        alpha = 0.8
        ax.plot([0, 1], [start, end], marker='o', color=color, linewidth=1.5, markersize=5, alpha=alpha)
        
        if abs(sy - start) > 0.5:
            ax.plot([-0.02, -0.05], [start, sy], color='#cccccc', linewidth=0.5)
        if abs(ey - end) > 0.5:
            ax.plot([1.02, 1.05], [end, ey], color='#cccccc', linewidth=0.5)
            
        ax.text(-0.06, sy, f"{course} ({start:.0f}%)", ha='right', va='center', fontsize=9, color='#444444')
        ax.text(1.06, ey, f"{course} ({end:.0f}%)", ha='left', va='center', fontsize=9, color='#444444')

    ax.set_xlim(-0.5, 1.5)
    ax.set_ylim(-5, 105)
    ax.set_xticks([0, 1])
    ax.set_xticklabels([label_map.get(first_period, first_period), label_map.get(last_period, last_period)],
                       fontsize=11, fontweight='bold', color='#333333')
    ax.set_title("Shift in Positive Sentiment: First vs Last Period", fontsize=14, pad=20, color='#222222', loc='left')
    plt.tight_layout()

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)

def create_score_distribution_by_period(class_df, course, period_order, label_map, pdf):
    print(f"  > Generating Score Distribution by Period for '{course}'...")
    from student_feedback_analyzer import _short_labels_for
    short_labels = _short_labels_for(period_order, label_map)

    plot_data = class_df[['Survey_Period', CLASS_SCORE_COL, OVERALL_SCORE_COL]].dropna(subset=[CLASS_SCORE_COL]).copy()
    if plot_data.empty:
        return

    plot_data['Survey_Period'] = pd.Categorical(plot_data['Survey_Period'], categories=period_order, ordered=True)
    plot_data = plot_data.sort_values('Survey_Period')

    plot_long = plot_data.melt(
        id_vars='Survey_Period',
        value_vars=[CLASS_SCORE_COL, OVERALL_SCORE_COL],
        var_name='Score Type',
        value_name='Score'
    ).dropna(subset=['Score'])
    plot_long['Score Type'] = plot_long['Score Type'].map({
        CLASS_SCORE_COL:   'Class Rating',
        OVERALL_SCORE_COL: 'School Rating'
    })

    CLASS_COLOR  = '#4f8ca6'
    SCHOOL_COLOR = '#c78163'
    palette = {'Class Rating': CLASS_COLOR, 'School Rating': SCHOOL_COLOR}

    fig, ax = plt.subplots(figsize=(max(9, len(period_order) * 2), 5))
    setup_tufte_axes(ax)

    sns.violinplot(
        x='Survey_Period', y='Score', hue='Score Type',
        data=plot_long, order=period_order,
        palette=palette, inner=None, cut=0,
        split=False, dodge=True, gap=0.05,
        linewidth=0, alpha=0.3, ax=ax
    )

    sns.boxplot(
        x='Survey_Period', y='Score', hue='Score Type',
        data=plot_long, order=period_order,
        palette=palette, width=0.1, showcaps=False,
        boxprops={'zorder': 3, 'linewidth': 1, 'alpha': 0.8}, 
        whiskerprops={'zorder': 3, 'linewidth': 1, 'color': '#777777'},
        medianprops={'color': 'white', 'linewidth': 2, 'zorder': 4},
        showfliers=False, dodge=True, gap=0.05,
        ax=ax
    )

    x_base = np.arange(len(period_order))
    hue_offsets = {'Class Rating': -0.2, 'School Rating': 0.2}
    for score_type, col, color, marker in [
        ('Class Rating',  CLASS_SCORE_COL,   CLASS_COLOR,  'D'),
        ('School Rating', OVERALL_SCORE_COL, SCHOOL_COLOR, 's'),
    ]:
        means = plot_data.groupby('Survey_Period', observed=True)[col].mean().reindex(period_order)
        x_pos = x_base + hue_offsets[score_type]
        ax.plot(x_pos, means.values, color=color, marker=marker,
                linewidth=1.5, markersize=5, zorder=6,
                label=f'{score_type} mean', linestyle='--', alpha=0.9)

    handles, labels_leg = ax.get_legend_handles_labels()
    seen, unique_handles, unique_labels = set(), [], []
    for h, lbl in zip(handles, labels_leg):
        if lbl not in seen:
            seen.add(lbl)
            unique_handles.append(h)
            unique_labels.append(lbl)
    
    ax.legend(unique_handles, unique_labels, frameon=False, loc='upper right', bbox_to_anchor=(1, 1.05), ncol=2, fontsize=9)

    ax.set_xticks(range(len(period_order)))
    ax.set_xticklabels(short_labels, rotation=0, ha='center', fontsize=10, color='#333333')
    ax.set_ylim(1, 7)
    ax.set_yticks(range(1, 8))
    ax.set_title(f"Score Distributions Over Time: {course}", fontsize=14, pad=20, color='#222222', loc='left')
    ax.set_ylabel("Score (1–7)", fontsize=11, color='#555555')
    ax.set_xlabel(None)
    plt.tight_layout()

    pdf.savefig(fig)
    plt.close(fig)

def create_response_volume_bar(df, period_order, label_map, pdf):
    print("  > Generating Response Volume per Period...")
    from student_feedback_analyzer import _short_labels_for
    short_labels = _short_labels_for(period_order, label_map)
    counts = df.groupby('Survey_Period', observed=True).size().reindex(period_order, fill_value=0)

    fig, ax = plt.subplots(figsize=(max(8, len(period_order) * 1.2), 3))
    
    bars = ax.bar(range(len(period_order)), counts.values, color='#c9c9c9', width=0.4)

    setup_tufte_axes(ax, ygrid=False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color('#aaaaaa')
    ax.tick_params(axis='y', length=0)
    ax.set_yticks([])

    for bar, val in zip(bars, counts.values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(counts.values) * 0.02,
                str(int(val)), ha='center', va='bottom', fontsize=10, color='#444444')

    ax.set_xticks(range(len(period_order)))
    ax.set_xticklabels(short_labels, rotation=0, ha='center', fontsize=10, color='#333333')
    ax.set_title("Responses Collected per Period", fontsize=14, pad=15, color='#222222', loc='left')
    ax.set_ylabel(None)
    plt.tight_layout()

    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)
