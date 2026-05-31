import ast
import sys

with open('student_feedback_analyzer.py', 'r') as f:
    src = f.read()

lines = src.split('\n')

class FuncVisitor(ast.NodeVisitor):
    def __init__(self):
        self.funcs = {}
    def visit_FunctionDef(self, node):
        self.funcs[node.name] = (node.lineno - 1, node.end_lineno)
        self.generic_visit(node)

v = FuncVisitor()
v.visit(ast.parse(src))

funcs_to_replace = {
    'create_feeling_words_heatmap': '''def create_feeling_words_heatmap(class_df, course, period_order, label_map, pdf, top_n=15):
    print(f"  > Generating Feeling-Words Heatmap for '{course}'...")
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
    
    # Use GnBu instead of Greys to add a hint of color
    sns.heatmap(counts.T, annot=True, fmt='d', cmap='GnBu', cbar=False,
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

    if pdf:
        pdf.savefig(fig)
    plt.close(fig)''',

    'create_score_distribution_by_period': '''def create_score_distribution_by_period(class_df, course, period_order, label_map, pdf):
    print(f"  > Generating Score Distribution by Period for '{course}'...")
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

    # Added thick white edgecolor to create a visual gap between box and whiskers
    sns.boxplot(
        x='Survey_Period', y='Score', hue='Score Type',
        data=plot_long, order=period_order,
        palette=palette, width=0.1, showcaps=False,
        boxprops={'zorder': 3, 'linewidth': 2, 'edgecolor': 'white', 'alpha': 0.8}, 
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
    
    # Moved legend to avoid overlapping data
    ax.legend(unique_handles, unique_labels, frameon=False, loc='upper left', bbox_to_anchor=(1.02, 1), ncol=1, fontsize=9)

    ax.set_xticks(range(len(period_order)))
    ax.set_xticklabels(short_labels, rotation=0, ha='center', fontsize=10, color='#333333')
    ax.set_ylim(1, 7)
    ax.set_yticks(range(1, 8))
    ax.set_title(f"Score Distributions Over Time: {course}", fontsize=14, pad=20, color='#222222', loc='left')
    ax.set_ylabel("Score (1–7)", fontsize=11, color='#555555')
    ax.set_xlabel(None)
    plt.tight_layout()

    if pdf:
        pdf.savefig(fig)
    plt.close(fig)''',

    'create_feelings_histogram': '''def create_feelings_histogram(df, course_name, pdf=None):
    print(f"Generating Feelings Histogram for {course_name}...")

    if course_name == "All Classes Combined":
        class_df = df.copy()
    else:
        class_df = df[df[CLASS_COL] == course_name]
    
    word_counts = class_df[FEELING_WORD_COL].value_counts()
    if word_counts.empty:
        return

    plot_df = pd.DataFrame({'Word': word_counts.index, 'Count': word_counts.values})
    plot_df['Sentiment'] = plot_df['Word'].map(WORD_TO_SENTIMENT).fillna('Other')
    
    sentiment_order = ['Positive', 'Neutral', 'Negative', 'Other']
    plot_df['Sentiment'] = pd.Categorical(plot_df['Sentiment'], categories=sentiment_order, ordered=True)
    plot_df = plot_df.sort_values('Sentiment')
    
    # Muted palette for histogram
    tufte_colors = {'Positive': '#74a089', 'Neutral': '#c2c2c2', 'Negative': '#cb7a77', 'Other': '#e0e0e0'}
    bar_colors = plot_df['Sentiment'].map(tufte_colors)

    fig, ax = plt.subplots(figsize=(11, 6))
    
    ax.bar(plot_df['Word'], plot_df['Count'], color=bar_colors, alpha=0.9, width=0.6)

    setup_tufte_axes(ax, ygrid=False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.tick_params(axis='both', length=0)
    
    plt.setp(ax.get_xticklabels(), rotation=45, ha='right', color='#333333', fontsize=10) 
    ax.set_yticks([])

    for i, count in enumerate(plot_df['Count']):
        ax.text(i, count + max(plot_df['Count']) * 0.02, str(count), ha='center', va='bottom', fontsize=9, color='#444444')

    ax.set_title(f'Distribution of Student Feelings: {course_name} (n={len(class_df)})', fontsize=14, pad=20, color='#222222', loc='left')
    ax.set_ylabel(None)
    ax.set_xlabel(None)
    
    present_sentiments = plot_df['Sentiment'].unique()
    legend_patches = [plt.Rectangle((0,0),1,1, color=tufte_colors.get(s, '#cccccc')) for s in sentiment_order if s in present_sentiments]
    legend_labels = [s for s in sentiment_order if s in present_sentiments]
    if legend_patches:
        ax.legend(legend_patches, legend_labels, frameon=False, loc='upper right', bbox_to_anchor=(1.02, 1))

    plt.tight_layout()
    if pdf:
        pdf.savefig(fig)
    plt.close(fig)''',

    'create_comparative_bubble_chart': '''def create_comparative_bubble_chart(df, course_name, pdf=None):
    print(f"Generating 'How's it Going?' Bubble Chart for {course_name}...")
    
    if course_name == "All Classes Combined":
        class_df = df.copy()
    else:
        class_df = df[df[CLASS_COL] == course_name].copy()
    
    class_scores = class_df[[CLASS_SCORE_COL]].rename(columns={CLASS_SCORE_COL: 'Score'})
    class_scores['Category'] = f'In {course_name}'
    
    overall_scores = class_df[[OVERALL_SCORE_COL]].rename(columns={OVERALL_SCORE_COL: 'Score'})
    overall_scores['Category'] = 'In School Overall'
    
    plot_data = pd.concat([class_scores, overall_scores]).dropna(subset=['Score'])

    if plot_data.empty:
        return

    fig, ax = plt.subplots(figsize=(8, 6))

    counts = plot_data.groupby(['Category', 'Score']).size().reset_index(name='Count')

    tufte_palette = {'In School Overall': '#c78163', f'In {course_name}': '#4f8ca6'}
    
    sns.scatterplot(x='Category', y='Score', hue='Category', size='Count', sizes=(100, 1500), 
                    data=counts, palette=tufte_palette,
                    alpha=0.8, edgecolor='white', linewidth=1, ax=ax, legend=False)

    for _, row in counts.iterrows():
        ax.text(row['Category'], row['Score'], str(row['Count']), 
                color='white', ha='center', va='center', fontsize=9, fontweight='bold')

    for i, category in enumerate(plot_data['Category'].unique()):
        median_val = plot_data[plot_data['Category'] == category]['Score'].median()
        ax.plot([i-0.3, i+0.3], [median_val, median_val], color='#444444', lw=2, linestyle='--')

    setup_tufte_axes(ax, ygrid=False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.tick_params(axis='both', length=0)
    ax.set_yticks([])

    ax.set_title(f"Score Comparison: Class vs. School (n={len(class_df)})", fontsize=14, pad=20, color='#222222', loc='left')
    ax.set_ylabel(None)
    ax.set_xlabel(None)
    ax.set_ylim(0.5, 7.5)

    plt.tight_layout()
    if pdf:
        pdf.savefig(fig)
    plt.close(fig)''',

    'create_sentiment_score_boxplot': '''def create_sentiment_score_boxplot(df, course_name, pdf=None):
    print(f"Generating Sentiment vs Score Boxplot for {course_name}...")
    
    if course_name == "All Classes Combined":
        class_df = df.copy()
    else:
        class_df = df[df[CLASS_COL] == course_name].copy()
        
    class_df['Sentiment'] = class_df[FEELING_WORD_COL].map(WORD_TO_SENTIMENT).fillna('Other')
    
    if class_df[CLASS_SCORE_COL].dropna().empty:
        return
        
    sentiment_order = ['Positive', 'Neutral', 'Negative'] 
    class_df['Sentiment'] = pd.Categorical(class_df['Sentiment'], categories=sentiment_order, ordered=True)
    
    fig, ax = plt.subplots(figsize=(9, 6))
    
    tufte_colors = {'Positive': '#74a089', 'Neutral': '#c2c2c2', 'Negative': '#cb7a77'}
    
    sns.boxplot(x='Sentiment', y=CLASS_SCORE_COL, data=class_df, palette=tufte_colors, ax=ax, order=sentiment_order,
                width=0.4, boxprops={'linewidth': 1, 'alpha': 0.8, 'edgecolor': 'white'},
                whiskerprops={'linewidth': 1, 'color': '#777777'},
                medianprops={'linewidth': 2, 'color': 'white'}, fliersize=0)
    
    counts = class_df.groupby(['Sentiment', CLASS_SCORE_COL], observed=True).size().reset_index(name='Count')
    counts = counts[counts['Count'] > 0]
    
    if not counts.empty:
        sns.scatterplot(x='Sentiment', y=CLASS_SCORE_COL, size='Count', sizes=(100, 1500), 
                        data=counts, color='#444444', alpha=0.6, edgecolor='white', linewidth=1, ax=ax, legend=False)
        x_map = {cat: i for i, cat in enumerate(sentiment_order)}
        for _, row in counts.iterrows():
            if pd.notna(row['Sentiment']):
                ax.text(x_map[row['Sentiment']], row[CLASS_SCORE_COL], str(row['Count']), 
                        color='white', ha='center', va='center', fontsize=9, fontweight='bold')
    
    setup_tufte_axes(ax, ygrid=True)
    
    ax.set_title(f"Score Distribution by Sentiment: {course_name} (n={len(class_df)})", fontsize=14, pad=20, color='#222222', loc='left')
    ax.set_ylabel("Score (1-7)", fontsize=11, color='#555555')
    ax.set_xlabel(None)
    ax.set_ylim(0.5, 7.5)
    ax.set_yticks(range(1, 8))
    
    plt.tight_layout()
    if pdf:
        pdf.savefig(fig)
    plt.close(fig)''',

    'create_score_correlation_plot': '''def create_score_correlation_plot(df, course_name, pdf=None):
    print(f"Generating Score Correlation Plot for {course_name}...")
    
    if course_name == "All Classes Combined":
        class_df = df.copy()
    else:
        class_df = df[df[CLASS_COL] == course_name].copy()
    
    plot_data = class_df[[CLASS_SCORE_COL, OVERALL_SCORE_COL]].dropna()
    if len(plot_data) < 2:
        return
        
    fig, ax = plt.subplots(figsize=(7, 7))
    
    counts = plot_data.groupby([OVERALL_SCORE_COL, CLASS_SCORE_COL]).size().reset_index(name='Count')
    
    sns.scatterplot(x=OVERALL_SCORE_COL, y=CLASS_SCORE_COL, size='Count', sizes=(100, 1500), 
                    data=counts, color='#4f8ca6', alpha=0.8, edgecolor='white', linewidth=1, ax=ax, legend=False)
                    
    for _, row in counts.iterrows():
        ax.text(row[OVERALL_SCORE_COL], row[CLASS_SCORE_COL], str(row['Count']), 
                color='white', ha='center', va='center', fontsize=9, fontweight='bold')
    
    x_vals = plot_data[OVERALL_SCORE_COL].values
    y_vals = plot_data[CLASS_SCORE_COL].values
    
    if np.std(x_vals) > 0:
        z = np.polyfit(x_vals, y_vals, 1)
        p = np.poly1d(z)
        x_trend = np.linspace(min(x_vals), max(x_vals), 100)
        ax.plot(x_trend, p(x_trend), color='#cb7a77', linestyle='--', alpha=0.8, linewidth=2, label='Linear Trend')
    
    correlation = plot_data[OVERALL_SCORE_COL].corr(plot_data[CLASS_SCORE_COL])
    
    setup_tufte_axes(ax, ygrid=True)
    ax.xaxis.grid(True, color='#f4f4f4', linestyle='-', linewidth=0.8)
    
    ax.set_title(f"Score Correlation: Class vs. School\nr = {correlation:.2f} (n={len(plot_data)})", fontsize=14, pad=20, color='#222222', loc='left')
    ax.set_xlabel("Overall School Score (1-7)", fontsize=11, color='#555555')
    ax.set_ylabel("In-Class Score (1-7)", fontsize=11, color='#555555')
    ax.set_xlim(0.5, 7.5)
    ax.set_ylim(0.5, 7.5)
    ax.set_xticks(range(1, 8))
    ax.set_yticks(range(1, 8))
    ax.legend(frameon=False, loc='upper left')
    
    plt.tight_layout()
    if pdf:
        pdf.savefig(fig)
    plt.close(fig)'''
}

# Replace functions from bottom to top to avoid shifting line numbers
sorted_funcs = sorted(funcs_to_replace.keys(), key=lambda k: v.funcs[k][0], reverse=True)

for func_name in sorted_funcs:
    start_line, end_line = v.funcs[func_name]
    replacement = funcs_to_replace[func_name].split('\n')
    lines = lines[:start_line] + replacement + lines[end_line:]

with open('student_feedback_analyzer.py', 'w') as f:
    f.write('\n'.join(lines))

print("Patch applied successfully.")
