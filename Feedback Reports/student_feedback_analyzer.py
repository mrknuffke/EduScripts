import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import numpy as np
import textwrap
import emoji
from scipy.stats import f_oneway, ttest_rel
from statsmodels.stats.multicomp import pairwise_tukeyhsd
from matplotlib.ticker import MaxNLocator
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import ssl
import json
from pathlib import Path

try:
    _create_unverified_https_context = ssl._create_unverified_context
except AttributeError:
    pass
else:
    ssl._create_default_https_context = _create_unverified_https_context

# --- SETTINGS ---
CLASS_COL = "Which of my classes are you in?"
CLASS_SCORE_COL = "Question 1: How are things going for you in this class?"
OVERALL_SCORE_COL = "How are things going overall as a student across all of your classes)?"
FEELING_WORD_COL = "Question 2: If you had to choose a single word to describe your experience in this class right now as a student, what would it be?"
TIMESTAMP_COL = "Timestamp"

SENTIMENT_MAP = {
    'Positive': ['Affirmed', 'Calm', 'Comfortable', 'Hopeful', 'Motivated', 'Supported', 'Purposeful', 'Curious', 'Proud'],
    'Neutral': ['Challenged'],
    'Negative': ['Bored', 'Frustrated', 'Concerned', 'Exhausted', 'Stressed', 'Unmotivated', 'Unsuccessful', 'Distracted', 'Hopeless', 'Unsupported']
}

SENTIMENT_COLORS = { 
    'Positive': '#69b3a2', 
    'Neutral': '#779ecb', 
    'Negative': '#e57373',
    'Other': '#cccccc'  
}

WORD_TO_SENTIMENT = {word: sentiment for sentiment, words in SENTIMENT_MAP.items() for word in words}

# --- SCRIPT START ---

CONFIG_PATH = Path.home() / ".student_feedback_analyzer.json"
MAX_URL_HISTORY = 10

def load_config():
    if not CONFIG_PATH.exists():
        return {"url_history": []}
    try:
        with open(CONFIG_PATH) as f:
            data = json.load(f)
        if "url_history" not in data:
            data["url_history"] = []
        return data
    except (json.JSONDecodeError, OSError):
        return {"url_history": []}

def save_config(config):
    try:
        with open(CONFIG_PATH, "w") as f:
            json.dump(config, f, indent=2)
    except OSError:
        pass

def push_url_history(config, url):
    history = [u for u in config.get("url_history", []) if u != url]
    history.insert(0, url)
    config["url_history"] = history[:MAX_URL_HISTORY]
    save_config(config)

def clean_filename(name):
    return "".join(c for c in name if c.isalnum() or c in (' ', '_')).rstrip().replace(" ", "_")

def launch_gui():
    root = tk.Tk()
    root.title("Student Feedback Analyzer")
    root.geometry("650x780")

    config = load_config()
    data_df = [None]
    analysis_mode = tk.StringVar(value="point_in_time")
    selected_period = tk.StringVar()
    periods_list = []
    course_vars = {}
    
    def parse_google_sheet_url(url):
        if "edit?usp=sharing" in url:
            return url.replace("edit?usp=sharing", "export?format=csv")
        elif "/edit" in url:
            base_url = url.split("/edit")[0]
            return f"{base_url}/export?format=csv"
        return url

    def load_data(filepath_or_url, source_value=None):
        try:
            df = pd.read_csv(filepath_or_url, engine='python')
            df.columns = df.columns.str.strip()
            
            required_cols = [CLASS_COL, CLASS_SCORE_COL, OVERALL_SCORE_COL, FEELING_WORD_COL]
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                messagebox.showerror("Error", f"Missing required columns:\n{', '.join(missing_cols)}")
                return
            
            if TIMESTAMP_COL in df.columns:
                df[TIMESTAMP_COL] = pd.to_datetime(df[TIMESTAMP_COL], errors='coerce')
                df = df.dropna(subset=[TIMESTAMP_COL]).sort_values(TIMESTAMP_COL)
                
                df['Time_Gap'] = df[TIMESTAMP_COL].diff()
                df['Period_Group'] = (df['Time_Gap'] > pd.Timedelta(days=7)).cumsum()
                
                period_labels = {}
                for group, group_df in df.groupby('Period_Group'):
                    start_date = group_df[TIMESTAMP_COL].min().strftime('%b %d, %Y')
                    end_date = group_df[TIMESTAMP_COL].max().strftime('%b %d, %Y')
                    label = f"Period {group + 1}: {start_date} - {end_date}"
                    period_labels[group] = label
                    
                df['Survey_Period'] = df['Period_Group'].map(period_labels)
                
                nonlocal periods_list
                periods_list = list(period_labels.values())
                period_listbox.delete(0, tk.END)
                for p in periods_list:
                    period_listbox.insert(tk.END, p)
                period_listbox.selection_set(0)
            else:
                df['Survey_Period'] = "All Data"
                period_listbox.delete(0, tk.END)
                period_listbox.insert(tk.END, "All Data")
                period_listbox.selection_set(0)
            
            data_df[0] = df
            status_label.config(text=f"Loaded {len(df)} responses successfully.", fg="green")

            if source_value:
                push_url_history(config, source_value)
                url_combo['values'] = config.get("url_history", [])
            
            for widget in courses_frame.winfo_children():
                widget.destroy()
            course_vars.clear()
            
            var_all = tk.BooleanVar(value=True)
            course_vars["All Classes Combined"] = var_all
            tk.Checkbutton(courses_frame, text="All Classes Combined (Overall)", variable=var_all).pack(anchor=tk.W)
            
            unique_classes_found = sorted(df[CLASS_COL].dropna().unique())
            for c in unique_classes_found:
                var = tk.BooleanVar(value=True)
                course_vars[c] = var
                tk.Checkbutton(courses_frame, text=c, variable=var).pack(anchor=tk.W)
                
            run_btn.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data:\n{e}")

    def on_browse():
        filepath = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if filepath:
            url_combo.set(filepath)
            load_data(filepath, source_value=filepath)

    def on_load_url():
        url = url_combo.get().strip()
        if url:
            if url.startswith("http"):
                csv_url = parse_google_sheet_url(url)
                load_data(csv_url, source_value=url)
            else:
                load_data(url, source_value=url)
            
    def on_run():
        if data_df[0] is None:
            return
            
        selection = period_listbox.curselection()
        if selection:
            selected_period.set(period_listbox.get(selection[0]))
        else:
            selected_period.set("All Data")
            
        root.withdraw()
        root.update_idletasks()
        root.quit()
        
    # UI Layout
    tk.Label(root, text="Step 1: Load Data", font=("Arial", 14, "bold")).pack(pady=10)
    
    url_frame = tk.Frame(root)
    url_frame.pack(fill=tk.X, padx=20)
    tk.Label(url_frame, text="Google Sheet URL\nor CSV Path:").pack(side=tk.LEFT)
    url_history = config.get("url_history", [])
    url_combo = ttk.Combobox(url_frame, width=38, values=url_history)
    url_combo.pack(side=tk.LEFT, padx=5)
    if url_history:
        url_combo.set(url_history[0])
    tk.Button(url_frame, text="Load Data", command=on_load_url).pack(side=tk.LEFT)
    
    tk.Label(root, text="- OR -").pack(pady=5)
    tk.Button(root, text="Browse for Local CSV", command=on_browse).pack()
    
    status_label = tk.Label(root, text="No data loaded.", fg="gray")
    status_label.pack(pady=5)
    
    tk.Label(root, text="Step 2: Analysis Mode", font=("Arial", 14, "bold")).pack(pady=10)
    
    mode_frame = tk.Frame(root)
    mode_frame.pack()
    tk.Radiobutton(mode_frame, text="Point in Time (Select Period Below)", variable=analysis_mode, value="point_in_time").pack(anchor=tk.W)
    tk.Radiobutton(mode_frame, text="Longitudinal Trends (Compare All Periods)", variable=analysis_mode, value="longitudinal").pack(anchor=tk.W)
    
    tk.Label(root, text="Select Survey Period (for Point in Time):").pack(pady=5)
    period_listbox = tk.Listbox(root, height=5, selectmode=tk.SINGLE, width=60)
    period_listbox.pack()
    
    tk.Label(root, text="Select Courses to Analyze:", font=("Arial", 12, "bold")).pack(pady=10)
    courses_frame = tk.Frame(root)
    courses_frame.pack()
    
    run_btn = tk.Button(root, text="Run Analysis", font=("Arial", 14, "bold"), state=tk.DISABLED, command=on_run)
    run_btn.pack(pady=20)
    
    root.mainloop()
    
    # After mainloop ends
    try:
        root.destroy()
    except Exception:
        pass
    selected_courses_list = [c for c, var in course_vars.items() if var.get()]
    return data_df[0], analysis_mode.get(), selected_period.get(), selected_courses_list

def build_period_date_labels(df):
    """Build a map from Survey_Period -> short date-range label (e.g., 'Oct 3-7' or 'Oct 3')."""
    if TIMESTAMP_COL not in df.columns:
        return {p: p for p in df['Survey_Period'].unique()}

    labels = {}
    for period, group in df.groupby('Survey_Period'):
        ts = group[TIMESTAMP_COL].dropna()
        if ts.empty:
            labels[period] = period
            continue
        start, end = ts.min(), ts.max()
        if start.date() == end.date():
            labels[period] = start.strftime('%b %-d')
        elif start.month == end.month:
            labels[period] = f"{start.strftime('%b %-d')}-{end.strftime('%-d')}"
        elif start.year == end.year:
            labels[period] = f"{start.strftime('%b %-d')}-{end.strftime('%b %-d')}"
        else:
            labels[period] = f"{start.strftime('%b %-d, %Y')}-{end.strftime('%b %-d, %Y')}"
    return labels

def _short_labels_for(period_order, label_map):
    return [label_map.get(p, p) for p in period_order]

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

def create_reference_page(pdf, is_longitudinal):
    import textwrap
    fig, ax = plt.subplots(figsize=(8.5, 11))
    ax.axis('off')
    
    title = "Student Feedback Reports: Reference Guide"
    subtitle = "Point-in-Time Class Analysis" if not is_longitudinal else "Longitudinal Trends Analysis"
    
    # Title
    ax.text(0.05, 0.95, title, fontsize=18, fontweight='bold', color='#1a1a1a', ha='left', va='top')
    ax.text(0.05, 0.91, subtitle, fontsize=12, fontweight='bold', color='#555555', ha='left', va='top')
    ax.plot([0.05, 0.95], [0.89, 0.89], color='#cccccc', linewidth=1)
    
    y = 0.85
    
    # Overview
    overview_text = (
        "This report is generated from student survey feedback data to track experiences, "
        "sentiment, and performance trends. To keep the visualizations clean and highly "
        "impactful, this guide provides a reference explanation for each chart included in "
        "this document."
    )
    wrapped_overview = textwrap.fill(overview_text, width=80)
    ax.text(0.05, y, wrapped_overview, fontsize=10, color='#333333', ha='left', va='top')
    
    y -= 0.12
    
    # Section header
    ax.text(0.05, y, "Understanding the Charts", fontsize=12, fontweight='bold', color='#2a9d8f', ha='left', va='top')
    y -= 0.03
    
    if not is_longitudinal:
        explanations = [
            ("1. Distribution of Student Feelings",
             "A categorical bar chart showing the count of students choosing specific feeling words to "
             "describe their experience. Bars are grouped and color-coded by sentiment: Positive (green), "
             "Neutral (gray), Negative (red), and Other (light gray). Helps gauge the emotional tone of the class."),
            
            ("2. Score Comparison: Class vs. School (Bubble Chart)",
             "Compares student ratings of this class against their overall rating of school on a scale of 1 to 7. "
             "Each bubble represents a score level; its size corresponds to the number of students who gave "
             "that score (labels show the exact count). Muted horizontal dashed lines show the median score "
             "for each group. Allows quick comparison of class satisfaction against general school-wide feelings."),
            
            ("3. Score Distribution by Sentiment",
             "Displays the distribution of class scores (1-7) grouped by student sentiment (Positive, Neutral, Negative). "
             "The boxes show the interquartile range (IQR) and medians, while overlaid bubbles show the individual student "
             "counts at each score. Reveals if students with specific sentiments (e.g. Negative) are also giving lower ratings."),
            
            ("4. Score Correlation: Class vs. School",
             "A 2D scatter plot mapping individual student class ratings (y-axis) against their school ratings (x-axis). "
             "Bubble sizes represent the count of students at that coordinate. The red dashed line shows the linear trend, "
             "and the Pearson correlation coefficient 'r' measures the strength of the relationship. Tells you if a student's "
             "class rating is strongly tied to their general outlook, or if the class experience is independent.")
        ]
    else:
        explanations = [
            ("1. Responses Collected per Period",
             "A simple bar chart displaying the count of responses in each survey period. Provides crucial sample-size "
             "context to determine if trends are driven by a representative number of students."),
            
            ("2. In-Class Score Trends by Course (Small Multiples)",
             "Displays a grid of individual course ratings over time (green line) overlaid on top of the average "
             "across all combined classes (gray line). Allows easy, side-by-side identification of specific courses "
             "that are outperforming or underperforming the department average."),
            
            ("3. Shift in Positive Sentiment: First vs Last Period (Slopegraph)",
             "A slopegraph linking the percentage of positive student sentiment in the first survey period to the "
             "last survey period. Improved sentiment is drawn in green, while declining sentiment is drawn in red. "
             "Provides a high-level view of class sentiment trajectory."),
            
            ("4. Average Score Over Time",
             "Tracks the average class rating (green, solid) compared to the overall school rating (orange, dashed) "
             "across all survey periods. Helps see if class-specific trends are moving in tandem with general school satisfaction."),
            
            ("5. Sentiment Distribution Over Time",
             "A stacked bar chart showing the composition of student sentiment (Positive: green, Neutral: gray, Negative: red, "
             "Other: light gray) for each period. Useful to see if the overall class vibe is shifting over time."),
            
            ("6. Feeling-Word Frequency Heatmap",
             "Maps the frequency of the top 15 feeling words chosen by students over time. Rows are sorted and colored "
             "by sentiment. Darker colors indicate higher frequency. Quickly highlights shifting emotional themes in the class."),
            
            ("7. Score Distributions Over Time (Violin Charts)",
             "Uses violin density curves to show the full distribution of scores for this class (blue) and school (orange) "
             "over time. Box-and-whisker overlays are removed for maximum legibility, and dashed lines represent the mean score "
             "in each period. Reveals skewness, bimodality, and detailed score shifts that simple means might hide.")
        ]
        
    for title_ex, desc in explanations:
        ax.text(0.05, y, title_ex, fontsize=10, fontweight='bold', color='#333333', ha='left', va='top')
        y -= 0.02
        wrapped_desc = textwrap.fill(desc, width=95)
        ax.text(0.07, y, wrapped_desc, fontsize=9, color='#666666', ha='left', va='top')
        y -= 0.012 * len(wrapped_desc.split('\n')) + 0.015
        
    # Footer
    ax.text(0.5, 0.03, "Student Feedback Analyzer • Premium Visualization Engine", fontsize=8, color='#999999', ha='center')
    
    plt.tight_layout()
    pdf.savefig(fig)
    plt.close(fig)

def create_score_trends_chart(class_df, course, period_order, label_map, pdf, png_path=None):
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
    
    ax.set_xticklabels(_short_labels_for(period_order, label_map), rotation=0, ha='center', fontsize=10, color='#333333')
    
    ax.set_title(f"Average Score Over Time: {course}", fontsize=14, pad=20, color='#222222', loc='left')
    ax.set_ylabel("Score (1-7)", fontsize=11, color='#555555', labelpad=10)
    
    ax.legend(frameon=False, loc='upper right', bbox_to_anchor=(1, 1.05), ncol=2, fontsize=10, labelcolor='#444444')
    plt.tight_layout()

    if png_path:
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
    pdf.savefig(fig)
    plt.close(fig)

def create_sentiment_evolution_chart(class_df, course, period_order, label_map, pdf, png_path=None):
    print(f"  > Generating Sentiment Evolution for '{course}'...")
    class_df = class_df.copy()
    class_df['Sentiment'] = class_df[FEELING_WORD_COL].astype(str).str.strip().str.title().map(WORD_TO_SENTIMENT).fillna('Other')

    sentiment_counts = class_df.groupby(['Survey_Period', 'Sentiment'], observed=True).size().unstack(fill_value=0)
    cols_present = [col for col in ['Positive', 'Neutral', 'Negative', 'Other'] if col in sentiment_counts.columns]
    sentiment_counts = sentiment_counts[cols_present]

    sentiment_pct = sentiment_counts.div(sentiment_counts.sum(axis=1), axis=0) * 100
    sentiment_pct = sentiment_pct.reindex(period_order).dropna()
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
    
    ticks = ax.get_yticks()
    ax.set_yticks(ticks)
    ax.set_yticklabels([f"{int(y)}%" for y in ticks], color='#555555')
    
    plt.xticks(rotation=0, ha='center', fontsize=10, color='#333333')
    
    ax.legend(title='', bbox_to_anchor=(1.02, 1), loc='upper left', frameon=False, labelcolor='#444444')
    plt.tight_layout()

    if png_path:
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
    pdf.savefig(fig)
    plt.close(fig)

def create_small_multiples_score_trends(df, classes, period_order, label_map, pdf, png_path=None):
    print("  > Generating Score Trends by Course...")
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

    if png_path:
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)

def create_feeling_words_heatmap(class_df, course, period_order, label_map, pdf, top_n=15, png_path=None):
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

    if png_path:
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
    if pdf:
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

def create_sentiment_slopegraph(df, classes, period_order, label_map, pdf, png_path=None):
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

    if png_path:
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)

def create_score_distribution_by_period(class_df, course, period_order, label_map, pdf, png_path=None):
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

    if png_path:
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
    if pdf:
        pdf.savefig(fig)
    plt.close(fig)

def create_response_volume_bar(df, period_order, label_map, pdf, png_path=None):
    print("  > Generating Response Volume per Period...")
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

    if png_path:
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)

def create_longitudinal_stats_page(df, classes, period_order, label_map, pdf):
    """Hypothesis tests: trend regression, first-vs-last comparison, sentiment chi-square."""
    print("  > Generating Longitudinal Statistical Analysis...")
    from scipy.stats import linregress, ttest_ind, mannwhitneyu, chi2_contingency

    MIN_N = 5
    lines = []

    lines.append("=" * 90)
    lines.append("LONGITUDINAL STATISTICAL ANALYSIS")
    lines.append("=" * 90)
    lines.append("")
    lines.append(f"Periods analyzed ({len(period_order)}): " + ", ".join(label_map.get(p, p) for p in period_order))
    lines.append("")

    # 1. Per-class linear trend: mean score vs period index
    lines.append("-" * 90)
    lines.append("1. LINEAR TREND TEST — Is the in-class score rising or falling over time?")
    lines.append("   Regression of mean in-class score on period index. Significant p-value (<0.05)")
    lines.append("   indicates a reliable upward/downward trend beyond random variation.")
    lines.append("-" * 90)
    lines.append(f"{'Class':<35} {'Slope/period':>13} {'R²':>7} {'p-value':>9}  Interpretation")

    for course in classes:
        class_df = df[df[CLASS_COL] == course]
        period_means = class_df.groupby('Survey_Period')[CLASS_SCORE_COL].mean().reindex(period_order).dropna()
        if len(period_means) < 3:
            lines.append(f"{course[:34]:<35} {'n/a':>13} {'n/a':>7} {'n/a':>9}  Need ≥3 periods with data.")
            continue
        x = np.arange(len(period_means))
        y = period_means.values
        result = linregress(x, y)
        slope, r2, pval = result.slope, result.rvalue ** 2, result.pvalue
        if pval < 0.05:
            direction = "rising" if slope > 0 else "falling"
            interp = f"Significant {direction} trend."
        else:
            interp = "No significant trend (noise)."
        lines.append(f"{course[:34]:<35} {slope:>+13.3f} {r2:>7.2f} {pval:>9.3f}  {interp}")
    lines.append("")

    # 2. First vs last period comparison
    if len(period_order) >= 2:
        first_p, last_p = period_order[0], period_order[-1]
        lines.append("-" * 90)
        lines.append(f"2. FIRST VS LAST PERIOD — Did scores shift from '{label_map.get(first_p, first_p)}' to '{label_map.get(last_p, last_p)}'?")
        lines.append("   Mann-Whitney U when n<15 per group, independent t-test otherwise.")
        lines.append("   Cohen's d shows effect size (0.2=small, 0.5=medium, 0.8=large).")
        lines.append("-" * 90)
        lines.append(f"{'Class':<35} {'n_first':>8} {'n_last':>7} {'mean_Δ':>8} {'p-value':>9} {'d':>6}  Interpretation")

        for course in classes:
            class_df = df[df[CLASS_COL] == course]
            first_scores = class_df[class_df['Survey_Period'] == first_p][CLASS_SCORE_COL].dropna()
            last_scores = class_df[class_df['Survey_Period'] == last_p][CLASS_SCORE_COL].dropna()
            n1, n2 = len(first_scores), len(last_scores)
            if n1 < 2 or n2 < 2:
                lines.append(f"{course[:34]:<35} {n1:>8} {n2:>7} {'n/a':>8} {'n/a':>9} {'n/a':>6}  Insufficient data.")
                continue
            mean_delta = last_scores.mean() - first_scores.mean()
            # Pooled SD for Cohen's d
            pooled_sd = np.sqrt(((n1 - 1) * first_scores.var(ddof=1) + (n2 - 1) * last_scores.var(ddof=1)) / (n1 + n2 - 2))
            d = mean_delta / pooled_sd if pooled_sd > 0 else 0
            if min(n1, n2) < 15:
                try:
                    _, pval = mannwhitneyu(last_scores, first_scores, alternative='two-sided')
                except ValueError:
                    pval = np.nan
            else:
                _, pval = ttest_ind(last_scores, first_scores, equal_var=False)
            if np.isnan(pval):
                interp = "Test undefined (identical values)."
            elif pval < 0.05:
                direction = "increased" if mean_delta > 0 else "decreased"
                interp = f"Significantly {direction}."
            else:
                interp = "No significant shift."
            if min(n1, n2) < MIN_N:
                interp += " [small n — caution]"
            lines.append(f"{course[:34]:<35} {n1:>8} {n2:>7} {mean_delta:>+8.2f} {pval:>9.3f} {d:>+6.2f}  {interp}")
        lines.append("")

        # 3. Sentiment distribution chi-square: first vs last
        lines.append("-" * 90)
        lines.append("3. SENTIMENT SHIFT — Did the mix of positive/neutral/negative change from first to last?")
        lines.append("   Chi-square test of independence on a 2×3 contingency table.")
        lines.append("   Small expected counts (<5) weaken reliability; flagged below.")
        lines.append("-" * 90)
        lines.append(f"{'Class':<35} {'χ²':>8} {'p-value':>9}  Interpretation")

        df_local = df.copy()
        df_local['Sentiment'] = df_local[FEELING_WORD_COL].astype(str).str.strip().str.title().map(WORD_TO_SENTIMENT).fillna('Other')

        for course in classes:
            class_df = df_local[df_local[CLASS_COL] == course]
            first_sent = class_df[class_df['Survey_Period'] == first_p]['Sentiment']
            last_sent = class_df[class_df['Survey_Period'] == last_p]['Sentiment']
            categories = ['Positive', 'Neutral', 'Negative']
            table = np.array([
                [(first_sent == cat).sum() for cat in categories],
                [(last_sent == cat).sum() for cat in categories],
            ])
            if table.sum() == 0 or (table.sum(axis=1) == 0).any():
                lines.append(f"{course[:34]:<35} {'n/a':>8} {'n/a':>9}  Insufficient data.")
                continue
            # Drop all-zero columns to avoid chi2 errors
            nonzero_cols = table.sum(axis=0) > 0
            trimmed = table[:, nonzero_cols]
            if trimmed.shape[1] < 2:
                lines.append(f"{course[:34]:<35} {'n/a':>8} {'n/a':>9}  Only one sentiment category present.")
                continue
            try:
                chi2, pval, _, expected = chi2_contingency(trimmed)
            except ValueError:
                lines.append(f"{course[:34]:<35} {'n/a':>8} {'n/a':>9}  Chi-square undefined.")
                continue
            low_expected = (expected < 5).any()
            if pval < 0.05:
                interp = "Sentiment mix significantly different."
            else:
                interp = "No significant change in sentiment mix."
            if low_expected:
                interp += " [small expected counts — caution]"
            lines.append(f"{course[:34]:<35} {chi2:>8.2f} {pval:>9.3f}  {interp}")
        lines.append("")

    lines.append("=" * 90)
    lines.append("Note: These tests treat each survey response as independent. If the same students")
    lines.append("responded across periods, p-values may be approximate. Use effect sizes + visual")
    lines.append("inspection alongside p-values when making decisions.")
    lines.append("=" * 90)

    text = "\n".join(lines)

    plt.style.use('default')
    fig = plt.figure(figsize=(14, max(8, 0.22 * len(lines))))
    ax = fig.add_subplot(111)
    ax.axis('off')
    ax.text(0.01, 0.99, text, ha='left', va='top', family='monospace', fontsize=8)
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)

def create_longitudinal_reports(df, unique_classes, run_dir, png_dir):
    print("\nGenerating Longitudinal Trend Reports...")
    from matplotlib.backends.backend_pdf import PdfPages

    period_order = list(df.sort_values('Period_Group')['Survey_Period'].unique())
    label_map = build_period_date_labels(df)
    per_class_only = [c for c in unique_classes if c != "All Classes Combined"]

    pdf_filename = os.path.join(run_dir, "Longitudinal_Trends_Report.pdf")
    with PdfPages(pdf_filename) as pdf:
        # Reference Page
        create_reference_page(pdf, is_longitudinal=True)

        if len(period_order) >= 2:
            create_response_volume_bar(df, period_order, label_map, pdf, png_path=os.path.join(png_dir, "responses_volume.png"))
            if per_class_only:
                create_small_multiples_score_trends(df, per_class_only, period_order, label_map, pdf, png_path=os.path.join(png_dir, "courses_small_multiples_score_trends.png"))
                create_sentiment_slopegraph(df, per_class_only, period_order, label_map, pdf, png_path=os.path.join(png_dir, "courses_sentiment_slopegraph.png"))

        for course in unique_classes:
            if course == "All Classes Combined":
                class_df = df.copy()
            else:
                class_df = df[df[CLASS_COL] == course].copy()

            if len(class_df['Survey_Period'].unique()) < 2:
                print(f"  > Skipping longitudinal reports for '{course}' (needs at least 2 time periods).")
                continue

            course_slug = clean_filename(course)
            create_score_trends_chart(class_df, course, period_order, label_map, pdf, png_path=os.path.join(png_dir, f"{course_slug}_score_trends.png"))
            create_sentiment_evolution_chart(class_df, course, period_order, label_map, pdf, png_path=os.path.join(png_dir, f"{course_slug}_sentiment_evolution.png"))
            create_feeling_words_heatmap(class_df, course, period_order, label_map, pdf, png_path=os.path.join(png_dir, f"{course_slug}_feeling_words_heatmap.png"))
            create_score_distribution_by_period(class_df, course, period_order, label_map, pdf, png_path=os.path.join(png_dir, f"{course_slug}_score_distribution.png"))

        if len(period_order) >= 2 and per_class_only:
            create_longitudinal_stats_page(df, per_class_only, period_order, label_map, pdf)

        print(f"  > Saved all longitudinal charts to {pdf_filename}")

def create_consolidated_report(df, unique_classes, qualitative_col_name, run_dir):
    print(f"\nGenerating Consolidated Feedback Report...")

    in_class_stats = pd.DataFrame(index=['Sample Size (n)', 'Mean', 'Median', 'Mode', 'Range', 'Std. Deviation'])
    overall_stats = pd.DataFrame(index=['Sample Size (n)', 'Mean', 'Median', 'Mode', 'Range', 'Std. Deviation'])
    sentiment_stats = pd.DataFrame(index=['% Positive', '% Neutral', '% Negative'])
    feeling_words_data = {}
    qualitative_feedback_data = {}

    df['Sentiment'] = df[FEELING_WORD_COL].map(WORD_TO_SENTIMENT).fillna('Other')

    for course in unique_classes:
        if course == "All Classes Combined":
            class_df = df.copy()
        else:
            class_df = df[df[CLASS_COL] == course].copy()

        class_scores = class_df[CLASS_SCORE_COL].dropna()
        in_class_stats[course] = [
            class_scores.count(),
            f"{class_scores.mean():.2f}",
            class_scores.median(),
            ', '.join(map(str, class_scores.mode().values)) if not class_scores.mode().empty else 'N/A',
            class_scores.max() - class_scores.min() if class_scores.count() > 0 else 'N/A',
            f"{class_scores.std():.2f}"
        ]
        
        overall_scores = class_df[OVERALL_SCORE_COL].dropna()
        overall_stats[course] = [
            overall_scores.count(),
            f"{overall_scores.mean():.2f}",
            overall_scores.median(),
            ', '.join(map(str, overall_scores.mode().values)) if not overall_scores.mode().empty else 'N/A',
            overall_scores.max() - overall_scores.min() if overall_scores.count() > 0 else 'N/A',
            f"{overall_scores.std():.2f}"
        ]
        
        sentiment_perc = class_df['Sentiment'].value_counts(normalize=True).mul(100)
        sentiment_stats[course] = [
            f"{sentiment_perc.get('Positive', 0):.1f}%",
            f"{sentiment_perc.get('Neutral', 0):.1f}%",
            f"{sentiment_perc.get('Negative', 0):.1f}%"
        ]

        word_counts = class_df[FEELING_WORD_COL].value_counts().head(10)
        feeling_words_data[course] = '\n'.join([f"{count} - {word}" for word, count in word_counts.items()])
        
        if qualitative_col_name:
            comments_series = class_df[qualitative_col_name].dropna()
            comments = [str(c).strip() for c in comments_series if str(c).strip() and str(c).lower() not in ['n/a', 'na', 'none']]
            demojized_comments = [emoji.demojize(comment) for comment in comments]
            
            if demojized_comments:
                wrapped_comments = [textwrap.fill(comment, width=100) for comment in demojized_comments]
                qualitative_feedback_data[course] = '\n\n'.join([f"• {comment}" for comment in wrapped_comments])
            else:
                qualitative_feedback_data[course] = "No feedback provided."
        else:
            qualitative_feedback_data[course] = "No feedback provided."

    fig, axes = plt.subplots(6, 1, figsize=(8.5, 32), gridspec_kw={'height_ratios': [2, 2, 1.5, 3, 8, 8]})
    fig.suptitle("Consolidated Student Feedback Report", fontsize=20, y=0.98)

    sections = [
        ("Question 1: 'How are things going in this class?'", in_class_stats, axes[0]),
        ("How are things going overall as a student?", overall_stats, axes[1]),
        ("Sentiment Category Breakdown", sentiment_stats, axes[2])
    ]

    for title, data, ax in sections:
        ax.axis('off')
        ax.set_title(title, fontsize=14, pad=20, weight='bold')
        table = ax.table(cellText=data.values, colLabels=data.columns, rowLabels=data.index, 
                         loc='center', cellLoc='center')
        table.scale(1, 1.8)
        table.auto_set_font_size(False)
        table.set_fontsize(10)

    ax4 = axes[3]
    ax4.axis('off')
    ax4.set_title("Top 10 Feeling Words by Frequency", fontsize=14, pad=20, weight='bold')
    
    num_classes = len(unique_classes)
    x_positions = np.linspace(0.05, 0.95, num_classes*2+1)[1::2]
    
    for x, course in zip(x_positions, unique_classes):
        ax4.text(x, 0.95, course, ha='center', fontsize=11, weight='bold')
        ax4.text(x, 0.9, feeling_words_data[course], ha='center', va='top', fontsize=10)
        
    ax5 = axes[4]
    ax5.axis('off')
    ax5.set_title("Qualitative Feedback", fontsize=14, pad=20, weight='bold')

    full_text = ""
    for course in unique_classes:
        full_text += f'{course}\n'
        full_text += f'{"-"*len(course)}\n'
        full_text += f'{qualitative_feedback_data.get(course, "No feedback provided.")}\n\n'
    
    ax5.text(0.01, 0.95, full_text.strip(), ha='left', va='top', fontsize=9, wrap=True)
    
    ax6 = axes[5]
    ax6.axis('off')
    ax6.set_title("Statistical Significance Testing", fontsize=14, pad=20, weight='bold')
    
    def cohen_d_paired(x, y):
        n = len(x)
        if n < 2:
            return 0, "N/A"
        
        diff = x - y
        mean_diff = np.mean(diff)
        std_diff = np.std(diff, ddof=1)
        
        if std_diff == 0:
            return np.inf if mean_diff != 0 else 0, "N/A"
            
        d = mean_diff / std_diff
        
        magnitude = "N/A"
        if abs(d) < 0.2:
            magnitude = "Negligible"
        elif abs(d) < 0.5:
            magnitude = "Small"
        elif abs(d) < 0.8:
            magnitude = "Medium"
        else:
            magnitude = "Large"
            
        return d, magnitude

    MIN_SAMPLE_SIZE_T_TEST = 15
    unique_classes_for_stats = [c for c in unique_classes if c != "All Classes Combined"]
    
    test_results_text = "--- ANOVA: Comparing Mean 'In-Class' Scores Between Classes ---\n"
    test_results_text += "\n1. 'How are things going in this class?'\n"
    class_scores_grouped = [df[df[CLASS_COL] == course][CLASS_SCORE_COL].dropna() for course in unique_classes_for_stats]
    
    valid_groups = [g for g in class_scores_grouped if len(g) > 1]
    
    if len(valid_groups) > 1:
        f_stat, p_val = f_oneway(*valid_groups)
        test_results_text += f"   - ANOVA Result: p-value = {p_val:.3f}\n"
        if p_val < 0.05:
            test_results_text += "   - Interpretation: There is a statistically significant difference in the average 'in-class'\n     scores between at least two of your classes.\n"
            
            all_scores_list = []
            valid_class_names = [course for course, g in zip(unique_classes, class_scores_grouped) if len(g) > 1]
            
            for course_name, s in zip(valid_class_names, valid_groups):
                all_scores_list.append(pd.DataFrame({'score': s, 'group': course_name}))

            if all_scores_list:
                tukey_data = pd.concat(all_scores_list)
                if len(tukey_data['group'].unique()) > 1:
                    tukey_result = pairwise_tukeyhsd(endog=tukey_data['score'], groups=tukey_data['group'], alpha=0.05)
                    test_results_text += "   - Pairwise Comparison (Tukey's HSD): The table below shows which specific classes differ.\n     If 'reject' is True, their means are significantly different.\n"
                    test_results_text += str(tukey_result) + "\n"
                else:
                    test_results_text += "   - Pairwise Comparison: Not enough groups with data for Tukey's HSD test.\n"
        else:
            test_results_text += "   - Interpretation: The differences in the average 'in-class' scores between your classes are\n     not statistically significant and could be due to random chance.\n"
    else:
        test_results_text += "   - ANOVA Result: Not enough data (at least 2 classes with >1 student) to perform this test.\n"
    
    test_results_text += "\n--- Paired T-Test: Comparing 'In-Class' vs 'Overall' Scores Within Each Class ---\n"
    for course in unique_classes_for_stats:
        class_df_test = df[df[CLASS_COL] == course][[CLASS_SCORE_COL, OVERALL_SCORE_COL]].dropna()
        n = len(class_df_test)
        
        test_results_text += f"\n- {course} (n={n}):\n"
        
        if n > 1:
            if n < MIN_SAMPLE_SIZE_T_TEST:
                 test_results_text += f"  - Warning: Sample size is small. Results should be interpreted with caution.\n"

            t_stat, p_val = ttest_rel(class_df_test[CLASS_SCORE_COL], class_df_test[OVERALL_SCORE_COL])
            test_results_text += f"  - T-Test Result: p-value = {p_val:.3f}\n"
            
            d, magnitude = cohen_d_paired(class_df_test[CLASS_SCORE_COL], class_df_test[OVERALL_SCORE_COL])
            test_results_text += f"  - Effect Size (Cohen's d): {d:.2f} ({magnitude} Effect)\n"

            mean_class = class_df_test[CLASS_SCORE_COL].mean()
            mean_overall = class_df_test[OVERALL_SCORE_COL].mean()
            
            if p_val < 0.05:
                direction = "higher" if mean_class > mean_overall else "lower"
                test_results_text += f"  - Interpretation: Students in this class rate their 'in-class' experience (mean={mean_class:.2f})\n    significantly {direction} than their 'overall' school experience (mean={mean_overall:.2f}).\n"
            else:
                test_results_text += "  - Interpretation: There is no statistically significant difference between how students rate\n    this class and how they rate their school experience overall.\n"
        else:
            test_results_text += "  - T-Test Result: Not enough data (n < 2) to perform paired t-test.\n"
            
    ax6.text(0.01, 0.95, test_results_text.strip(), ha='left', va='top', fontsize=8, family='monospace')

    fig.tight_layout(rect=[0, 0, 1, 0.96])
    
    filename = os.path.join(run_dir, "Consolidated_Feedback_Report.pdf")
    plt.savefig(filename)
    plt.close()
    print(f"  > Saved as {filename}")

def create_feelings_histogram(df, course_name, pdf=None, png_path=None):
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
    if png_path:
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
    if pdf:
        pdf.savefig(fig)
    plt.close(fig)

def create_comparative_bubble_chart(df, course_name, pdf=None, png_path=None):
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

    setup_tufte_axes(ax, ygrid=True)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.tick_params(axis='x', length=0)
    ax.set_yticks(range(1, 8))

    ax.set_title(f"Score Comparison: Class vs. School (n={len(class_df)})", fontsize=14, pad=20, color='#222222', loc='left')
    ax.set_ylabel("Score (1: Not well – 7: Very well)", fontsize=11, color='#555555')
    ax.set_xlabel(None)
    ax.set_ylim(0.5, 7.5)

    plt.tight_layout()
    if png_path:
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
    if pdf:
        pdf.savefig(fig)
    plt.close(fig)


def create_sentiment_score_boxplot(df, course_name, pdf=None, png_path=None):
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
    if png_path:
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
    if pdf:
        pdf.savefig(fig)
    plt.close(fig)

def create_score_correlation_plot(df, course_name, pdf=None, png_path=None):
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
    if png_path:
        fig.savefig(png_path, dpi=300, bbox_inches='tight')
    if pdf:
        pdf.savefig(fig)
    plt.close(fig)

def show_completion_popup(output_files, run_dir):
    """Modal popup shown when analysis finishes. User clicks OK to exit."""
    import os
    import sys
    import subprocess
    popup = tk.Tk()
    popup.title("Analysis Complete")
    popup.geometry("650x500")

    tk.Label(popup, text="✓ Analysis complete", font=("Arial", 16, "bold"), fg="#2e8b57").pack(pady=(15, 5))
    tk.Label(popup, text=f"Output directory:\n{run_dir}", font=("Arial", 10), justify=tk.LEFT).pack(pady=5)
    tk.Label(popup, text=f"{len(output_files)} files written:", font=("Arial", 11, "bold")).pack(pady=(10, 5))

    list_frame = tk.Frame(popup)
    list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Menlo", 10))
    
    sorted_files = sorted(output_files)
    for f in sorted_files:
        listbox.insert(tk.END, os.path.relpath(f, run_dir))
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=listbox.yview)

    def open_selected(event=None):
        selection = listbox.curselection()
        if not selection: return
        rel_path = listbox.get(selection[0])
        file_path = os.path.join(run_dir, rel_path)
        if sys.platform == "win32":
            os.startfile(file_path)
        elif sys.platform == "darwin":
            subprocess.call(["open", file_path])
        else:
            subprocess.call(["xdg-open", file_path])

    listbox.bind("<Double-1>", open_selected)
    tk.Label(popup, text="(Double-click any file above to open it)", font=("Arial", 9, "italic"), fg="gray").pack(pady=5)

    btn_frame = tk.Frame(popup)
    btn_frame.pack(pady=15)
    
    # Check if there is a main PDF to offer a direct open button
    pdfs = [f for f in sorted_files if f.endswith('.pdf')]
    if pdfs:
        def open_pdf():
            file_path = pdfs[0]
            if sys.platform == "win32": os.startfile(file_path)
            elif sys.platform == "darwin": subprocess.call(["open", file_path])
            else: subprocess.call(["xdg-open", file_path])
        tk.Button(btn_frame, text="Open PDF Report", font=("Arial", 12, "bold"), fg="#2e8b57", width=15, command=open_pdf).pack(side=tk.LEFT, padx=10)

    def on_ok():
        try:
            popup.withdraw()
            popup.update_idletasks()
            popup.destroy()
        except Exception:
            pass

    tk.Button(btn_frame, text="OK", font=("Arial", 12, "bold"), width=10, command=on_ok).pack(side=tk.LEFT, padx=10)
    popup.mainloop()

if __name__ == "__main__":
    import os
    import time
    result = launch_gui()
    if len(result) == 4:
        df, analysis_mode, selected_period, selected_courses = result
    else:
        df = None

    if df is not None:
        print(f"\nProcessing loaded dataset...")
        run_start = time.time() - 1  # small buffer for filesystem timestamp resolution

        df[FEELING_WORD_COL] = df[FEELING_WORD_COL].astype(str).str.strip().str.title()

        qualitative_col_name = None
        try:
            qualitative_col_name = [col for col in df.columns if col.startswith('[Optional]')][0]
        except IndexError:
            pass

        unique_classes = selected_courses

        if not unique_classes:
            print("\nError: No classes found in the specified class column.")
        else:
            # Set up output directories recursively inside 'reports/'
            reports_root = os.path.join(os.getcwd(), "reports")
            os.makedirs(reports_root, exist_ok=True)
            
            if analysis_mode == "longitudinal":
                run_dir = os.path.join(reports_root, "longitudinal")
            else:
                period_slug = clean_filename(selected_period).lower() if selected_period else "all_data"
                run_dir = os.path.join(reports_root, period_slug)
                
            os.makedirs(run_dir, exist_ok=True)
            png_dir = os.path.join(run_dir, "pngs")
            os.makedirs(png_dir, exist_ok=True)

            if analysis_mode == "longitudinal":
                create_longitudinal_reports(df, unique_classes, run_dir, png_dir)
            else:
                if selected_period and selected_period != "All Data":
                    df = df[df['Survey_Period'] == selected_period]
                    print(f"\nFiltering data for {selected_period}...")

                print(f"Found {len(unique_classes)} classes to analyze.")
                from matplotlib.backends.backend_pdf import PdfPages

                pdf_filename = os.path.join(run_dir, "Class_Level_Charts.pdf")
                with PdfPages(pdf_filename) as pdf:
                    # Create Reference Page
                    create_reference_page(pdf, is_longitudinal=False)
                    
                    for course in unique_classes:
                        print(f"Processing '{course}'...")
                        course_slug = clean_filename(course)
                        create_feelings_histogram(df, course, pdf=pdf, png_path=os.path.join(png_dir, f"{course_slug}_feelings_histogram.png"))
                        create_comparative_bubble_chart(df, course, pdf=pdf, png_path=os.path.join(png_dir, f"{course_slug}_comparative_bubble_chart.png"))
                        create_sentiment_score_boxplot(df, course, pdf=pdf, png_path=os.path.join(png_dir, f"{course_slug}_sentiment_boxplot.png"))
                        create_score_correlation_plot(df, course, pdf=pdf, png_path=os.path.join(png_dir, f"{course_slug}_score_correlation.png"))
                        print("-" * 20)

                print(f"  > Saved all class-level charts to {pdf_filename}")

                if qualitative_col_name:
                    create_consolidated_report(df, unique_classes, qualitative_col_name, run_dir)
                else:
                    print("\nSkipping consolidated report because no qualitative feedback column was found.")

            print("\nAnalysis complete. All requested plots and reports have been saved in this directory.")

            output_files = []
            for root, dirs, files in os.walk(run_dir):
                for fname in files:
                    fpath = os.path.join(root, fname)
                    if fname.lower().endswith(('.pdf', '.png')):
                        if os.path.getmtime(fpath) >= run_start:
                            output_files.append(fpath)

            show_completion_popup(output_files, run_dir)
    else:
        print("\nScript manually closed or no file was processed.")
