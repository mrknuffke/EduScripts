# Student Feedback Analyzer

**Version**: 31.May.2026

## Overview
Student Feedback Analyzer is a premium Python application designed to analyze, visualize, and report student survey feedback data. Drawing heavy visual design inspiration from Edward Tufte's principles of data presentation (maximizing data-ink ratio, removing chart junk, and utilizing muted, harmonious color palettes), it transforms student surveys into beautiful, diagnostic, self-explanatory PDF booklets and presentation-ready PNGs.

The tool provides both **Point-in-Time** (deep dive into a specific survey period) and **Longitudinal** (chronological trend tracking across multiple periods) analysis modes, backed by robust statistical calculations.

---

## Features

### 1. Smart Data Loading & GUI
- **User-Friendly Tkinter Interface**: Quick setup using a file selector, URL combo box, and toggle switches.
- **Google Sheet Syncing**: Parses sharing URLs to automatically download direct CSV feeds.
- **Smart Timeline Chunking**: Groups survey timestamps into chronological "Survey Periods" automatically if they are spaced apart by more than a week.
- **Configuration Cache**: Stores your URL history locally in `~/.student_feedback_analyzer.json` for rapid data reloading.

### 2. High-Resolution Visualizations (Tufte-inspired Aesthetics)
All charts are rendered with minimalist frames, custom grid lines, curated colors, and **self-documenting descriptions** embedded at the bottom of the canvas:
- **Distribution of Student Feelings**: Dynamic histogram tracking student mood and sentiment (Positive, Neutral, Negative).
- **Class vs. School Overall**: Bubble scatterplot comparing student feedback inside your class against their overall school satisfaction.
- **Score Distribution by Sentiment**: Boxplots overlaid with bubble counts to map satisfaction ratings against student emotional states.
- **Score Correlation**: 2D scatterplot plotting class vs. school satisfaction, showing trend regression lines and Pearson correlation $r$.
- **Longitudinal Trend Dashboard**: Multi-period trend analysis featuring:
  - *Response Volume*: Count of responses across periods.
  - *In-Class Trends (Small Multiples)*: Grids showing individual course trends overlaid on the combined class average.
  - *Sentiment Slopegraph*: Direct line comparison of positive sentiment shifting between the first and last periods.
  - *Average Score Over Time*: Class rating line tracked alongside school rating line.
  - *Sentiment Evolution*: Stacked composition bars showing sentiment shifts over time.
  - *Feeling-Word Heatmap*: Frequency map tracking shift in selected feeling words.
  - *Score Violins Over Time*: Violin distribution curves showing full density shifts.

### 3. Integrated Statistical Insights
- Paired statistical analyses to check for significant differences (Pearson $r$, Cohen's $d$ effect sizes).
- Multi-period trend hypothesis testing (linear regression slopes, Mann-Whitney U tests, independent t-tests, and Chi-square sentiment shifts).
- **Consolidated Textual Report**: A formatted Markdown/text document compiling numerical statistics, top-10 feeling words, and categorizing qualitative student comments.
- **Master Reference Guide**: A dedicated explanation sheet generated inside your report directory to help you interpret every single chart.

---

## Intake Spreadsheet Structure

To successfully analyze your data, the input Google Sheet or CSV must follow a precise header scheme. The script matches columns by their exact header text:

| Dimension / Column | Required Header Name | Data Format & Expectations |
| :--- | :--- | :--- |
| **Submission Time** | `Timestamp` | Datetime string (e.g., `YYYY-MM-DD HH:MM:SS` or standard form date). Mandatory for *Longitudinal Trends* to segment surveys into weekly periods. |
| **Course Indicator** | `Which of my classes are you in?` | Text string indicating the specific course or class period (e.g. `AP Biology`, `Chemistry Period 1`). |
| **In-Class Rating** | `Question 1: How are things going for you in this class?` | Numeric integer rating between `1` (lowest satisfaction) and `7` (highest satisfaction). |
| **Overall School Rating** | `How are things going overall as a student across all of your classes)?` | Numeric integer rating between `1` and `7`. *(Note: Make sure to include the closing parenthesis `)` before the question mark!)* |
| **Feeling Word Descriptor** | `Question 2: If you had to choose a single word to describe your experience in this class right now as a student, what would it be?` | A single word capturing their mood. The script maps responses to the following sentiments:<br>• **Positive**: *Affirmed, Calm, Comfortable, Hopeful, Motivated, Supported, Purposeful, Curious, Proud*<br>• **Neutral**: *Challenged*<br>• **Negative**: *Bored, Frustrated, Concerned, Exhausted, Stressed, Unmotivated, Unsuccessful, Distracted, Hopeless, Unsupported* |
| **Qualitative Comments** | *Any column header starting with* `[Optional]` | Text strings containing open-ended student notes (e.g. `[Optional] Is there anything else you want to share?`). These comments are grouped by course and exported in the final report text summary. |

---

## Installation & Requirements

The analyzer requires **Python 3** and several open-source data science libraries.

1. Clone or download this project into your workspace.
2. Install the required Python dependencies:
   ```bash
   pip install pandas matplotlib seaborn numpy emoji scipy statsmodels
   ```
3. Verify that your system has `tkinter` support installed (included by default on macOS and Windows).

---

## Usage

1. Open your terminal in the `Feedback Reports` directory.
2. Launch the analyzer:
   ```bash
   python3 student_feedback_analyzer.py
   ```
3. Use the GUI:
   - **Load Data**: Paste your Google Sheet URL or click **Browse for Local CSV** to select a CSV.
   - **Configure Analysis**: Select **Point in Time** or **Longitudinal Trends**.
   - **Filter**: Choose the specific survey periods and courses to include.
   - **Analyze**: Click **Run Analysis**.
4. **View Outputs**: The analyzer generates a new folder inside the `reports/` directory:
   - `Class_Level_Charts.pdf` or `Longitudinal_Trends_Report.pdf` (comprehensive PDF reports).
   - `Consolidated_Feedback_Report.pdf` (compiling stats and qualitative comment lists).
   - `Report_Reference_Guide.pdf` (master guide defining all charts).
   - `pngs/` sub-folder containing 300 DPI high-resolution standalone images of every single chart, perfect for presentations or emails.

---

## Data Privacy Focus
This repository contains a local `.gitignore` configured to keep your student data private:
- Automatically ignores generated directories under `reports/`.
- Automatically ignores all `*.csv`, `*.pdf`, and `*.png` files from git indexing to prevent accidental commits of sensitive student records.
