# EduScripts

A comprehensive repository of Google Apps Scripts and Python analytical tools designed for educators, classroom automation, gradebook reporting, and educational data visualization.

---

## 📚 Repository Overview

This repository houses standalone tools organized into modular subdirectories. Tools are split into two primary categories:

1. **Google Apps Scripts (`.gs`)**: Embedded automation scripts that run directly inside Google Sheets or Google Docs to streamline administrative, reporting, and instructional tasks.
2. **Python Data Science Applications (`.py`)**: Powerful data analysis and visualization suites built with pandas, seaborn, and matplotlib to analyze student survey feedback and AP exam correlation data using Edward Tufte visual design principles.

---

## 🛠️ Tool Directory Index

| Tool Directory | Type | Key Features & Purpose |
| :--- | :--- | :--- |
| [**`APExamScoreGraphs`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/APExamScoreGraphs/README.md) | Python | Analyzes correlation between in-class test averages and AP exam scores. Generates Tufte-style scatter plots, logistic regression probability models (AP 3+, 4+, 5), Tukey HSD post-hoc statistics, and multi-page PDF reports. |
| [**`Calendarizer`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/Calendarizer/README.md) | Apps Script | Transforms instructional pacing spreadsheets into printable "Wall Calendar" (monthly) or "Lateral Calendar" (horizontal linear) views with automatic holiday/block highlighting. |
| [**`Feedback Reports`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/Feedback Reports/README.md) | Python + GUI | Premium student survey feedback analyzer. Generates point-in-time diagnostic booklets, longitudinal trend dashboards, sentiment slopegraphs, and 300 DPI high-resolution chart images. |
| [**`GradebookReporter`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/GradebookReporter/README.md) | Apps Script | Generates individual student progress reports directly from a spreadsheet gradebook. Supports HTML email summaries to students/parents and compiled Google Docs for printing. |
| [**`Randomizer`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/Randomizer/README.md) | Apps Script | Utility to randomize student lists, construct balanced lab groups, create presentation orders, and shuffle seating charts directly within Google Sheets. |
| [**`Schedule Tools`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/Schedule Tools/README.md) | Apps Script | Utilities for managing rotational school schedules, period time calculations, and daily timetable formatting. |
| [**`Miscellaneous Tools`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/Miscellaneous%20Tools/README.md) | Mixed | Collection of standalone utility scripts and single-purpose educational tools. |

---

## 📖 Best Practices: Setting Up & Running Google Apps Scripts

For scripts in `Calendarizer`, `GradebookReporter`, `Randomizer`, and `Schedule Tools`:

### 1. Installation into Google Sheets
1. Open the Google Sheet where you want to use the tool.
2. In the top menu, click **Extensions** > **Apps Script**.
3. Clear out any sample code in the default `Code.gs` file.
4. Copy the entire contents of the relevant `.gs` file (e.g., `ReportScript.gs`, `calendarizer.gs`, `Randomizer.gs`, or `ScheduleTools.gs`).
5. Paste the code into the Apps Script editor and press `Cmd + S` (macOS) or `Ctrl + S` (Windows) to save.
6. Refresh your Google Sheet tab. A custom menu (e.g., **Gradebook Tools**, **Calendar Tools**) will appear in the Google Sheets toolbar upon reloading.

### 2. Google Authorization Walkthrough
The first time you select a action from the custom script menu, Google requires a one-time authorization step:
1. Click **Continue** when prompted with *Authorization Required*.
2. Select your Google Account.
3. If presented with a *"Google hasn't verified this app"* screen, click **Advanced** at the bottom.
4. Click **Go to [Script Name] (unsafe)**.
5. Click **Allow** to grant permission for the script to access your spreadsheet and run.

### 3. Setting Up Triggers (Automation)
Some scripts benefit from automated background triggers (e.g., processing Form responses on submission or daily attendance checks):
1. In the Apps Script editor sidebar, click the **Triggers** icon (the alarm clock).
2. Click **+ Add Trigger** (bottom right).
3. Select the function you want to run automatically.
4. Choose the event type:
   - **On edit**: Runs whenever a user edits a cell in the spreadsheet.
   - **On form submit**: Runs automatically whenever a linked Google Form receives a new response.
   - **Time-driven**: Runs on a recurring schedule (e.g., every morning at 7:00 AM).

---

## 🐍 Best Practices: Python Data Analysis Tools

For Python tools in `APExamScoreGraphs` and `Feedback Reports`:

### Environment Prerequisites
Ensure Python 3.9+ is installed along with common data science libraries:
```bash
pip install pandas matplotlib seaborn numpy scipy statsmodels scikit-learn emoji
```

### Direct Google Sheets CSV Integration
Both Python analyzers support pulling live data directly from Google Sheets without manually downloading CSV files every time:
1. Open your Google Sheet containing survey or assessment data.
2. Click **Share** > change permissions to *"Anyone with the link can view"*.
3. Copy the URL (e.g. `https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit?usp=sharing`).
4. Paste the URL into the application GUI or script prompt; the analyzer automatically converts the edit link to an export feed (`/export?format=csv`) to pull fresh data dynamically.

---

## 🔒 Data Privacy & GitHub Guidelines

> [!IMPORTANT]
> **Protecting Student PII (Personally Identifiable Information)**:
> This repository is configured with a strict root `.gitignore` to ensure no confidential student data, grade records, or survey responses are ever pushed to GitHub.

- **Excluded File Types**: All `*.csv`, `*.pdf`, `*.png`, and `*.svg` files are ignored globally across all directories.
- **Cache Exclusion**: Python byte-caches (`__pycache__/`, `*.pyc`) and editor metadata (`.DS_Store`, `.claude/`, `.ipynb_checkpoints/`) are ignored automatically.
- **Commit Safety Check**: Before pushing changes, verify that your staging area contains only script source code (`.gs`, `.py`, `.md`) using:
  ```bash
  git status
  ```
