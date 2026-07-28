# AP Exam Score Graphs & Correlation Analyzer

A Python analytical tool designed to evaluate the relationship between student in-class test averages and their ultimate AP Exam scores. Built around Edward Tufte visual design principles, `MasterAnalysis.py` cleans raw student assessment data, flags grade mismatches interactively, runs advanced statistical models (ANOVA, Tukey HSD, Logistic Regression), and generates high-resolution chart publications and multi-page PDF reports.

---

## 📊 Features & Analysis Pipeline

1. **Interactive Data Validation**:
   - Compares reported letter grades (`Test Average Grade`) against calculated test averages based on a 10-point scale.
   - Interactively flags mismatches and offers four resolution modes: auto-correct, step-by-step interactive review, dataset exclusion, or continuing as-is.
   - Automatically exports flagged discrepancies to `flagged_for_review.csv`.

2. **Core Visualizations**:
   - **Figure 1 — Correlation Scatterplot**: Translucent frequency bubble plot with regression trend line, 95% confidence band, Pearson $r$, and $R^2$ coefficient.
   - **Figure 2 — Frequency Bubble Chart**: Frequency distribution mapping letter grades ($D$ to $A+$) against AP scores ($1$ to $5$).
   - **Figure 3 — Descriptive Statistics Table**: N, Mean, Median, Range, and SD formatted table.
   - **Figure 4 — Probability Heatmap**: Conditional probabilities of achieving each AP score based on test average ranges.
   - **Figure 5 — Year-over-Year Performance Trends**: Dual stacked trend panels with 95% confidence intervals, ANOVA test statistics, and automated Tukey HSD post-hoc pairwise comparisons.
   - **Figure 6 — Distribution Boxplots**: Tufte minimalist boxplots overlaid with jitterless individual data points.
   - **Figure 7 — "3 or Above" Passing Rates**: Side-by-side bar charts showing passing rates ($AP \ge 3$) by letter grade and school year.
   - **Figure 8 — Logistic Regression Predictive Model**: Multi-threshold logistic regression modeling the probability of achieving AP scores of $\ge 3$, $\ge 4$, and $= 5$ based on in-class test averages, complete with accuracy metrics and threshold cutoff points ($50\%$, $75\%$, $90\%$).
   - **Figure 9 — Predicted Probabilities Grid**: Cell-shaded probability grid for test averages in 5-point increments.
   - **Figure 10 — Section Breakdown Table**: Detailed breakdown by school year and section block.

---

## 📥 Required CSV Data Structure

The script reads data from `AP Test Score-Exam Score Tracking - Data.csv` (or prompts for a file path). The input spreadsheet must include the following headers:

| Column Name | Type | Expected Values & Format | Description |
| :--- | :--- | :--- | :--- |
| `Student` | String | Student identifier (name or ID) | Identifies the individual student. |
| `Schoolyear` | String | e.g., `2022-2023`, `2023-2024` | The academic school year. |
| `Block` | String | e.g., `Period 1`, `Block A` | Class section identifier. |
| `Test Average` | Float | `0.0` to `100.0` | Student's average score on in-class tests. |
| `Test Average Grade` | String | `A+`, `A`, `B+`, `B`, `C+`, `C`, `D+`, `D` | Reported letter grade corresponding to the test average. |
| `AP Exam Score` | Integer | `1`, `2`, `3`, `4`, `5` | The score earned on the official AP Exam. |

---

## 📐 Grade Scale Thresholds (10-Point Scale)

The script uses the following threshold boundaries for expected grade validation:

- **A+**: $> 90.0$
- **A**: $> 80.0$ and $\le 90.0$
- **B+**: $> 70.0$ and $\le 80.0$
- **B**: $> 60.0$ and $\le 70.0$
- **C+**: $> 50.0$ and $\le 60.0$
- **C**: $> 40.0$ and $\le 50.0$
- **Below C**: $\le 40.0$

---

## 🚀 Execution & Usage

1. Open a terminal in the `APExamScoreGraphs` directory:
   ```bash
   cd APExamScoreGraphs
   ```
2. Run the analysis script:
   ```bash
   python3 MasterAnalysis.py
   ```
3. If grade mismatches are detected, follow the interactive terminal prompts:
   - Type `fix` to auto-correct all letter grades.
   - Type `review` to step through each student mismatch manually.
   - Press `Enter` to proceed with the raw data as-is.
4. **Outputs**: Generated PDF reports and charts are compiled for review.

---

## 🔒 Data Privacy Note

All `*.csv` data files (containing raw student grades) and output plots/PDFs are automatically ignored by Git via `.gitignore`. Never commit raw student scores to version control.
