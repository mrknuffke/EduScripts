# AP Exam Score Predictor & Analytics Dashboard (`v.26-27`)

An interactive, zero-install single-page web application designed for educators, students, and counselors to determine and justify **predicted AP Exam grades for UCAS applications** (and general university advising) based on in-class test performance, visualize historical score correlations, evaluate probability distributions, and identify target growth gaps.

---

## 🏫 Institutional Context: Singapore American School

This model is calibrated specifically on historical student performance data from **Singapore American School (SAS)** ($N = 193$ students across the 2020–2021 through 2025–2026 academic school years). The assessment pacing, difficulty, grading standards, and student demographics reflect the rigor of the SAS curriculum.

> [!WARNING]
> ### ⚠️ Provisional Nature of Score Predictions
> - **Statistical Estimate, Not a Guarantee**: Output scores and probability bands are **strictly provisional estimates** generated from ordinary least squares (OLS) linear regression ($R^2 = 0.4527$, $\text{RMSE} = 0.744$). They do **not** represent a guarantee or contract regarding official College Board AP Exam scores.
> - **UCAS Predicted Grades Advisory Role**: When used to establish or corroborate predicted AP grades for **UCAS (UK university applications)**, predictions reflect a student's observed achievement trajectory at a specific point in time. Final AP results in May depend heavily on second-semester conceptual retention, dedicated cumulative review, pacing on Free Response Questions (FRQs), and testing conditions on exam day.
> - **In-Progress Coursework**: Early-semester test averages are drawn from small sample sizes (1–2 unit assessments) and carry greater variance. Predicted grades should always be interpreted alongside professional educator judgment, qualitative engagement, and historical exam prep rigor.

---

## 🎯 What Is This Tool?

The **AP Exam Score Predictor & Analytics Dashboard** provides an empirical, defensible standard for projecting student AP Exam outcomes—specifically tailored for submitting **provisional predicted grades for UCAS (UK university admissions)** and academic counseling at Singapore American School.

Trained on a multi-year SAS cohort ($N = 193$ students across the 2020–2026 school years) and versioned annually for the **2026–2027 academic cycle (`v.26-27`)**, the application translates a student's current in-class test average into:
1. **A continuous predicted AP score** (e.g., $3.82 / 5.00$).
2. **A rounded discrete AP score tier** (1 through 5) with College Board qualification badges.
3. **A Gaussian probability distribution** estimating the percentage likelihood of scoring each integer value (1, 2, 3, 4, or 5).
4. **Target score gap analysis** showing the exact percentage increase required to advance to the next AP qualification tier.
5. **Actionable instructional guidance** tailored to the student's risk profile.
6. **An interactive historical scatter plot** with year and block filtering to explore cohort trends against the empirical regression model.

The tool runs **100% locally in the browser** without requiring Python, servers, build tools, or command-line execution.

---

## 🚀 Key Features

### 1. Interactive Score Calculator
* **Bi-directional Controls**: Adjust student test averages using either a direct numeric entry field or a real-time slider.
* **Quick Presets**: Jump instantly to standard benchmark averages ($45\%$, $60\%$, $75\%$, $85\%$, $92\%$).
* **Hero Score Badge**: Displays the predicted discrete AP score along with official College Board descriptors (*Extremely Well Qualified*, *Well Qualified*, *Qualified*, *Possibly Qualified*, *No Recommendation*).
* **Continuous Score Precision**: Shows fine-grained continuous scores (clamped between $1.00$ and $5.00$) to evaluate where a student sits within a given score band.

### 2. Gaussian Probability Distribution
* Rather than providing only a single deterministic prediction, the tool models the reality of testing variance using a standard normal distribution around the predicted continuous mean.
* Uses the empirical model's Root Mean Squared Error ($\sigma = 0.744$) to calculate cumulative distribution function (CDF) probabilities for bins:
  * **Score 1**: $< 1.5$
  * **Score 2**: $[1.5, 2.5)$
  * **Score 3**: $[2.5, 3.5)$
  * **Score 4**: $[3.5, 4.5)$
  * **Score 5**: $\ge 4.5$
* Dynamic color-coded probability bars identify the most likely outcome while making confidence bounds transparent.

### 3. Score Advancement Gap Analysis
* Automatically computes the remaining margin needed on in-class tests to reach higher benchmarks:
  * **AP 3 Cutoff (Passing / Credit)**: $\ge 52.1\%$
  * **AP 4 Cutoff (Well Qualified)**: $\ge 69.5\%$
  * **AP 5 Cutoff (Top Tier)**: $\ge 87.0\%$
* When a student surpasses $87.0\%$, the panel confirms top-tier status and displays their buffer above the cutoff.

### 4. Historical Scatter Plot & Analytics
* **Cohort Visualization**: Plots student records against the linear regression trendline ($y = 0.05731x - 0.4851$).
* **Filter Controls**: Slice historical data by academic school year (2020–2021 through 2025–2026) or class block (A2, A3, A4, B1, B2, B3).
* **Cohort Metrics Summary**:
  * **Sample Size**: $N = 193$ students
  * **Coefficient of Determination ($R^2$)**: $0.4527$ (explaining $\sim 45.3\%$ of AP score variance)
  * **Residual Error (RMSE)**: $0.744$ score units
  * **$\pm 1$ Score Accuracy**: $94.3\%$ ($54.9\%$ exact integer match)

### 5. Empirical Cutoff Benchmarks & Instructional Matrix
Provides an educator-facing reference matrix connecting in-class averages with targeted pedagogical interventions:

| AP Score Tier | Required Test Average | College Credit Status | Instructional & Review Strategy |
| :---: | :---: | :---: | :--- |
| **AP 5** | $\ge 87.0\%$ | Universal Placement & Credit | Maintain pacing; refine timing on full-point Free Response Question (FRQ) rubrics. |
| **AP 4** | $69.5\% - 86.9\%$ | Widespread College Credit | Target specific unit deficits to push into the top-tier AP 5 score boundary. |
| **AP 3** | $52.1\% - 69.4\%$ | Passing / State Credit | Reinforce core multiple-choice test strategy and common misconception traps. |
| **AP 2** | $34.6\% - 52.0\%$ | Non-Passing / At-Risk | Requires structured review sessions and core concept remediation to cross $52.1\%$. |
| **AP 1** | $< 34.6\%$ | No College Credit | Urgent 1-on-1 tutoring and foundational prerequisite recovery required. |

---

## 📐 Mathematical Model & Formulas

The predictor utilizes an ordinary least squares (OLS) linear model fitted on historical course data:

$$\text{Predicted AP Score} = 0.05731 \times (\text{In-Class Test Average \%}) - 0.4851$$

Bounded strictly within the official score range:

$$\hat{y} = \max\left(1.0, \min\left(5.0, 0.05731x - 0.4851\right)\right)$$

### Inverting for Grade Thresholds
Setting $\hat{y}$ to the midpoints of the integer score boundaries yields the empirical cutoffs:
* **Score 5 cutoff** ($\hat{y} = 4.5$): $\frac{4.5 + 0.4851}{0.05731} \approx 87.0\%$
* **Score 4 cutoff** ($\hat{y} = 3.5$): $\frac{3.5 + 0.4851}{0.05731} \approx 69.5\%$
* **Score 3 cutoff** ($\hat{y} = 2.5$): $\frac{2.5 + 0.4851}{0.05731} \approx 52.1\%$
* **Score 2 cutoff** ($\hat{y} = 1.5$): $\frac{1.5 + 0.4851}{0.05731} \approx 34.6\%$

### Normal Error Probability Formulation
Probabilities are derived from the Gaussian Cumulative Distribution Function:

$$\Phi(x; \mu, \sigma) = \frac{1}{2} \left[1 + \text{erf}\left(\frac{x - \mu}{\sigma \sqrt{2}}\right)\right]$$

where $\mu = \hat{y}$ and residual standard deviation $\sigma = 0.744$. The error function $\text{erf}(z)$ is approximated using the high-precision Abramowitz & Stegun Chebyshev polynomial formula.

---

## 🔒 Data Privacy & Zero Student PII

* **100% Client-Side Execution**: All calculations, sliders, and chart filtering execute entirely inside your local web browser. No student data, grade values, or inputs are transmitted over the network or saved externally.
* **Anonymous Dataset**: The embedded historical dataset contains zero Personally Identifiable Information (PII), student names, or student IDs.
* Fully compliant with educational privacy guidelines (FERPA).

---

## 💻 How to Use

1. Navigate to the `AP Score Predictor` folder.
2. Open [`index.html`](file:///Users/davidknuffke/Documents/Programming/EduScripts/AP%20Score%20Predictor/index.html) in any modern web browser (Google Chrome, Safari, Firefox, or Edge).
   * Or from the terminal:
     ```bash
     open "AP Score Predictor/index.html"
     ```
3. Use the **Interactive Score Calculator** tab for student conferences and goal-setting.
4. Switch to the **Historical Plot & Cutoff Boundaries** tab to inspect cohort data distributions and grade cutoffs.

---

## 🔗 Relationship to Other EduScripts Tools

This dashboard serves as a browser-friendly, interactive companion to [`APExamScoreGraphs`](file:///Users/davidknuffke/Documents/Programming/EduScripts/APExamScoreGraphs/README.md):
* **`APExamScoreGraphs/MasterAnalysis.py`**: The heavy-duty Python data science pipeline that ingests raw gradebook CSVs, detects grade mismatches, runs ANOVA and Tukey HSD post-hoc tests, and generates multi-page publication-quality PDF reports.
* **`AP Score Predictor`**: The lightweight, zero-dependency interactive dashboard intended for daily classroom access, student check-ins, and department meetings.

---

## 🔄 Annual Model Updating (Version Cycle)

This dashboard is updated year-on-year to reflect incoming AP cohort results:
* **Current Version (`v.26-27`)**: Trained on historical exam data through the 2025–2026 academic year ($N = 193$), parameterized for UCAS predicted grades and advising during the 2026–2027 school year.
* **Annual Workflow**: Following the release of summer AP Exam scores, re-run [`APExamScoreGraphs/MasterAnalysis.py`](file:///Users/davidknuffke/Documents/Programming/EduScripts/APExamScoreGraphs/MasterAnalysis.py) to recalculate the OLS slope, intercept, $R^2$, and RMSE. Update the `MODEL` constants at the top of [`index.html`](file:///Users/davidknuffke/Documents/Programming/EduScripts/AP%20Score%20Predictor/index.html#L380) to produce the next annual release (e.g., `v.27-28`).

---

## 📄 License

This project is licensed under the terms of the [MIT License](file:///Users/davidknuffke/Documents/Programming/EduScripts/AP%20Score%20Predictor/LICENSE).
