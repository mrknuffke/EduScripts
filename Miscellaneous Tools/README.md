# Miscellaneous Tools

A dedicated folder for standalone utility scripts, single-purpose automation helpers, micro Python scripts, and experimental Google Apps Scripts.

---

## 🛠️ Script Index

| Script File | Type | Description | Requirements / Setup |
| :--- | :--- | :--- | :--- |
| [**`DocFontTools.gs`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/Miscellaneous%20Tools/DocFontTools.gs) | Apps Script | Google Docs utility script that adds a custom **Font Tools** menu. Features **Highlight Non-Brand Fonts** (scans document text attribute runs and highlights non-brand fonts in yellow) and **Generate Font Report** (creates a new Google Sheet report in Google Drive root auditing all text snippets, font families, sizes, and styling). | Open Google Doc → Extensions → Apps Script → Paste code. |
| [**`DocStandardizer.gs`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/Miscellaneous%20Tools/DocStandardizer.gs) | Apps Script | Standalone batch tool to standardize formatting across Google Docs (batch folder or single doc). Converts straight to smart quotes, fixes Unicode subscripts/superscripts to native formatting, promotes "✎Note Space N:" paragraphs to Heading 1, cleans table row heights & in-cell paragraph spacing, clears paragraph bottom borders, and applies Garamond/Montserrat typography while preserving custom author text colors. | Requires **Google Docs API** Advanced Service:<br>Apps Script Editor → Services (`+`) → Google Docs API → Add. |

---

## 📝 Best Practices for Adding Scripts

When adding new scripts into this directory, please adhere to the following best practices:

### 1. File Naming Conventions
- **Google Apps Scripts**: Use CamelCase or descriptive names with the `.gs` extension (e.g. `DocFontTools.gs`, `DocStandardizer.gs`).
- **Python Utilities**: Use lowercase snake_case with the `.py` extension (e.g. `csv_cleaner.py`, `seating_chart_exporter.py`).
- **Shell Scripts**: Use lowercase with the `.sh` extension (e.g. `backup_reports.sh`).

### 2. In-Script Documentation
Every script added to this directory should begin with a header comment/docstring containing:
- **Title & Description**: Brief explanation of what the script does.
- **Author & Date**: Maintainer info.
- **Dependencies**: List of required Python packages or Google APIs.
- **Usage Example**: Command line snippet or Apps Script function entry point.

### 3. Data Privacy & Security Rules
- **No Hardcoded Secrets**: Never commit passwords, Google API credentials, webhooks, or private tokens directly in script files. Use environment variables or Google Apps Script `PropertiesService.getUserProperties()`.
- **No Student PII**: Never save sample dataset files containing actual student names, IDs, or grades. Place sample data in ignored `.csv` files or use synthetic anonymized placeholders.
