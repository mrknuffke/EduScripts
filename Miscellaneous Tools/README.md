# Miscellaneous Tools

A dedicated folder for standalone utility scripts, single-purpose automation helpers, micro Python scripts, and experimental Google Apps Scripts.

---

## 🛠️ Overview & Purpose

As educational workflows evolve, smaller utility scripts often serve targeted needs that don't warrant an entire multi-file project directory. The `Miscellaneous Tools` directory acts as a clean, structured repository home for these scripts.

---

## 📝 Best Practices for Adding Scripts

When adding new scripts into this directory, please adhere to the following best practices:

### 1. File Naming Conventions
- **Google Apps Scripts**: Use CamelCase or descriptive names with the `.gs` extension (e.g. `AttendanceAlert.gs`, `RosterParser.gs`).
- **Python Utilities**: Use lowercase snake_case with the `.py` extension (e.g. `csv_cleaner.py`, `seating_chart_exporter.py`).
- **Shell Scripts**: Use lowercase with the `.sh` extension (e.g. `backup_reports.sh`).

### 2. In-Script Documentation
Every script added to this directory should begin with a header comment/docstring containing:
- **Title & Description**: Brief explanation of what the script does.
- **Author & Date**: Maintainer info.
- **Dependencies**: List of required Python packages or Google APIs.
- **Usage Example**: Command line snippet or Apps Script function entry point.

#### Apps Script Header Example:
```javascript
/**
 * Title: Form Response Auto-Archive
 * Description: Moves older Google Form responses to an archive tab automatically.
 * Author: David Knuffke
 * Target: Google Sheets / Apps Script
 */
```

#### Python Script Header Example:
```python
"""
Title: Student CSV Anonymizer
Description: Strips names and email addresses from a spreadsheet CSV before public sharing.
Requirements: pandas
Usage: python3 csv_anonymizer.py input.csv output.csv
"""
```

### 3. Data Privacy & Security Rules
- **No Hardcoded Secrets**: Never commit passwords, Google API credentials, webhooks, or private tokens directly in script files. Use environment variables or Google Apps Script `PropertiesService.getUserProperties()`.
- **No Student PII**: Never save sample dataset files containing actual student names, IDs, or grades. Place sample data in ignored `.csv` files or use synthetic anonymized placeholders.
