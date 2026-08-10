# Miscellaneous Tools

A dedicated folder for standalone utility scripts, single-purpose automation helpers, micro Python scripts, and experimental Google Apps Scripts.

---

## 🛠️ Script Index

| Script File | Type | Description | Requirements / Setup |
| :--- | :--- | :--- | :--- |
| [**`DocFontTools.gs`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/Miscellaneous%20Tools/DocFontTools.gs) | Apps Script | Google Docs utility script that adds a custom **Font Tools** menu. Features **Highlight Non-Brand Fonts** (scans document text attribute runs and highlights non-brand fonts in yellow) and **Generate Font Report** (creates a new Google Sheet report in Google Drive root auditing all text snippets, font families, sizes, and styling). | Open Google Doc → Extensions → Apps Script → Paste code. |
| [**`DocStandardizer.gs`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/Miscellaneous%20Tools/DocStandardizer.gs) | Apps Script | Standalone batch tool to standardize formatting across Google Docs (batch folder or single doc). Converts straight to smart quotes, fixes Unicode subscripts/superscripts to native formatting, promotes "✎Note Space N:" paragraphs to Heading 1, cleans table row heights & in-cell paragraph spacing, clears paragraph bottom borders, and applies Garamond/Montserrat typography while preserving custom author text colors. | Requires **Google Docs API** Advanced Service:<br>Apps Script Editor → Services (`+`) → Google Docs API → Add. |
| [**`move_to_root_and_modify.py`**](file:///Users/davidknuffke/Documents/Programming/EduScripts/Miscellaneous%20Tools/move_to_root_and_modify.py) | Python Utility | Recursive directory file flattener & batch renamer. Moves all files nested inside subdirectories into a root folder, appends optional prefixes to all files, resolves filename collisions (`_1`, `_2`), and cleans up empty subdirectories. Features an **Interactive Guided Wizard** and CLI options (`--dry-run`, `--prefix`, `--remove-empty-dirs`). | Python 3.6+ (Standard Library). Run interactively with `python3 move_to_root_and_modify.py` or with CLI flags. |

---

## 🚀 Script Highlights & Usage

### 📂 `move_to_root_and_modify.py` (Directory Flattener & Renamer)

#### 1. Interactive Guided Wizard (Default)
Run the script without flags to launch an interactive terminal prompt:
```bash
python3 move_to_root_and_modify.py
```
The wizard will guide you through:
1. **Target Directory**: Folder to flatten (defaults to current directory).
2. **File Prefix**: Optional prefix string to append to all files (e.g. `2026_` or `Unit1_`).
3. **Empty Folder Cleanup**: Option to auto-delete subdirectories after emptying them.
4. **Dry-Run Preview**: Preview all file moves and renames before modifying anything.

#### 2. Command Line Arguments (Power Users)
```bash
# Preview operations without changing files (Dry-Run)
python3 move_to_root_and_modify.py /path/to/folder --dry-run

# Flatten folder and prepend prefix to all files
python3 move_to_root_and_modify.py /path/to/folder --prefix "2026_"

# Flatten, prefix, remove empty directories, and skip confirmation prompts
python3 move_to_root_and_modify.py /path/to/folder -p "Unit1_" -r -y
```

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
