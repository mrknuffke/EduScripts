# Gradebook Reporter

**Version**: 25.Aug.2026

## Overview
Gradebook Reporter is a Google Apps Script utility that generates individual student progress reports from a spreadsheet gradebook. It can send these reports directly via email or compile them into a Google Doc for printing/archiving.

## Features
-   **Flexible Reporting**: Automatically groups complex assessments (Quizzes, Labs, WebAssigns) based on structural clues (header keywords or standard rows), regardless of the Category name.
-   **Participation Metrics**: Identifies and reports summary statistics (e.g., "% Incomplete", "Completion Rate") in a dedicated, color-coded table (Green = Good, Red/Orange = Needs Work).
-   **Smart Layouts**:
    -   **Chemistry Mode**: Separates Formative work from Summative Standard Mastery.
    -   **AP Bio Mode**: Automatically detects "Topic Quest Labs" and other complex headers.
-   **Missing Work Logic**: Intelligently suppresses "Congratulations" messages if summary stats indicate missing assignments, even if individual items aren't flagged.
-   **Robust Handling**: Works on gradebooks with or without a "Standards" row, automatically falling back to simpler grouping methods.
-   **Layout-Agnostic Roster Parsing**: Nothing about the roster is hard-coded to a fixed row or column. The script finds the roster header row by its labels (it does not have to be Row 2), reads student data from the row below it, and identifies checkbox columns so Sheets' `TRUE`/`FALSE` display values are never mistaken for section names or graded work. A row is treated as a student only if it has a name plus real evidence of a person (an email, a parent email, graded work, or a `Last, First` name), so course titles and banner rows become section headings instead of phantom students.
-   **Email Reports**: Sends personalized HTML emails to students with their grades and missing assignment alerts.
-   **Drive Reports**: Generates a single Google Doc with page breaks between student reports.
-   **Student Selector**: A UI dialog that groups students into one collapsible card per class section, with a section filter, per-section select-all, and a live count of what's selected.
-   **Preview Mode**: Preview up to 10 reports with clear separation to verify layout before sending.
-   **Fun Feedback**: Includes a library of silly, encouraging emoji puns for students with no missing work.

## Installation
1.  Open the Google Sheet where you wish to use the reporter.
2.  Navigate to `Extensions` > `Apps Script`.
3.  If there is any code in the default `Code.gs` file, delete it.
4.  Copy the entire content of `ReportScript.gs` from this repository.
5.  Paste the code into the Apps Script editor.
6.  Save the project (Click the disk icon or press `Cmd/Ctrl + S`).
7.  Reload your Google Sheet.
8.  **Authorization**: The first time you run a function from the new menu, Google will ask for permission.
    -   Click `Continue`.
    -   Select your Google Account.
    -   Click `Advanced` (if a "Google hasn't verified this app" warning appears).
    -   Click `Go to (Script Name) (unsafe)`.
    -   Click `Allow`.

## Setup
-   Requires columns for "Name" and "Email" (found by header label anywhere in rows 1-3; Column B is the default for "Name").
-   Recognizes standard headers.
-   **Class Sections** (optional, two interchangeable options):
    -   Put the section name (e.g. `Block 1`) in Column A of every student row, **or**
    -   Give each class a heading row containing only the section name - no email, no grades. Row styling is optional.
-   **Keywords**: Supports "Quiz", "Test", "Quest", "Lab", "WebAssign", "unit", "assess" for automatic grouping.
-   Use `Gradebook Tools` > `Generate Demo Gradebook` to see the expected format.

## Testing
Apps Script cannot be run locally, so the roster-parsing logic is covered by fixture sheets that mirror real gradebook layouts (checkbox columns, merged banner dividers, roster headers below the assignment header, per-row section columns, unstyled headings).

```bash
./test/run-tests.sh
```

The runner extracts the roster functions straight out of `ReportScript.gs` and executes them, so the tests always run against the shipped code. It uses `node` when available and otherwise falls back to macOS's built-in JavaScript engine, so it needs no installation. **When a gradebook is parsed wrongly, add its shape to `test/fixtures.js` before changing the parser** - that is what keeps one sheet's fix from breaking another's.

## Troubleshooting
Run `Gradebook Tools` > `🛠️ Setup Checker & Guide` on the misbehaving sheet. The first result card reports which row the roster headers were read from and which row student data starts at. If those row numbers are wrong, every other result will be wrong too, and that is the thing to fix first.

## Usage
1.  Open your gradebook sheet.
2.  Go to `Gradebook Tools` > `Email Reports` (or `Preview` / `Drive`).
3.  Select students from the dialog.
4.  Click "Run".
