# Gradebook Reporter

**Version**: 27.Jan.2026

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
-   **Email Reports**: Sends personalized HTML emails to students with their grades and missing assignment alerts.
-   **Drive Reports**: Generates a single Google Doc with page breaks between student reports.
-   **Student Selector**: A UI dialog to filter students by section and select specific individuals.
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
-   Requires columns for "Name" and "Email".
-   Recognizes standard headers.
-   **Keywords**: Supports "Quiz", "Test", "Quest", "Lab", "WebAssign", "unit", "assess" for automatic grouping.
-   Use `Gradebook Tools` > `Generate Demo Gradebook` to see the expected format.

## Usage
1.  Open your gradebook sheet.
2.  Go to `Gradebook Tools` > `Email Reports` (or `Preview` / `Drive`).
3.  Select students from the dialog.
4.  Click "Run".
