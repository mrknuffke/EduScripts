# Gradebook Reporter

**Version**: 18.Jan.2026

## Overview
Gradebook Reporter is a Google Apps Script utility that generates individual student progress reports from a spreadsheet gradebook. It can send these reports directly via email or compile them into a Google Doc for printing/archiving.

## Features
-   **Email Reports**: Sends personalized HTML emails to students with their grades and missing assignment alerts.
-   **Drive Reports**: Generates a single Google Doc with page breaks between student reports.
-   **Student Selector**: A UI dialog to filter students by section and select specific individuals.
-   **Demo Mode**: Generates a template gradebook to demonstrate functionality.
-   **Safety Checks**: identifying potential email/name mismatches before sending.

## Setup
-   Requires columns for "Name" and "Email".
-   Recognizes standard headers (Assignments, Standards, Categories).
-   Use `Gradebook Tools` > `Generate Demo Gradebook` to see the expected format.

## Usage
1.  Open your gradebook sheet.
2.  Go to `Gradebook Tools` > `Email Reports` or `Generate Reports (Drive)`.
3.  Select students from the dialog.
4.  Click "Run".
