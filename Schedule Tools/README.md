# Schedule Tools

**Version**: 7.May.2026

## Overview
Schedule Tools is a Google Apps Script toolset for generating, formatting, and managing department master schedules in Google Sheets. It parses raw exported CSV data into clean, color-coded semester grids with A/B/C/D day block layouts.

## Features
-   **Department Setup Wizard**: Creates a native config sheet pre-filled with courses scanned from your data. Define hex colors, abbreviations, and formatting preferences directly in the spreadsheet.
-   **Master Schedule Builder**: Parses raw CSV data and generates two formatted schedule sheets (Semester 1 & 2) with teachers, days, and blocks automatically mapped.
-   **Course Entry Editor**: Selectively strip or keep lines of cell content (section info, room numbers, block/day, term) and apply abbreviations from your config. Supports batch processing across all schedule sheets.
-   **Color and Format Courses**: Applies background colors, font colors, font families, bold titles, alignment, and wrap settings from your config. Can format a single sheet or all schedule sheets at once.
-   **Day Swapper (C/D Generator)**: Generates C and D day columns by swapping block order from A and B days, preserving all formatting.
-   **Refresh Courses**: Scans schedule sheets for any new course names not yet in your config and appends them with default styling.
-   **Config Validation**: Checks your config sheet for invalid hex codes, missing `#` prefixes, and course names that don't match any schedule data.

## Installation
1.  Open the Google Sheet where you wish to use the schedule tools.
2.  Navigate to `Extensions` > `Apps Script`.
3.  If there is any code in the default `Code.gs` file, delete it.
4.  Copy the entire content of `ScheduleTools.gs` from this repository.
5.  Paste the code into the Apps Script editor.
6.  Save the project (Click the disk icon or press `Cmd/Ctrl + S`).
7.  Reload your Google Sheet.
8.  **Authorization**: The first time you run a function from the new menu, Google will ask for permission.
    -   Click `Continue`.
    -   Select your Google Account.
    -   Click `Advanced` (if a "Google hasn't verified this app" warning appears).
    -   Click `Go to (Script Name) (unsafe)`.
    -   Click `Allow`.

## Usage
1.  Use the `Department Tools` menu to access all tools.
2.  Start with `How to Use / Tutorial` for a guided walkthrough.
3.  Run `Department Setup Wizard` to create your config sheet.
4.  Paste your raw schedule data, then run `Create Initial Schedule Grid`.
5.  Use `Course Entry Editor` and `Color and Format Courses` to polish the output.
6.  Run `Generate C/D Days` to populate C and D day columns from A and B.
