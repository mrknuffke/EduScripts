# Calendarizer

**Version**: 19.Jan.2026

## Overview
Calendarizer is a Google Apps Script tool designed to visualize instructional pacing charts as beautiful, printable calendars. It supports both "Wall Calendar" (traditional month view) and "Lateral Calendar" (linear horizontal view) layouts.

## Features
-   **Custom Views**: generate `Wall Calendar` or `Lateral Calendar` sheets from your pacing data.
-   **Smart Parsing**: Automatically identifies class blocks (A/B/C/D), holidays, and notes from a structured list.
-   **Configuration**: Customize start month, holiday keywords, and color palettes via a settings menu.
-   **Styling**: Features a modern pastel color palette and clean typography.
-   **Tutorial**: Integrated sidebar tutorial to help new users get started.

## Installation
1.  Open the Google Sheet where you wish to use the calendarizer.
2.  Navigate to `Extensions` > `Apps Script`.
3.  If there is any code in the default `Code.gs` file, delete it.
4.  Copy the entire content of `calendarizer.gs` from this repository.
5.  Paste the code into the Apps Script editor.
6.  Save the project (Click the disk icon or press `Cmd/Ctrl + S`).
7.  Reload your Google Sheet.
8.  **Authorization**: The first time you run a function from the new menu, Google will ask for permission.
    -   Click `Continue`.
    -   Select your Google Account.
    -   Click `Advanced` (if a "Google hasn't verified this app" warning appears).
    -   Click `Go to (Script Name) (unsafe)`.
    -   Click `Allow`.

## Configuration
-   **Start Month**: Set the academic year start (e.g., July, August, September).
-   **Keywords**: Define words that trigger "Holiday" styling (e.g., "Break", "No School").
-   **Colors**: Customize the 12-month color cycle.

## Usage
1.  Format your source sheet with columns: Class Number, Start Date, End Date, Block/Type, Notes.
2.  Name the sheet `YYYY-YYYY` (e.g., `2025-2026`).
3.  Use the `Calendar Tools` menu to generate your desired view.
