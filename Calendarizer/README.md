# Calendarizer

**Version**: 18.Jan.2026

## Overview
Calendarizer is a Google Apps Script tool designed to visualize instructional pacing charts as beautiful, printable calendars. It supports both "Wall Calendar" (traditional month view) and "Lateral Calendar" (linear horizontal view) layouts.

## Features
-   **Custom Views**: generate `Wall Calendar` or `Lateral Calendar` sheets from your pacing data.
-   **Smart Parsing**: Automatically identifies class blocks (A/B/C/D), holidays, and notes from a structured list.
-   **Configuration**: Customize start month, holiday keywords, and color palettes via a settings menu.
-   **Styling**: Features a modern pastel color palette and clean typography.
-   **Tutorial**: Integrated sidebar tutorial to help new users get started.

## Configuration
-   **Start Month**: Set the academic year start (e.g., July, August, September).
-   **Keywords**: Define words that trigger "Holiday" styling (e.g., "Break", "No School").
-   **Colors**: Customize the 12-month color cycle.

## Usage
1.  Format your source sheet with columns: Class Number, Start Date, End Date, Block/Type, Notes.
2.  Name the sheet `YYYY-YYYY` (e.g., `2025-2026`).
3.  Use the `Calendar Tools` menu to generate your desired view.
