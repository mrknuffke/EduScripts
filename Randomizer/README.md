# Randomizer

**Version**: 19.Jan.2026

## Overview
Randomizer is a comprehensive classroom management tool for Google Sheets. It facilitates creating random groups, managing seating charts, and balancing student distributions based on custom attributes.

## Features
-   **Roster Setup Wizard**: Easily configure sections, rooms, and attributes.
-   **Randomization**: Assign students to tables/groups with sophisticated constraints.
    -   **Social Mixer**: Maximizes variety in partners over time.
    -   **Balancing**: Distributes students explicitly by attributes (e.g., Gender, Skill Level).
    -   **Constraints**: Supports "Student Buddies" (must sit together) and "Separations" (must not sit together).
    -   **Preferential Seating**: Force specific students to specific tables.
-   **Layout Manager**: Save and load different seating arrangements.
-   **History Tracking**: Remembers past groups to avoid repeats.
-   **Onboarding**: Includes a tutorial sidebar and demo class generator.

## Installation
1.  Open the Google Sheet where you wish to use the randomizer.
2.  Navigate to `Extensions` > `Apps Script`.
3.  If there is any code in the default `Code.gs` file, delete it.
4.  Copy the entire content of `Randomizer.gs` from this repository.
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
1.  Use the `Randomizer` menu to access all tools.
2.  Start with `Roster Setup Wizard` if setting up a new sheet.
3.  Configure your Rooms and Tables.
4.  Run `Randomly Assign Students` to generate groups.
