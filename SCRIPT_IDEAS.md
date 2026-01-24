# Potential Google Apps Script Projects

Here are 10 ideas for new Google Apps Scripts tools to support teachers. Feel free to add comments, questions, or modify these ideas directly in this document.

## 1. SlideDecker (Docs to Slides Converter)
*   **Problem:** Teachers often draft lesson plans or outlines in Google Docs but find the process of manually copy-pasting content into Slides for presentation tedious.
*   **Solution:** A script that parses a Google Doc based on header hierarchy (e.g., `H1` = New Slide Title, `Normal Text` = Bullet points) and automatically generates a formatted Google Slide presentation.
*   **Tech Stack:** Docs API, Slides API.

## 2. FormRoster Sync (Sheets to Forms)
*   **Problem:** Teachers maintain master rosters in Sheets but frequently create new Google Forms for quizzes or rubrics. Manually ensuring the "Student Name" dropdown matches the current roster is prone to error and repetitive.
*   **Solution:** A script links specific Google Forms to a Master Roster Sheet. Updates to the roster automatically propagate to the "Student Name" dropdowns in all linked Forms.
*   **Tech Stack:** Forms API, Sheets Triggers.

## 3. AssetTracker (Classroom Inventory System)
*   **Problem:** Managing classroom resources like calculators, books, or lab equipment is chaotic. Knowing who has what item is difficult.
*   **Solution:** A system using a Google Form and a barcode scanner (via phone camera) to check items in and out. The script processes responses to update a "Current Inventory" sheet, tracking status ("In"/"Out") and current holder. Can include automated email reminders for overdue items.
*   **Tech Stack:** Forms, Sheets, GmailApp.

## 4. QuizBuilder (Sheets to Forms)
*   **Problem:** The Google Forms UI is slow for creating lengthy quizzes or tests with many distractors.
*   **Solution:** A rapid-entry spreadsheet template where teachers list questions, options, and correct answers. The script reads this data and generates a fully configured Google Form Quiz with answer keys pre-set.
*   **Tech Stack:** SpreadsheetApp, FormsApp.

## 5. ParentLog (Gmail to Sheets)
*   **Problem:** Documenting parent communication is required but copying emails from Gmail to a log is friction-heavy.
*   **Solution:** A Gmail sidebar add-on. When viewing an email, a "Log to Sheet" button parses the sender, subject, date, and body snippet, appending it to a "Communication Log" spreadsheet for the relevant student.
*   **Tech Stack:** Gmail Add-on, SpreadsheetApp.

## 6. DriveTidy (Student Portfolio Automator)
*   **Problem:** Student digital work submitted via Drive becomes disorganized. Creating individual portfolios and distributing materials is time-consuming.
*   **Solution:** A script that generates a structured folder hierarchy for every student on a roster (e.g., `Class 2026 > [Student Name]`). It includes a "distribution" feature: dropping a file into a "Handout" folder copies it to every student's individual folder.
*   **Tech Stack:** DriveApp.

## 7. RubricMash (Form Responses to Doc)
*   **Problem:** Grading via Google Forms produces a row of data, which is not a user-friendly format for student feedback.
*   **Solution:** On Form submission, the script takes the grading data, populates a Google Doc template (replacing placeholders like `{{StudentName}}` and `{{Score}}`), converts the Doc to a PDF, and emails it to the student.
*   **Tech Stack:** DocsApp, DriveApp (for PDF conversion), GmailApp.

## 8. EventPlanner (Sheet to Calendar Events)
*   **Problem:** School schedules (exams, PD days, rotations) often arrive as tables or documents, requiring manual entry into Google Calendar.
*   **Solution:** A script that ingests a spreadsheet containing Date, Title, Description, and Time, then bulk-creates actual Google Calendar events. It can also handle guest invitations.
*   **Tech Stack:** CalendarApp.

## 9. AbsenteeAlerter (Attendance Dashboard)
*   **Problem:** Identifying attendance patterns (e.g., chronic absenteeism) requires manual tracking and calculation.
*   **Solution:** A daily attendance sheet where checking a box marks a student absent. The script calculates running totals and, upon triggering specific thresholds (3, 5, 10 absences), drafts a "Notice of Attendance" email to parents/guidance for teacher review.
*   **Tech Stack:** SpreadsheetApp, GmailApp.

## 10. GroupWork Guardian (Docs Activity Monitor)
*   **Problem:** Grading group work is difficult when contribution levels are unclear ("I did all the work!").
*   **Solution:** A script that analyzes the revision history of a Google Doc (or folder of Docs) to report approximate contribution percentages or "last edited by" stats for each user involved.
*   **Tech Stack:** DriveApp (Advanced Drive Service).
