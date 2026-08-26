/**
 * MIT License
 * Copyright (c) 2026 David Knuffke
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

/**
 * GRADEBOOK REPORT GENERATOR
 * Instructions:
 * 1. Click "Gradebook Tools" > "Email Reports" or "Generate Reports (Drive)".
 * 2. Select students.
 * 3. Process.
 */

/**
 * Expands shorthand prefixes in assignment names for display.
 * @param {string} name - The assignment name to expand
 * @return {string} The expanded name
 */
function expandAssignmentPrefix(name) {
  if (!name) return name;
  // Case-insensitive prefix expansion
  return name
    .replace(/^AC:/i, 'Activity:')
    .replace(/^ID:/i, 'InfoDoc:');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gradebook Tools')
    .addItem('ℹ️ Help & Tutorial', 'showTutorialSidebar')
    .addItem('🛠️ Setup Checker & Guide', 'showSetupGuide')
    .addSeparator()
    .addItem('📧 Email Reports (Selector)', 'openEmailSelector')
    .addItem('📂 Generate Reports (Drive)', 'openDriveSelector')
    .addSeparator()
    .addItem('⚙️ Set Reply-To Email', 'setReplyToEmail')
    .addItem('📘 Generate Demo Gradebook', 'generateGradebookTemplate')
    .addToUi();
}

/**
 * Gets the stored reply-to email for this spreadsheet.
 * @return {string} The reply-to email or empty string if not set
 */
function getReplyToEmail() {
  const props = PropertiesService.getDocumentProperties();
  return props.getProperty('REPLY_TO_EMAIL') || '';
}

/**
 * Prompts the user to set/update the reply-to email address.
 */
function setReplyToEmail() {
  const ui = SpreadsheetApp.getUi();
  const currentEmail = getReplyToEmail();

  const promptMessage = currentEmail
    ? `Current reply-to email: ${currentEmail}\n\nEnter a new email address (or leave blank to clear):`
    : 'Enter the email address that students should reply to:';

  const response = ui.prompt('Set Reply-To Email', promptMessage, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    const newEmail = response.getResponseText().trim();
    const props = PropertiesService.getDocumentProperties();

    if (newEmail === '') {
      props.deleteProperty('REPLY_TO_EMAIL');
      ui.alert('Reply-to email has been cleared. Emails will be sent without a reply-to address.');
    } else if (newEmail.includes('@')) {
      props.setProperty('REPLY_TO_EMAIL', newEmail);
      ui.alert(`Reply-to email set to: ${newEmail}`);
    } else {
      ui.alert('Invalid email address. Please include an @ symbol.');
    }
  }
}

function openEmailSelector() { showStudentSelector('email'); }
function openDriveSelector() { showStudentSelector('drive'); }

/**
 * Generates a template sheet for testing/onboarding.
 */
function generateGradebookTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Demo Gradebook");
  if (sheet) {
    SpreadsheetApp.getUi().alert("A sheet named 'Demo Gradebook' already exists.");
    return;
  }

  sheet = ss.insertSheet("Demo Gradebook");

  // Set Headers using colors from the script logic
  const headers = ["Section", "Name", "Email", "Parent Email", "Assignment 1", "Assignment 2", "Assignment 3", "Summative Exam"];
  const categories = ["", "", "", "", "Classwork", "Classwork", "Homework", "Assessments"];
  const standards = ["", "", "", "", "Standard 1", "Standard 1", "Standard 2", "Standard 3"];

  sheet.getRange(1, 1, 1, headers.length).setValues([categories]).setFontWeight("bold").setBackground("#e0e0e0");
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#434343").setFontColor("white");
  sheet.getRange(3, 1, 1, headers.length).setValues([standards]).setFontStyle("italic").setBackground("#f3f3f3");

  // Dummy Data
  const data = [
    ["Block 1", "Potter, Harry", "harry@hogwarts.edu", "james.potter@hogwarts.edu", "1", "1", "0", "95"],
    ["Block 1", "Granger, Hermione", "hermione@hogwarts.edu", "mr.granger@londondentist.com", "1", "1", "1", "100"],
    ["Block 1", "Weasley, Ron", "ron@hogwarts.edu", "molly.weasley@hogwarts.edu", "0", "1", "Missing", "85"],
    ["Block 2", "Malfoy, Draco", "draco@hogwarts.edu", "lucius.malfoy@hogwarts.edu", "1", "Exempt", "1", "90"],
    ["Block 2", "Lovegood, Luna", "luna@hogwarts.edu", "xenophilius.lovegood@quibbler.org", "1", "1", "1", "92"]
  ];

  sheet.getRange(4, 1, data.length, data[0].length).setValues(data);

  // Format
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(3);

  SpreadsheetApp.getUi().alert("Demo Gradebook created! You can now test the reporting tools.");
}

// --- ROSTER SCANNING (shared by the Student Selector and the Setup Checker) ---

/** Words that mark a row as a class/section divider rather than a person. */
const SECTION_LABEL_WORDS = /\b(block|period|section|class|hour|homeroom|semester|quarter|term|group)\b/i;

/** Words that mark a row as a summary/statistics row rather than a person. */
const SUMMARY_ROW_WORDS = /\b(average|avg|sum|total|median|mean|count|stdev|max|min|stats?)\b/i;

/**
 * Reads a trimmed cell value from a row, tolerating a missing column index.
 */
function rosterCell(row, colIndex) {
  if (colIndex < 0 || colIndex >= row.length) return "";
  return row[colIndex] ? String(row[colIndex]).trim() : "";
}

/**
 * Returns the first cell in a row that has any content.
 */
function firstNonEmptyCell(row) {
  for (let c = 0; c < row.length; c++) {
    if (row[c] && String(row[c]).trim() !== "") return String(row[c]).trim();
  }
  return "";
}

/**
 * Flags rows where the email address doesn't look like it belongs to the name.
 */
function hasEmailNameMismatch(name, email) {
  if (!email) return false;
  if (email.indexOf('@') === -1) return true;
  const lastName = name.includes(',') ? name.split(',')[0].trim() : name.split(' ')[0].trim();
  const cleanName = lastName.replace(/[^a-zA-Z]/g, '').toLowerCase();
  const cleanEmail = email.split('@')[0].replace(/[^a-zA-Z]/g, '').toLowerCase();
  return cleanName.length > 1 && !cleanEmail.includes(cleanName);
}

/**
 * Finds the row carrying the roster labels (Name / Email / Parent Email).
 *
 * This is NOT always the assignment header row. A gradebook may list assignment
 * names in Row 2 and put the roster labels further down, above the first class
 * divider. Scoring each candidate row on the labels it contains locates it
 * wherever it sits, instead of assuming a fixed position.
 */
function findRosterHeaderRow(data) {
  const limit = Math.min(12, data.length);
  let bestRow = 0;
  let bestScore = 0;

  for (let r = 0; r < limit; r++) {
    let hasName = false, hasEmail = false, hasParent = false, hasPreferred = false;

    for (let c = 0; c < data[r].length; c++) {
      if (!data[r][c]) continue;
      const text = String(data[r][c]).trim().toLowerCase().replace(/:$/, '');
      if (text.includes('preferred')) hasPreferred = true;
      else if (text.includes('parent') || text.includes('guardian')) hasParent = true;
      else if (text.includes('email')) hasEmail = true;
      else if (text === 'name' || text === 'student name' || text === 'student') hasName = true;
    }

    const score = (hasName ? 3 : 0) + (hasEmail ? 3 : 0) + (hasParent ? 1 : 0) + (hasPreferred ? 1 : 0);
    if (score > bestScore) {                 // ties keep the earliest row
      bestScore = score;
      bestRow = r;
    }
  }

  // Nothing recognisable: fall back to the documented Row 2.
  return bestScore > 0 ? bestRow : Math.min(1, data.length - 1);
}

/**
 * Finds columns whose data rows hold nothing but checkbox values. Sheets renders
 * a checkbox as the display string "TRUE"/"FALSE", which must never be mistaken
 * for a section name or for evidence of graded work.
 */
function findCheckboxColumns(data, startRow) {
  let width = 0;
  for (let r = 0; r < data.length; r++) width = Math.max(width, data[r].length);

  const checkboxCols = {};
  for (let c = 0; c < width; c++) {
    let booleans = 0;
    let others = 0;
    for (let r = startRow; r < data.length; r++) {
      const value = rosterCell(data[r], c);
      if (value === "") continue;
      if (isBooleanText(value)) booleans++; else others++;
    }
    // Majority rather than all-or-nothing: a checkbox column often also carries
    // a merged divider label or a headcount on a banner row.
    if (booleans >= 2 && booleans / (booleans + others) >= 0.75) checkboxCols[c] = true;
  }
  return checkboxCols;
}

/**
 * True when a cell is filled like a divider banner: a dark solid fill, or any
 * non-white fill carrying white text. A banner row is never a student.
 */
function isBannerCell(background, fontColor) {
  const bg = (background || "").toLowerCase();
  const font = (fontColor || "").toLowerCase();
  if (bg === '#000000' || bg === '#434343' || bg === '#666666') return true;
  return bg !== "" && bg !== '#ffffff' && (font === '#ffffff' || font === '#fff');
}

/**
 * Reads fill and font colours for just the columns that can carry a divider
 * banner, rather than pulling formatting for the whole sheet.
 */
function readBannerStyles(sheet, rowCount, colIndexes) {
  const styles = {};
  colIndexes.forEach(function (c) {
    if (c < 0 || styles[c]) return;
    const range = sheet.getRange(1, c + 1, rowCount, 1);
    styles[c] = { bg: range.getBackgrounds(), font: range.getFontColors() };
  });
  return styles;
}

/**
 * Locates the Name / Email / Parent Email / Section columns by scanning the
 * three header rows. Falls back to the documented positions when a label is
 * missing; nameFallback records whether that fallback was needed.
 */
function findRosterColumns(data, headerRow, checkboxCols) {
  const cols = { name: -1, preferred: -1, email: -1, parentEmail: -1, section: -1, nameFallback: false };

  // Read the roster header row first, then sweep the rows above it for any
  // label it did not carry (some gradebooks split them across header rows).
  const rowsToScan = [headerRow];
  for (let r = 0; r < headerRow; r++) rowsToScan.push(r);

  for (let i = 0; i < rowsToScan.length; i++) {
    const r = rowsToScan[i];
    for (let c = 0; c < data[r].length; c++) {
      if (!data[r][c]) continue;
      const text = String(data[r][c]).trim().toLowerCase();

      if (text.includes('parent') || text.includes('guardian')) {
        if (cols.parentEmail === -1) cols.parentEmail = c;
      } else if (text.includes('email')) {
        if (cols.email === -1) cols.email = c;
      } else if (cols.section === -1 && c <= 2 && /^(section|block|period|class|group)\b/.test(text)) {
        // Only the leftmost columns can be the Section column; otherwise an
        // assignment header like "Class Discussion" would be mistaken for one.
        cols.section = c;
      } else if (cols.preferred === -1 && text.includes('preferred')) {
        cols.preferred = c;
      } else if (cols.name === -1 && text.includes('name')) {
        cols.name = c;
      }
    }
  }

  if (cols.name === -1) {
    cols.name = 1;                 // Column B is the documented default
    cols.nameFallback = true;
  }

  // A checkbox column is not a section column, however it is labelled.
  if (cols.section > -1 && checkboxCols[cols.section]) cols.section = -1;

  // Fall back to Column A only if it is unclaimed and actually holds section
  // text. Gradebooks often use Column A for selection checkboxes instead.
  if (cols.section === -1 &&
      cols.name !== 0 && cols.preferred !== 0 && cols.email !== 0 && cols.parentEmail !== 0 &&
      !checkboxCols[0] && columnHasLabelText(data, 0, headerRow + 1)) {
    cols.section = 0;
  }

  // Everything up to and including this column is roster data, never an
  // assignment. Report generation uses it to find where assignments begin.
  cols.lastRosterCol = Math.max(
    cols.name, cols.preferred, cols.email, cols.parentEmail, cols.section);

  return cols;
}

/**
 * Resolves the roster columns from sheet values already in hand, without
 * re-reading the sheet. Used by report generation, which loads the data itself.
 */
function resolveRosterColumns(data) {
  const headerRow = findRosterHeaderRow(data);
  const checkboxCols = findCheckboxColumns(data, headerRow + 1);
  return findRosterColumns(data, headerRow, checkboxCols);
}

/**
 * True when a column holds at least one non-numeric, non-boolean label below
 * the header rows - the mark of a real Section column.
 */
function columnHasLabelText(data, colIndex, startRow) {
  for (let r = startRow; r < data.length; r++) {
    const value = rosterCell(data[r], colIndex);
    if (value === "") continue;
    const upper = value.toUpperCase();
    if (upper === "TRUE" || upper === "FALSE") continue;
    if (!isNaN(Number(value))) continue;
    return true;
  }
  return false;
}

/**
 * Decides whether a column holds graded assessment work.
 *
 * The standards row and the category label are matched against the full
 * keyword list. The column header is matched too, but only against keywords
 * long enough to be unambiguous: the two-letter "wa" would otherwise fire on
 * ordinary headers such as "Water Stations".
 *
 * Without the header check, a scored column whose standards row and category
 * say nothing recognisable - "Formative 1.1", say - is never reported, because
 * a score is not a missing-work issue and nothing else classifies it.
 */
function matchesAssessmentColumn(header, standard, category) {
  const ASSESSMENT_KEYWORDS = ['quiz', 'test', 'exam', 'assess', 'wa', 'webassign',
                               'unit', 'quest', 'lab', 'formative', 'summative'];

  const lowerHeader = String(header || "").toLowerCase().trim();
  const lowerStandard = String(standard || "").toLowerCase().trim();
  const lowerCategory = String(category || "").toLowerCase().trim();

  const inStandardOrCategory = ASSESSMENT_KEYWORDS.some(function (k) {
    return lowerStandard.includes(k) || lowerCategory.includes(k);
  });
  if (inStandardOrCategory) return true;

  return ASSESSMENT_KEYWORDS.some(function (k) {
    return k.length >= 4 && lowerHeader.includes(k);
  });
}

// --- 0-4 RUBRIC SCORING -----------------------------------------------------

/** Descriptors for the 0-4 rubric used by Chem formatives/summatives and AP Bio labs. */
const RUBRIC_LABELS = {
  '0': 'Not Yet Evident',
  '1': 'Emerging',
  '2': 'Developing',
  '3': 'Meeting',
  '4': 'Meeting with Distinction'
};

/**
 * Converts a 0-4 rubric score into its descriptor.
 *
 * Returns null for anything that is not exactly 0, 1, 2, 3 or 4, so half marks,
 * percentages, checkbox text and blanks fall through to the normal handling.
 * A blank is deliberately not a zero: "no score yet" is not "Not Yet Evident".
 */
function rubricLabelFor(value) {
  const raw = String(value === null || value === undefined ? "" : value).trim();
  if (raw === "") return null;

  const num = Number(raw);
  if (!isFinite(num)) return null;

  const key = String(num);
  return Object.prototype.hasOwnProperty.call(RUBRIC_LABELS, key) ? RUBRIC_LABELS[key] : null;
}

/**
 * An AP Biology lab below this rubric score counts as outstanding work.
 *
 * 3 (Meeting) is a legitimate stop on the way to 4, not a problem: the lab
 * standard note carries the "get to 4 by the end of the semester" message, so
 * only work below Meeting is flagged as outstanding.
 */
const AP_BIO_LAB_CONCERN_BELOW = 3;

/**
 * Decides whether a report shows outstanding work, which suppresses the
 * congratulations message and drives the encouragement tiers.
 *
 * AP Biology labs are judged on their rubric score rather than on a Missing or
 * Incomplete marker, and are checked before the assessment exclusion below.
 * They are classified as assessments, so folding them in with everything else
 * would silently exclude them and leave the rule unable to fire at all.
 */
function hasOutstandingWork(reportRows, subjectName) {
  return reportRows.some(function (item) {
    if (item.isSummaryStat) return false;

    if (subjectName === "AP Biology" && item.isRubricScored) {
      return item.rawScore !== null && item.rawScore !== undefined &&
             item.rawScore < AP_BIO_LAB_CONCERN_BELOW;
    }

    // Scored assessments are reported for information, not as work owed.
    if (item.isQuizOrWebAssign) return false;

    return item.value === 'Missing' || item.value === 'Incomplete';
  });
}

/**
 * True for columns marked on the 0-4 rubric rather than as complete/missing:
 * Chemistry formatives and summatives, and AP Biology labs.
 *
 * This matters before display: the ordinary mapping turns a 1 into "Complete"
 * and a 0 into "Missing", which is wrong for a rubric where 1 means Emerging
 * and 0 means Not Yet Evident.
 */
function isRubricScoredColumn(subjectName, header, category, finalName, isSummativeStandard) {
  const text = (String(header || "") + " " + String(category || "")).toLowerCase();
  const name = String(finalName || "").toLowerCase().trim();

  if (subjectName === "AP Biology") {
    return name.indexOf('lab') === 0 || /\blabs?\b/.test(text) || /\blabs?\b/.test(name);
  }
  if (subjectName === "Chemistry" || subjectName === "XL Chemistry") {
    return !!isSummativeStandard ||
           text.indexOf('formative') > -1 || text.indexOf('summative') > -1 ||
           name.indexOf('formative') > -1 || name.indexOf('summative') > -1;
  }
  return false;
}

/**
 * Builds the per-column category labels from the category header row.
 *
 * Labels are filled rightwards so a merged header spanning several assignment
 * columns applies to all of them. Anything sitting above the roster block is
 * cleared first: a label like "Admin" over Name/Preferred Name/Email describes
 * those columns, and must not flow into the first assignment column beside it.
 *
 * @param {Array} categoryRow - the category header row's display values
 * @param {number} lastRosterCol - index of the final roster column
 */
function resolveCategoryRow(categoryRow, lastRosterCol) {
  const categories = (categoryRow || []).map(function (v) {
    return (v === null || v === undefined) ? "" : String(v);
  });

  for (let i = 0; i <= lastRosterCol && i < categories.length; i++) categories[i] = "";

  for (let i = 1; i < categories.length; i++) {
    if (categories[i] === "" && categories[i - 1] !== "") {
      categories[i] = categories[i - 1];
    }
  }
  return categories;
}

/**
 * Walks the gradebook below the header rows and splits it into students and
 * class-section dividers.
 *
 * A row counts as a student only when it has a name AND some evidence of being
 * a real person: an email, a parent email, any graded work, or a "Last, First"
 * style name. A row with text but none of that evidence is a class heading, so
 * it becomes the section label for the students beneath it instead of being
 * reported as a student. Section labels also come from a per-row Section
 * column when the gradebook uses one.
 */
function scanGradebookRoster(sheet) {
  const data = sheet.getDataRange().getDisplayValues();

  // Locate the roster header wherever it sits; data begins on the next row.
  const headerRow = findRosterHeaderRow(data);
  const FIRST_DATA_ROW = headerRow + 1;

  const checkboxCols = findCheckboxColumns(data, FIRST_DATA_ROW);
  const cols = findRosterColumns(data, headerRow, checkboxCols);

  // Roster and checkbox columns are excluded when looking for graded work.
  const reserved = {};
  [cols.name, cols.email, cols.parentEmail, cols.section].forEach(function (c) {
    if (c > -1) reserved[c] = true;
  });
  Object.keys(checkboxCols).forEach(function (c) { reserved[c] = true; });

  const styles = readBannerStyles(sheet, data.length, [0, cols.name]);

  const students = [];
  const sectionOrder = [];       // sheet order, not alphabetical
  let currentSection = "";
  let dividerCount = 0;

  const noteSection = function (label) {
    if (label && sectionOrder.indexOf(label) === -1) sectionOrder.push(label);
  };

  const isBannerRow = function (r) {
    const keys = Object.keys(styles);
    for (let i = 0; i < keys.length; i++) {
      const s = styles[keys[i]];
      if (s.bg[r] && isBannerCell(s.bg[r][0], s.font[r][0])) return true;
    }
    return false;
  };

  for (let r = FIRST_DATA_ROW; r < data.length; r++) {
    const row = data[r];
    const name = rosterCell(row, cols.name);
    const sectionCell = rosterCell(row, cols.section);
    const email = rosterCell(row, cols.email);
    const parentEmail = rosterCell(row, cols.parentEmail);

    // Anything outside the roster/checkbox columns counts as graded work.
    let gradeCount = 0;
    for (let c = 0; c < row.length; c++) {
      if (reserved[c]) continue;
      if (row[c] && String(row[c]).trim() !== "") gradeCount++;
    }

    // Truly empty rows only. A banner's label can live in a reserved column
    // (a merged cell in Column A), so gradeCount alone cannot decide this.
    if (firstNonEmptyCell(row) === "") continue;

    // Summary/statistics and leftover header rows are neither students nor dividers.
    if (SUMMARY_ROW_WORDS.test(name) || SUMMARY_ROW_WORDS.test(sectionCell)) continue;
    if (name === '0' || name === 'Student' || name === 'Preferred Name' || name.includes('Name:')) continue;

    // A row styled as a banner is a class divider no matter what it contains -
    // these often carry a headcount or a formula alongside the label.
    const styledDivider = isBannerRow(r);

    const looksLikePerson = name.indexOf(',') > -1 && !SECTION_LABEL_WORDS.test(name);
    const isStudent = !styledDivider && name !== "" && (
      email.indexOf('@') > -1 ||
      parentEmail.indexOf('@') > -1 ||
      gradeCount > 0 ||
      looksLikePerson);

    if (!isStudent) {
      // A row with data but no name is a standards / date / filter row, not a
      // class divider - unless it is styled as a banner.
      let label = "";
      if (styledDivider) label = name || sectionCell || firstNonEmptyCell(row);
      else if (gradeCount === 0) label = name || sectionCell;

      if (label && !isBooleanText(label)) {
        currentSection = label;
        dividerCount++;
        noteSection(label);
      }
      continue;
    }

    const section = (!isBooleanText(sectionCell) && sectionCell) || currentSection || "Ungrouped";
    noteSection(section);

    students.push({
      row: r,
      name: name,
      email: email,
      parentEmail: parentEmail,
      section: section,
      isMismatch: hasEmailNameMismatch(name, email)
    });
  }

  return {
    students: students,
    sections: sectionOrder,
    cols: cols,
    dividerCount: dividerCount,
    headerRow: headerRow,
    firstDataRow: FIRST_DATA_ROW
  };
}

/**
 * True for the display text Sheets gives a checkbox cell.
 */
function isBooleanText(value) {
  const upper = String(value || "").trim().toUpperCase();
  return upper === "TRUE" || upper === "FALSE";
}

/**
 * Groups scanned students into ordered, id-tagged sections for the selector UI.
 * Sections keep their sheet order; students are alphabetised within a section.
 */
function groupStudentsBySection(students, sectionOrder) {
  const groups = [];
  const byName = {};

  const ensureGroup = function (secName) {
    if (!byName[secName]) {
      byName[secName] = { id: 'sec-' + groups.length, name: secName, students: [] };
      groups.push(byName[secName]);
    }
    return byName[secName];
  };

  sectionOrder.forEach(ensureGroup);

  students.forEach(function (s) {
    const group = ensureGroup(s.section);
    s.sectionId = group.id;
    group.students.push(s);
  });

  const populated = groups.filter(function (g) { return g.students.length > 0; });
  populated.forEach(function (g) {
    g.students.sort(function (a, b) { return a.name.localeCompare(b.name); });
  });
  return populated;
}

/**
 * Scans the sheet and opens the Student Selector Dialog
 */
function showStudentSelector(mode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const roster = scanGradebookRoster(sheet);

  if (roster.students.length === 0) {
    SpreadsheetApp.getUi().alert("No students found. Check your Gradebook format.");
    return;
  }

  const sections = groupStudentsBySection(roster.students, roster.sections);

  // Offering "Parent" destinations is pointless when the sheet holds no parent
  // addresses, whether the column is absent or present but entirely empty.
  const hasParentEmails = roster.students.some(function (s) {
    return s.parentEmail.indexOf('@') > -1;
  });

  // Generate and Show UI
  const html = buildStudentSelectorHtml(sections, mode, hasParentEmails);
  SpreadsheetApp.getUi().showModalDialog(html.setWidth(600).setHeight(700), 'Student Selector');
}

/**
 * CORE PROCESSOR
 */
function runReportBatch(mode, rowIndices, emailDest) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName().toLowerCase();

    let subjectName = "Grade";
    let semester = "Semester 1";
    if (sheetName.includes("sem 2")) semester = "Semester 2";
    if (sheetName.includes("ap bio")) subjectName = "AP Biology";
    else if (sheetName.includes("xl chem")) subjectName = "XL Chemistry";
    else if (sheetName.includes("chem")) subjectName = "Chemistry";

    const reportTitle = `${semester} ${subjectName} Progress Report`;

    return processGradebook(sheet, reportTitle, subjectName, mode, rowIndices, emailDest || 'both');
  } catch (e) {
    throw new Error("Line " + e.lineNumber + ": " + e.message);
  }
}

// --- REPORT GENERATION LOGIC ---
function processGradebook(sheet, titlePrefix, subjectName, mode, targetRows, emailDest) {
  const data = sheet.getDataRange().getDisplayValues();
  const backgrounds = sheet.getDataRange().getBackgrounds();
  const fontColors = sheet.getDataRange().getFontColors();
  const notes = sheet.getDataRange().getNotes();

  const headerRowIndex = 1;
  const categoryRowIndex = 0;
  const standardsRowIndex = 2;

  // Assignment headers keep their fixed rows, but the roster columns are found
  // by label: Name is not always Column B, and the Email column may sit in a
  // header row of its own. Getting these wrong produces nameless reports and
  // renders the email column as if it were an assignment.
  const rosterCols = resolveRosterColumns(data);
  const nameColIndex = rosterCols.name;
  const lastRosterCol = rosterCols.lastRosterCol;

  const coolMessages = [
    // Original fun puns
    "🐳 done! You're doing swimmingly! 🌊", "You are the 🐝's 🦵s! (The bee's knees!) 🍯", "You 🍩 have any missing work! Sweet! 🍩",
    "🌮 'bout a great job! You crushed it! 🌮", "You are 👻-tacular! (No missing work to haunt you!)", "I'm not 🦁, you did a great job! 🦁",
    "You are on 🔥! (Metaphorically, please don't pull the alarm) 🚒", "You're a 🌟! Don't let anyone dim your shine.",
    "I am officially creating a fan club for your gradebook.", "Your organizational skills are terrifyingly good.",
    "I checked twice. You really did everything.", "You are legally required to high-five yourself right now.",
    "This report is boring in the best way possible. Nothing missing!", "You have defeated the final boss of procrastination.",
    "I am applause. Just pure applause.", "Go buy yourself a treat. You earned it.", "Zero missing assignments. Is this real life?",
    "You are a productivity wizard. Teach me your ways.", "Your parents are probably going to frame this report.",
    "I tried to find a mistake. I failed. Good job.", "You are the MVP of turning things in on time.", "Your work ethic is legendary.",
    "Absolute perfection. No notes.", "You are winning at school right now.", "This is the gold standard of studenting.",
    "You are a homework ninja. Silent, deadly, effective.", "Boom. Done. Everything submitted.", "You are officially on top of your game.",
    "You're 🍇 (grape) at this! Keep it up!", "Don't stop be-🍃-ing (believing) in yourself!", "You are 🦖-mite (dynamite)!",
    "Keep up the 🥚-cellent work!", "You really 🐮-ved (moved) mountains with this effort!", "You're kind of a big 🥒 (dill)!",
    "Lettuce 🥬 celebrate your success!", "You've got a lot of 🍕-zazz (pizzazz)!", "I'm 🍌 (bananas) about your work ethic!",
    "You are one smart 🍪 (cookie)!", "Gouda job! 🧀 (Cheesy, I know)", "You are s-🧊 (ice) cool with no missing work!",
    "You rose 🌹 to the occasion!", "You are ⚓️ (anchor)-ed in excellence!", "Sending you high-fives and 🌮 (tacos)!",
    "Orange 🍊 you glad you did all your work?", "You are pear-fect 🍐!", "Time to shell-ebrate! 🐢",
    "You are un-be-🍃-able!", "You are a fungal/fungi to have in class! 🍄", "I Dig ⛏️ your work ethic!",
    // Chemistry puns
    "You've got great chemistry with your assignments! ⚗️", "NaCl job! (That's a salt, but you're not basic!) 🧂",
    "You have all the right elements for success! 🔬", "Your work is Au-some! (That's gold!) 🥇",
    "You're in your element! Periodic table would be proud. 📊", "No missing work? That's a positive reaction! ⚛️",
    "You've bonded well with your responsibilities! 🔗", "Your grade is noble... like a noble gas! 💨",
    "All your work is accounted for - perfectly balanced, as all equations should be! ⚖️",
    "You're sodium funny... Na just kidding, you're brilliant! 🧪",
    // Biology puns
    "You've really evolved as a student! 🧬", "Cell-ebrate good times! All work complete! 🦠",
    "You're DNA-mite! (Get it? Dynamite?) 💥", "Mitosis be the best report I've seen today! 🔬",
    "You've got good genes... for turning in work! 👖🧬", "This is un-CELL-ievably good work! 🔬",
    "You're not just surviving, you're thriving! Natural selection approves! 🌿",
    "ATP-solutely amazing work! You've got energy! ⚡", "Your work ethic is phenotypically perfect! 🧬",
    "Organism-ized and on point! 🦎", "You've adapted well to the assignment environment! 🐸",
    // Physics puns
    "You have great potential (energy)! ⚡", "Your momentum is unstoppable! 🚀",
    "Newton would be proud - you stayed in motion! 🍎", "You've overcome all resistance! Ohm my! ⚡",
    "Your work is relatively excellent! Einstein approves! 🧠", "You're accelerating toward success! 📈",
    "Zero friction between you and your assignments! 🛷", "You've reached terminal velocity of awesomeness! 🪂",
    "Watt a great job! You're fully charged! 🔋", "Your grades are looking pretty stellar! ⭐",
    // General science puns
    "Scientifically speaking, you're crushing it! 🔬", "Hypothesis confirmed: You're awesome! 📋",
    "Your data supports the conclusion that you rock! 📊", "Lab-solutely fantastic work! 🥽",
    "You've completed all your trials successfully! 🧫", "Your results are reproducible: consistently great! 📈",
    "Control group? More like you're IN control! 🎮", "You've got the right formula for success! 📝"
  ];

  let summativeStartColIndex = -1;
  if (subjectName === "Chemistry") {
    for (let r = 0; r < Math.min(data.length, 10); r++) {
      const rowVals = data[r];
      const foundIdx = rowVals.findIndex(cell => cell && String(cell).toLowerCase().trim() === 'summatives');
      if (foundIdx !== -1) { summativeStartColIndex = foundIdx; break; }
    }
  }

  const headers = data[headerRowIndex];
  // Fill-right Categories for merged headers, ignoring any label that belongs
  // to the roster block rather than to the assignments.
  const categories = resolveCategoryRow(data[categoryRowIndex], lastRosterCol);
  let emailColIndex = rosterCols.email;
  let parentEmailColIndex = rosterCols.parentEmail;

  // Search headers (only for whatever the roster scan could not resolve)
  for (let i = 0; i < headers.length; i++) {
    if (!headers[i]) continue;
    const text = headers[i].toLowerCase();
    if (text.includes('parent') || text.includes('guardian')) {
      if (parentEmailColIndex === -1) parentEmailColIndex = i;
    } else if (text.includes('email') && !text.includes('parent') && !text.includes('guardian')) {
      if (emailColIndex === -1) emailColIndex = i;
    }
  }

  // Fallback to categories if not found in headers
  if (emailColIndex === -1 && categories) {
    for (let i = 0; i < categories.length; i++) {
      if (!categories[i]) continue;
      const text = categories[i].toLowerCase();
      if (text.includes('email') && !text.includes('parent') && !text.includes('guardian')) {
        emailColIndex = i;
        break;
      }
    }
  }
  if (parentEmailColIndex === -1 && categories) {
    for (let i = 0; i < categories.length; i++) {
      if (!categories[i]) continue;
      const text = categories[i].toLowerCase();
      if (text.includes('parent') || text.includes('guardian')) {
        parentEmailColIndex = i;
        break;
      }
    }
  }

  // Fallback for Student Email: if not found by explicit search but there's another column containing "email"
  if (emailColIndex === -1) {
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && headers[i].toLowerCase().includes('email') && i !== parentEmailColIndex) {
        emailColIndex = i;
        break;
      }
    }
  }
  if (emailColIndex === -1 && categories) {
    for (let i = 0; i < categories.length; i++) {
      if (categories[i] && categories[i].toLowerCase().includes('email') && i !== parentEmailColIndex) {
        emailColIndex = i;
        break;
      }
    }
  }

  let cutoffColIndex = headers.length;
  const headerBgColors = backgrounds[headerRowIndex];
  for (let i = lastRosterCol + 1; i < headers.length; i++) {
    if (headerBgColors[i] === '#000000') { cutoffColIndex = i; break; }
  }

  const standards = (data.length > standardsRowIndex) ? data[standardsRowIndex] : [];
  const headerFontColors = fontColors[headerRowIndex];

  const columnDefs = headers.map((header, i) => {
    if (i >= cutoffColIndex) return null;
    return {
      id: i,
      rawHeader: header,
      standard: (standards && standards[i]) ? standards[i] : null,
      rawCategory: (categories && categories[i]) ? categories[i] : null,
      bgColor: headerBgColors[i],
      fontColor: headerFontColors[i],
      finalName: header,
      finalCategory: "",
      isSummativeStandard: false,
      isQuizOrWebAssign: false,
      isSummaryStat: false
    };
  });

  // Left empty so the table renderers fall back to their neutral "General"
  // heading; an assignment column with no category above it had been picking
  // up whatever label happened to sit to its left.
  let lastCategory = "";
  columnDefs.forEach(col => {
    if (!col || col.id <= lastRosterCol) return;
    if (col.rawCategory && col.rawCategory.trim() !== "") lastCategory = col.rawCategory.trim();
    col.finalCategory = lastCategory;
  });

  let lastSeenHeader = "";
  columnDefs.forEach(col => {
    if (!col || col.id <= lastRosterCol) return;
    let currentHeader = col.rawHeader ? col.rawHeader.trim() : "";
    if (currentHeader !== "" && subjectName === "Chemistry") lastSeenHeader = currentHeader;

    const isStandardCol = col.standard && col.standard.trim() !== "" && !col.standard.toLowerCase().includes("standards") && !col.standard.toLowerCase().includes("admin");

    if (subjectName === "Chemistry" && isStandardCol) {
      const prefix = (currentHeader !== "") ? currentHeader : lastSeenHeader;
      col.finalName = `${prefix}, ${col.standard}`;
      if (prefix.toLowerCase().includes("summative") || (summativeStartColIndex !== -1 && col.id >= summativeStartColIndex)) {
        col.isSummativeStandard = true;
      }
    } else if (currentHeader === "" && col.finalCategory && col.finalCategory.trim() !== "") {
      col.finalName = col.finalCategory;
    } else {
      col.finalName = currentHeader;
    }

    // --- NEW: FLEXIBLE GROUPED ASSESSMENT & SUMMARY DETECTION ---
    const lowerHeader = col.rawHeader ? col.rawHeader.toLowerCase().trim() : "";
    const lowerStandard = col.standard ? col.standard.toLowerCase().trim() : "";
    const lowerCategory = col.finalCategory ? col.finalCategory.toLowerCase().trim() : "";

    // Heuristics
    const headerScoreKeywords = ['raw', 'score', 'percent', '%', 'letter', 'grade', 'points', 'pts'];
    const summaryKeywords = ['completion', 'missing', 'participation', 'rate'];

    const matchesAssessment = matchesAssessmentColumn(col.rawHeader, col.standard, col.finalCategory);
    const matchesScoreHeader = headerScoreKeywords.some(k => lowerHeader.includes(k));

    // Logic: It's a grouped assessment if Row 3 OR Category has a keyword OR (Row 3 exists/is used AND Row 2 looks like a score header)
    if (matchesAssessment || (lowerStandard !== "" && matchesScoreHeader)) {
      col.isQuizOrWebAssign = true;
      // If a specific name exists in the Standards row (row 3), preserve it as the category for grouping
      // BUT ONLY if it isn't empty and doesn't look like a Chemistry Standard ID (DCI, SEP, AC, etc)
      if (col.standard && col.standard.trim() !== "") {
        const std = col.standard.trim();
        const isChemStandard = /^(DCI|SEP|AC|CC)\./i.test(std);
        if (!isChemStandard) col.finalCategory = std;
      }
      if (currentHeader !== "") col.finalName = currentHeader;
    }

    // Checking for Summary Stats
    if (summaryKeywords.some(k => lowerHeader.includes(k))) {
      col.isSummaryStat = true;
    }

    col.isRubricScored = !col.isSummaryStat && isRubricScoredColumn(
      subjectName, col.rawHeader, col.finalCategory, col.finalName, col.isSummativeStandard);
  });

  // --- PREPARE OUTPUT ---
  let doc, docId, docBody;
  let previewHtml = "";
  let previewCount = 0;
  let processedCount = 0;

  if (mode === 'drive') {
    const docName = `${sheetName} - Selected Reports`;
    doc = DocumentApp.create(docName);
    docId = doc.getId();
    docBody = doc.getBody();
    docBody.setMarginTop(36).setMarginBottom(36).setMarginLeft(36).setMarginRight(36);
  }

  // --- PROCESS SELECTED ROWS ---
  for (let i = 0; i < targetRows.length; i++) {
    const r = targetRows[i];
    const row = data[r];
    const studentName = row[nameColIndex];
    const rowNotes = notes[r];
    const rowBgColors = backgrounds[r];
    const rowFontColors = fontColors[r];

    // Gather Data
    let reportRows = [];
    columnDefs.forEach((col, idx) => {
      if (!col || idx <= lastRosterCol || col.rawHeader === 'Assignment' || col.rawHeader === 'Preferred Name') return;

      const rawLower = col.rawHeader ? col.rawHeader.toLowerCase() : "";
      const finalLower = col.finalName ? col.finalName.toLowerCase() : "";
      if (rawLower.includes("excused") || rawLower.includes("i's and m's") || finalLower.includes("excused") || finalLower.includes("i's and m's") || idx === emailColIndex) return;
      if (!col.finalName && !col.standard) return;

      let value = row[idx];
      let displayValue = value;
      let isIssue = false;
      let isExempt = false;
      let rawScore = null;

      if (value) {
        const valStr = String(value).trim();
        const lowerVal = valStr.toLowerCase();
        const rawHeaderLower = col.rawHeader ? col.rawHeader.toLowerCase() : "";

        // A rubric score must be read before the complete/missing mapping,
        // which would otherwise turn 1 into "Complete" and 0 into "Missing".
        const rubricLabel = col.isRubricScored ? rubricLabelFor(valStr) : null;

        if (rubricLabel !== null) {
          displayValue = rubricLabel;
          rawScore = Number(valStr);
          if (rawScore === 0) isIssue = true;      // no evidence of the standard yet
        }
        else if ((lowerVal === 'true' || valStr === '1') && !col.isSummaryStat) displayValue = 'Complete';
        else if ((valStr === '0' || lowerVal === 'm' || lowerVal === 'false') && !col.isSummaryStat) { displayValue = 'Missing'; isIssue = true; }
        else if (lowerVal === 'ex') { displayValue = 'Exempt'; isExempt = true; }
        else if (lowerVal === 'i') { displayValue = 'Incomplete'; isIssue = true; }
        // Treat 0.5 or .5 as Incomplete for Activities (AC:) and InfoDocs (ID:)
        else if ((valStr === '0.5' || valStr === '.5') && !col.isSummaryStat) {
          const isActivityOrInfoDoc = rawHeaderLower.startsWith('ac:') || rawHeaderLower.startsWith('id:');
          if (isActivityOrInfoDoc) { displayValue = 'Incomplete'; isIssue = true; }
        }

        // Uses the rubric flag rather than a name prefix: a lab headed
        // "Topic Quest Lab 1" does not start with "lab" and was never matched.
        if (subjectName === "AP Biology" && col.isRubricScored) {
          const numVal = parseFloat(valStr);
          if (!isNaN(numVal) && numVal < AP_BIO_LAB_CONCERN_BELOW) isIssue = true;
        }
      } else {
        displayValue = "-";
      }

      let shouldReport = false;
      if (isIssue || isExempt) shouldReport = true;
      if (subjectName === "Chemistry" && col.isSummativeStandard) shouldReport = true;
      if (col.isQuizOrWebAssign && displayValue !== "-" && String(displayValue).trim() !== "") shouldReport = true;
      if (col.isSummaryStat) shouldReport = true;

      // A column with no header and no category has nothing to show a student.
      if (!col.finalName || String(col.finalName).trim() === "") shouldReport = false;

      if (shouldReport) {
        reportRows.push({
          category: col.finalCategory,
          name: expandAssignmentPrefix(col.finalName),
          value: displayValue,
          rawScore: rawScore,
          isRubricScored: !!col.isRubricScored,
          note: rowNotes[idx],
          bgColor: col.bgColor,
          fontColor: col.fontColor,
          rowBg: rowBgColors[idx],
          rowFont: rowFontColors[idx],
          isQuizOrWebAssign: col.isQuizOrWebAssign,
          isSummaryStat: col.isSummaryStat,
          isSummativeStandard: col.isSummativeStandard
        });
      }
    });

    const hasActualMissingWork = hasOutstandingWork(reportRows, subjectName);

    const hasSummaryIssue = reportRows.some(item =>
      item.isSummaryStat &&
      (item.name.toLowerCase().includes("missing") || item.name.toLowerCase().includes("incomplete")) &&
      parseFloat(String(item.value).replace('%', '')) > 0
    );

    const isStudentInTrouble = hasActualMissingWork || hasSummaryIssue;

    // --- ACTION HANDLERS ---

    if (mode === 'drive') {
      if (processedCount > 0) docBody.appendPageBreak();
      renderToDoc(docBody, titlePrefix, studentName, reportRows, isStudentInTrouble, subjectName, coolMessages);
      processedCount++;
      if (processedCount % 5 === 0) { doc.saveAndClose(); doc = DocumentApp.openById(docId); docBody = doc.getBody(); }
    }

    else if (mode === 'preview') {
      if (previewCount >= 10) continue;
      const studentEmail = (emailColIndex > -1) ? row[emailColIndex] : "";
      const parentEmail = (parentEmailColIndex > -1) ? row[parentEmailColIndex] : "";

      // Test for a real address: a missing column used to yield the non-empty
      // placeholder "No Parent Email", which previewed as a sendable parent.
      const showStudent = (emailDest === 'student' || emailDest === 'both');
      const showParent = (emailDest === 'parent' || emailDest === 'both') &&
                         parentEmail.indexOf('@') > -1;
      const studentEmailLabel = studentEmail.indexOf('@') > -1 ? studentEmail : "No Student Email";

      if (showStudent) {
        const htmlBody = generateHtmlReport(titlePrefix, studentName, reportRows, isStudentInTrouble, subjectName, coolMessages, false);
        previewHtml += `<div class="preview-box" style="margin-bottom: 40px; border-bottom: 4px solid #ccc; padding-bottom: 40px; background-color: #fcfcfc; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                          <div class="preview-header" style="background-color: #1a73e8; color: white; padding: 8px 12px; font-weight: bold; border-radius: 4px 4px 0 0; margin-bottom: 15px;">STUDENT PREVIEW ${previewCount + 1}: ${studentName} (${studentEmailLabel})</div>
                          ${htmlBody}
                        </div>`;
      }
      
      if (showParent) {
        const htmlBody = generateHtmlReport(titlePrefix, studentName, reportRows, isStudentInTrouble, subjectName, coolMessages, true);
        previewHtml += `<div class="preview-box" style="margin-bottom: 40px; border-bottom: 4px solid #ccc; padding-bottom: 40px; background-color: #fcfcfc; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                          <div class="preview-header" style="background-color: #34a853; color: white; padding: 8px 12px; font-weight: bold; border-radius: 4px 4px 0 0; margin-bottom: 15px;">PARENT PREVIEW ${previewCount + 1}: For Parent of ${studentName} (${parentEmail})</div>
                          ${htmlBody}
                        </div>`;
      }
      previewCount++;
    }

    else if (mode === 'email') {
      const studentEmail = (emailColIndex > -1) ? row[emailColIndex] : "";
      const parentEmail = (parentEmailColIndex > -1) ? row[parentEmailColIndex] : "";
      
      const sendToStudent = (emailDest === 'student' || emailDest === 'both') && studentEmail && studentEmail.includes('@');
      const sendToParent = (emailDest === 'parent' || emailDest === 'both') && parentEmail && parentEmail.includes('@');

      if (sendToStudent) {
        const htmlBody = generateHtmlReport(titlePrefix, studentName, reportRows, isStudentInTrouble, subjectName, coolMessages, false);
        try {
          const emailOptions = {
            to: studentEmail,
            subject: `${titlePrefix} - ${studentName}`,
            htmlBody: htmlBody
          };
          const replyTo = getReplyToEmail();
          if (replyTo) emailOptions.replyTo = replyTo;
          MailApp.sendEmail(emailOptions);
          processedCount++;
        } catch (e) { Logger.log(`Email error ${studentName}: ${e.message}`); }
      }

      if (sendToParent) {
        const htmlBody = generateHtmlReport(titlePrefix, studentName, reportRows, isStudentInTrouble, subjectName, coolMessages, true);
        try {
          const emailOptions = {
            to: parentEmail,
            subject: `${titlePrefix} - Parent/Guardian Progress Report for ${studentName}`,
            htmlBody: htmlBody
          };
          const replyTo = getReplyToEmail();
          if (replyTo) emailOptions.replyTo = replyTo;
          MailApp.sendEmail(emailOptions);
          processedCount++;
        } catch (e) { Logger.log(`Parent email error ${studentName}: ${e.message}`); }
      }
    }
  }

  // --- FINALIZE ---
  if (mode === 'drive') {
    doc.saveAndClose();
    const docFile = DriveApp.getFileById(docId);
    const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const parentFolders = DriveApp.getFileById(ssId).getParents();
    const folder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
    docFile.moveTo(folder);
    SpreadsheetApp.getUi().alert(`Generated ${processedCount} reports in Google Drive.`);
    return null;
  }
  else if (mode === 'email') {
    SpreadsheetApp.getUi().alert(`Success! Sent ${processedCount} emails.`);
    return null;
  }
  else if (mode === 'preview') {
    return previewHtml || "No data to preview.";
  }
}

// --- RENDER FUNCTIONS ---

function renderToDoc(body, title, studentName, rows, hasMissing, subjectName, messages) {
  const titleStyle = {};
  titleStyle[DocumentApp.Attribute.FONT_SIZE] = 16;
  titleStyle[DocumentApp.Attribute.BOLD] = true;
  titleStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';

  const p = body.appendParagraph(`${title}\n${studentName}`);
  p.setAttributes(titleStyle);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  if (!hasMissing) {
    const msg = messages[Math.floor(Math.random() * messages.length)];
    const pMsg = body.appendParagraph(`\n\n${msg}`);
    pMsg.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(14).setForegroundColor('#2E7D32');
    const pSub = body.appendParagraph("\nStatus: No missing formative work. Nothing is owed at this time.");
    pSub.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(10).setForegroundColor('#555555');

    // Stats for perfect students
    printSummaryStatsDoc(body, rows);
    printEncouragementDoc(body, rows, false);

    if (rows.length > 0) {
      const filteredRows = rows.filter(r => !r.isSummaryStat);

      if (subjectName === "Chemistry") {
        const summativeRows = filteredRows.filter(r => r.isSummativeStandard);
        const otherRows = filteredRows.filter(r => !r.isSummativeStandard);

        if (otherRows.length > 0) printGroupedTableDoc(body, otherRows);
        if (summativeRows.length > 0) {
          body.appendParagraph("\nSummative Standard Mastery:\n").setBold(true).setFontSize(12);
          printGroupedTableDoc(body, summativeRows);
        }
      } else {
        if (filteredRows.length > 0) printGroupedTableDoc(body, filteredRows);
      }
    }
    printLabStandardNoteDoc(body, rows, subjectName, false);
  } else {
    // Has Missing Work
    const filteredRows = rows.filter(r => !r.isSummaryStat);
    if (subjectName === "Chemistry") {
      const summativeRows = filteredRows.filter(r => r.isSummativeStandard);
      const otherRows = filteredRows.filter(r => !r.isSummativeStandard);

      if (otherRows.length > 0) printGroupedTableDoc(body, otherRows);
      if (summativeRows.length > 0) {
        body.appendParagraph("\nSummative Standard Mastery:\n").setBold(true).setFontSize(12);
        printGroupedTableDoc(body, summativeRows);
      }
    } else {
      printGroupedTableDoc(body, filteredRows);
    }
    // Stats for other students
    printSummaryStatsDoc(body, rows);
    printEncouragementDoc(body, rows, false);
    printLabStandardNoteDoc(body, rows, subjectName, false);
  }
}

function printSummaryStatsDoc(body, rows) {
  const stats = rows.filter(r => r.isSummaryStat);
  if (stats.length === 0) return;

  body.appendParagraph("\nParticipation Metrics").setBold(true).setFontSize(11).setForegroundColor('#444444');
  const table = body.appendTable();
  table.setBorderColor('#bbbbbb');

  // Doc table widths are tricky, we just rely on auto for now or set specifically if needed
  stats.forEach(item => {
    const tr = table.appendTableRow();
    tr.appendTableCell(item.name).setBackgroundColor('#f9f9f9').setFontSize(9).setBold(true).setWidth(230);
    const valCell = tr.appendTableCell(item.value);
    const para = valCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(9);

    // Color Logic
    const nameLower = item.name.toLowerCase();
    const valNum = parseFloat(String(item.value).replace('%', ''));
    if (!isNaN(valNum)) {
      if (nameLower.includes("incomplete") || nameLower.includes("missing")) {
        // Lower is better. 0 is best.
        if (valNum === 0) para.setForegroundColor('#2E7D32'); // Green
        else if (valNum > 0) para.setForegroundColor('#c62828'); // Red
      } else if (nameLower.includes("completion")) {
        // Higher is better.
        if (valNum === 100 || valNum === 1) para.setForegroundColor('#2E7D32');
        else if (valNum < 100) para.setForegroundColor('#ef6c00'); // Orange
      }
    }
  });
}

function printGroupedTableDoc(body, rows) {
  const groups = {};
  const order = [];
  rows.forEach(r => {
    if (r.isSummaryStat) return; // Skip stats
    const c = r.category || "General";
    if (!groups[c]) { groups[c] = []; order.push(c); }
    groups[c].push(r);
  });

  order.forEach(cat => {
    body.appendParagraph(`\n${cat}`).setBold(true).setFontSize(11).setForegroundColor('#444444');
    const table = body.appendTable();
    table.setBorderColor('#bbbbbb');
    const header = table.appendTableRow();
    header.appendTableCell("Assignment").setBackgroundColor('#EFEFEF').setBold(true).setFontSize(9).setWidth(230);
    const scoreHeader = header.appendTableCell("Score");
    scoreHeader.setBackgroundColor('#EFEFEF').setBold(true).setFontSize(9);
    scoreHeader.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    groups[cat].forEach(item => {
      const tr = table.appendTableRow();
      const c1 = tr.appendTableCell(item.name);
      c1.setBackgroundColor(item.bgColor !== '#ffffff' ? item.bgColor : '#ffffff');
      c1.getChild(0).asParagraph().setFontSize(9).setBold(true).setForegroundColor(item.fontColor);
      c1.setPaddingTop(2).setPaddingBottom(2);
      const c2 = tr.appendTableCell(item.value);
      c2.setBackgroundColor(item.rowBg !== '#ffffff' ? item.rowBg : '#ffffff');
      c2.getChild(0).asParagraph().setFontSize(9).setForegroundColor(item.rowFont).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      c2.setPaddingTop(2).setPaddingBottom(2);
    });
  });
}

function generateHtmlReport(title, studentName, rows, hasMissing, subjectName, messages, isParent) {
  let html = `<div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; line-height: 1.6;">`;
  html += `<h2 style="text-align: center; color: #222; margin-bottom: 5px;">${title}</h2>`;
  html += `<h3 style="text-align: center; color: #555; margin-top: 0; margin-bottom: 20px;">${studentName}</h3>`;

  if (isParent) {
    // Parent-specific welcoming and explanatory note
    html += `<div style="background-color: #f8f9fa; border-left: 4px solid #1a73e8; padding: 15px; margin-bottom: 20px; border-radius: 0 4px 4px 0; font-size: 13px;">`;
    html += `<p style="margin-top: 0; font-weight: bold; color: #1a73e8; font-size: 14px;">Dear Parent / Guardian,</p>`;
    html += `<p style="margin-bottom: 10px;">This academic progress report is provided to help keep you informed about your student's status in our class. Below, you will find details on their assignments, assessments, and overall participation metrics.</p>`;
    html += `<p style="margin-bottom: 10px;"><strong>What this report shows:</strong> It details specific assignments completed, scores earned on assessments, and any outstanding or incomplete formative work. A status of <strong>'Missing'</strong> or <strong>'Incomplete'</strong> indicates that the assignment was not submitted, which can have a significant impact on your student's learning and grade.</p>`;
    html += `<p style="margin-bottom: 10px;"><strong>How you can help:</strong> We encourage you to take a few minutes to talk to your student about this report and ask them about what they are learning in our course. Active conversations at home can be incredibly supportive of their academic growth!</p>`;
    html += `<p style="margin-bottom: 0;">If you have any questions, concerns, or if we can support your student in any way, please feel free to reply directly to this email. I would love to hear from you!</p>`;
    html += `</div>`;
  } else if (!hasMissing) {
    const msg = messages[Math.floor(Math.random() * messages.length)];
    html += `<div style="text-align: center; margin: 20px 0; padding: 15px; background-color: #e8f5e9; border-radius: 5px;">`;
    html += `<h3 style="color: #2E7D32; margin: 0;">${msg}</h3>`;
    html += `<p style="color: #555; font-size: 12px; margin-top: 5px;">Status: No missing formative work. Nothing is owed at this time.</p></div>`;
  }

  html += generateHtmlSummaryStats(rows);
  html += generateHtmlEncouragement(rows, isParent);

  if (rows.length > 0) {
    const filteredRows = rows.filter(r => !r.isSummaryStat);

    if (subjectName === "Chemistry") {
      const summativeRows = filteredRows.filter(r => r.isSummativeStandard);
      const otherRows = filteredRows.filter(r => !r.isSummativeStandard);

      if (otherRows.length > 0) html += generateHtmlTables(otherRows);
      if (summativeRows.length > 0) {
        html += `<h4 style="margin-top: 20px; border-bottom: 1px solid #ccc; padding-bottom: 3px; color: #444;">Summative Standard Mastery:</h4>`;
        html += generateHtmlTables(summativeRows);
      }
    } else {
      html += generateHtmlTables(filteredRows);
    }
  }

  html += generateHtmlLabStandardNote(rows, subjectName, isParent);

  html += `<p style="font-size: 10px; color: #888; text-align: center; margin-top: 30px;">Generated by Gradebook Tools</p>`;
  html += `</div>`;
  return html;
}

// --- COMPLETION ENCOURAGEMENT ---------------------------------------------
// Reports carry a supportive note whenever completion is below 100%, scaled to
// how much work is outstanding. The tone is never punitive and never implies
// the situation is beyond saving; every tier ends with a way to reach out.

/**
 * Reads the completion percentage out of the participation stats.
 *
 * Handles both directions (a "Completion" stat and an "% Incomplete" one) and
 * both scales, since a gradebook may store 0.67 or 67. An inverse stat is only
 * trusted when it is clearly a percentage, so a raw count of missing
 * assignments is never mistaken for one.
 *
 * @return {number|null} 0-100, or null when no completion stat exists.
 */
function findCompletionPercent(rows) {
  const stats = rows.filter(r => r.isSummaryStat);
  let direct = null;
  let inverse = null;

  stats.forEach(item => {
    const name = String(item.name || "").toLowerCase();
    const raw = String(item.value == null ? "" : item.value).trim();
    const num = parseFloat(raw.replace('%', ''));
    if (isNaN(num)) return;

    // 1 means 100%, 0.x is a fraction, anything larger is already a percentage.
    const pct = (num === 1) ? 100 : (num > 0 && num < 1 ? num * 100 : num);

    const isCompletion = name.indexOf('completion') > -1 ||
      (name.indexOf('complete') > -1 && name.indexOf('incomplete') === -1);
    const looksLikePercent = raw.indexOf('%') > -1 || name.indexOf('%') > -1 ||
      name.indexOf('percent') > -1 || name.indexOf('rate') > -1;

    if (isCompletion) {
      if (direct === null) direct = pct;
    } else if ((name.indexOf('incomplete') > -1 || name.indexOf('missing') > -1) && looksLikePercent) {
      if (inverse === null) inverse = 100 - pct;
    }
  });

  const value = (direct !== null) ? direct : inverse;
  if (value === null) return null;
  return Math.max(0, Math.min(100, value));
}

/**
 * Builds the encouragement note for a report, or null when completion is at
 * 100% (or unknown), in which case the existing congratulations stand.
 */
function buildEncouragementNote(rows, isParent) {
  const pct = findCompletionPercent(rows);
  if (pct === null || pct >= 100) return null;

  const shown = (Math.round(pct * 10) / 10) + "%";
  const replyTo = getReplyToEmail();
  const contact = replyTo ? ` You can reach me at ${replyTo}.` : "";

  let note;
  if (pct >= 80) {
    note = {
      tier: 'minor',
      accent: '#1a73e8',
      background: '#e8f0fe',
      heading: "Nearly there — a small push will finish this",
      intro: isParent
        ? `Your student's completion is at ${shown}. That's strong work, with a short list still outstanding.`
        : `Your completion is at ${shown}. That's strong work, and what's left is a short list.`,
      steps: isParent
        ? ["Ask them to look through the table below for anything marked Missing or Incomplete.",
           "The oldest items are usually the quickest to close out.",
           "If anything is unclear, they're welcome to email me — or you can just reply to this message."]
        : ["Look through the table below for anything marked Missing or Incomplete.",
           "Start with the oldest one — those are usually the quickest to close out.",
           "If any of it is unclear, email me and I'll point you in the right direction."],
      closing: isParent
        ? `Nothing here is cause for concern. I'm glad to help them finish up if that would be useful.${contact}`
        : `You're in good shape. Reply to this email, or catch me in class, if you'd like a hand finishing up.${contact}`
    };
  } else if (pct >= 50) {
    note = {
      tier: 'moderate',
      accent: '#ef6c00',
      background: '#fff4e5',
      heading: "There's a real gap here, and it's very fixable",
      intro: isParent
        ? `Your student's completion is at ${shown}. Enough work is outstanding that a deliberate plan would help — and this is a very recoverable position.`
        : `Your completion is at ${shown}. Enough is outstanding that it's worth making a deliberate plan — and this is a very recoverable position.`,
      steps: isParent
        ? ["A conversation at home about what's outstanding is a good first step.",
           "Encourage them to pick two items to finish this week rather than all of it at once.",
           "They can email me or book a time with me, and you're very welcome to reply to this message."]
        : ["Go through the table below and write down every Missing or Incomplete item.",
           "Pick two to finish this week. Trying to do all of it at once is what makes it stall.",
           "Email me or book a time with me, and we'll work out the order together."],
      closing: isParent
        ? `Please don't hesitate to get in touch. I'd far rather connect early than late.${contact}`
        : `Please do reach out — I'd far rather hear from you early than late, and I'm happy to help you build the plan.${contact}`
    };
  } else {
    note = {
      tier: 'urgent',
      accent: '#c62828',
      background: '#fdecea',
      heading: "Let's sort this out together",
      intro: isParent
        ? `Your student's completion is at ${shown}. That's a significant amount of outstanding work, and I wanted you to hear it from me directly. This is a starting point rather than a final result — students do recover from here with support.`
        : `Your completion is at ${shown}. That's a lot of outstanding work, and I don't want you facing it on your own. This is a starting point, not a final result — students climb out of exactly this position, and there's a route through it.`,
      steps: isParent
        ? ["Tackling everything at once rarely works; a short, prioritized list does.",
           "I'd welcome a conversation this week — please reply to this email or book a time with me.",
           "Together we can choose a few items to start with and build from there."]
        : ["Don't try to tackle everything at once. That's what makes it feel impossible.",
           "Email me or book an appointment this week — it's the single most useful thing you can do right now.",
           "We'll choose a few items to start with and build momentum from there."],
      closing: isParent
        ? `Please reach out. The work is recoverable, and the earlier we begin, the more options your student has.${contact}`
        : `Please get in touch. The work is recoverable, and the earlier we start, the more options you'll have.${contact}`
    };
  }

  note.percent = pct;
  return note;
}

/**
 * Explains the standard AP Biology labs have to reach.
 *
 * Standards-based work is only fully claimed once every lab finishes the
 * semester at 4 / Meeting with Distinction, which a list of rubric words does
 * not convey on its own. Returns null for other subjects, and for reports that
 * contain no labs.
 */
function buildLabStandardNote(rows, subjectName, isParent) {
  if (subjectName !== "AP Biology") return null;

  const labs = rows.filter(function (r) { return r.isRubricScored && !r.isSummaryStat; });
  if (labs.length === 0) return null;

  const scored = labs.filter(function (r) {
    return r.rawScore !== null && r.rawScore !== undefined;
  });
  const below = scored.filter(function (r) { return r.rawScore < 4; });

  const subject = isParent ? "they" : "you";

  const note = {
    heading: "What lab scores need to reach",
    body: "Labs are marked on the 0\u20134 rubric. To support a full grade claim for " +
          "standards-based work, every lab needs to be sitting at 4 \u2014 Meeting with " +
          "Distinction \u2014 by the end of the semester."
  };

  if (scored.length === 0) {
    note.status = "";
    note.accent = '#5f6368';
    note.background = '#f8f9fa';
    return note;
  }

  if (below.length === 0) {
    note.accent = '#2E7D32';
    note.background = '#e8f5e9';
    note.status = labs.length === 1
      ? `The lab on this report is already at 4. Keep it there.`
      : `Every lab on this report is already at 4. Keep them there.`;
    return note;
  }

  const count = below.length === 1
    ? "1 lab on this report is"
    : `${below.length} labs on this report are`;

  note.accent = '#1a73e8';
  note.background = '#e8f0fe';
  note.status = isParent
    ? `${count} not there yet. These scores are not final until the end of the semester, ` +
      `so there is still time to revise them \u2014 ${subject} can ask me what a revision would need to show.`
    : `${count} not there yet. These scores are not final until the end of the semester, ` +
      `so there is still time to revise them \u2014 ask me what a revision would need to show.`;
  return note;
}

/**
 * Renders the lab standard note as HTML, or "" when it does not apply.
 */
function generateHtmlLabStandardNote(rows, subjectName, isParent) {
  const note = buildLabStandardNote(rows, subjectName, isParent);
  if (!note) return "";

  let html = `<div style="background-color: ${note.background}; border-left: 4px solid ${note.accent}; padding: 15px; margin: 20px 0; border-radius: 0 4px 4px 0; font-size: 13px;">`;
  html += `<p style="margin-top: 0; margin-bottom: 8px; font-weight: bold; color: ${note.accent}; font-size: 14px;">${note.heading}</p>`;
  html += `<p style="margin: 0;">${note.body}</p>`;
  if (note.status) html += `<p style="margin: 10px 0 0 0;">${note.status}</p>`;
  html += `</div>`;
  return html;
}

/**
 * Appends the lab standard note to a Doc report.
 */
function printLabStandardNoteDoc(body, rows, subjectName, isParent) {
  const note = buildLabStandardNote(rows, subjectName, isParent);
  if (!note) return;

  body.appendParagraph(`\n${note.heading}`).setBold(true).setFontSize(11).setForegroundColor(note.accent);
  body.appendParagraph(note.body).setBold(false).setFontSize(10).setForegroundColor('#333333');
  if (note.status) {
    body.appendParagraph(note.status).setBold(false).setFontSize(10).setForegroundColor('#333333');
  }
}

/**
 * Renders the encouragement note as HTML, or "" when there is nothing to say.
 */
function generateHtmlEncouragement(rows, isParent) {
  const note = buildEncouragementNote(rows, isParent);
  if (!note) return "";

  let html = `<div style="background-color: ${note.background}; border-left: 4px solid ${note.accent}; padding: 15px; margin: 20px 0; border-radius: 0 4px 4px 0; font-size: 13px;">`;
  html += `<p style="margin-top: 0; margin-bottom: 8px; font-weight: bold; color: ${note.accent}; font-size: 14px;">${note.heading}</p>`;
  html += `<p style="margin: 0 0 10px 0;">${note.intro}</p>`;
  html += `<ul style="margin: 0 0 10px 0; padding-left: 20px;">`;
  note.steps.forEach(step => { html += `<li style="margin-bottom: 4px;">${step}</li>`; });
  html += `</ul>`;
  html += `<p style="margin: 0;">${note.closing}</p>`;
  html += `</div>`;
  return html;
}

/**
 * Appends the encouragement note to a Doc report.
 */
function printEncouragementDoc(body, rows, isParent) {
  const note = buildEncouragementNote(rows, isParent);
  if (!note) return;

  body.appendParagraph(`\n${note.heading}`).setBold(true).setFontSize(11).setForegroundColor(note.accent);
  body.appendParagraph(note.intro).setBold(false).setFontSize(10).setForegroundColor('#333333');
  note.steps.forEach(step => {
    body.appendListItem(step).setGlyphType(DocumentApp.GlyphType.BULLET).setFontSize(10).setForegroundColor('#333333');
  });
  body.appendParagraph(note.closing).setBold(false).setFontSize(10).setForegroundColor('#333333');
}

function generateHtmlSummaryStats(rows) {
  const stats = rows.filter(r => r.isSummaryStat);
  if (stats.length === 0) return "";

  let html = `<h4 style="margin-bottom: 5px; color: #444; border-bottom: 1px solid #ccc; padding-bottom: 3px; margin-top: 20px;">Participation Metrics</h4>`;
  html += `<table style="width: 100%; max-width: 400px; border-collapse: collapse; font-size: 12px; margin-bottom: 15px;">`;
  stats.forEach(item => {
    let colorStyle = "";
    const nameLower = item.name.toLowerCase();
    const valNum = parseFloat(String(item.value).replace('%', ''));

    if (!isNaN(valNum)) {
      if (nameLower.includes("incomplete") || nameLower.includes("missing")) {
        // Lower is better. 0 is best.
        if (valNum === 0) colorStyle = "color: #2E7D32; font-weight: bold;"; // Green
        else if (valNum > 0) colorStyle = "color: #c62828; font-weight: bold;"; // Red
      } else if (nameLower.includes("completion")) {
        // Higher is better.
        if (valNum === 100 || valNum === 1) colorStyle = "color: #2E7D32; font-weight: bold;";
        else if (valNum < 100) colorStyle = "color: #ef6c00; font-weight: bold;"; // Orange
      }
    }

    html += `<tr>
              <td style="padding: 5px; border: 1px solid #eee; background-color: #f9f9f9; width: 70%; font-weight: bold;">${item.name}</td>
              <td style="padding: 5px; border: 1px solid #eee; text-align: center; ${colorStyle}">${item.value}</td>
             </tr>`;
  });
  html += `</table>`;
  return html;
}

function generateHtmlTables(rows) {
  const groups = {};
  const order = [];
  rows.forEach(r => {
    if (r.isSummaryStat) return;
    const c = r.category || "General";
    if (!groups[c]) { groups[c] = []; order.push(c); }
    groups[c].push(r);
  });

  let html = "";
  order.forEach(cat => {
    html += `<h4 style="margin-bottom: 5px; color: #444; border-bottom: 1px solid #ccc; padding-bottom: 3px;">${cat}</h4>`;
    html += `<table style="width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 15px;">`;
    html += `<tr style="background-color: #EFEFEF;">
              <th style="text-align: left; padding: 5px; border: 1px solid #ccc; width: 75%;">Assignment</th>
              <th style="text-align: center; padding: 5px; border: 1px solid #ccc; width: 25%;">Score</th>
             </tr>`;
    groups[cat].forEach(item => {
      const bgStyle = item.bgColor !== '#ffffff' ? `background-color: ${item.bgColor};` : '';
      const fontStyle = `color: ${item.fontColor}; font-weight: bold;`;
      const rowBgStyle = item.rowBg !== '#ffffff' ? `background-color: ${item.rowBg};` : '';
      const rowFontStyle = `color: ${item.rowFont};`;
      html += `<tr>
                <td style="padding: 5px; border: 1px solid #ccc; ${bgStyle} ${fontStyle}">${item.name}</td>
                <td style="text-align: center; padding: 5px; border: 1px solid #ccc; ${rowBgStyle} ${rowFontStyle}">${item.value}</td>
               </tr>`;
    });
    html += `</table>`;
  });
  return html;
}

/**
 * Builds the HTML interface for student selection.
 */
function buildStudentSelectorHtml(sections, mode, hasParentEmails) {
  const template = HtmlService.createTemplate(`
    <style>
      body { font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; padding: 0; margin: 0; background: #fcfcfc; color: #3c4043; overflow: hidden; display: flex; flex-direction: column; height: 100vh; }
      
      /* Header Area */
      .header { padding: 20px; background: #fff; border-bottom: 1px solid #dadce0; }
      h3 { margin: 0; color: #202124; font-size: 18px; font-weight: 500; }
      .subtitle { color: #5f6368; font-size: 13px; margin-top: 5px; }

      /* Segmented Email Destination Control */
      .dest-section {
        margin-top: 15px;
        padding-top: 12px;
        border-top: 1px dashed #dadce0;
      }
      .dest-label-title {
        font-weight: 600;
        font-size: 12px;
        color: #5f6368;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .dest-container {
        display: flex;
        gap: 10px;
      }
      .dest-pill {
        flex: 1;
        border: 1px solid #dadce0;
        border-radius: 6px;
        padding: 8px 12px;
        text-align: center;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
        background: white;
        transition: all 0.2s ease;
        user-select: none;
      }
      .dest-pill:hover {
        background: #f8f9fa;
        border-color: #c0c1c4;
      }
      .dest-pill input[type="radio"] {
        display: none;
      }
      .dest-pill.active {
        background: #e8f0fe;
        border-color: #1a73e8;
        box-shadow: 0 1px 2px rgba(26, 115, 232, 0.15);
      }
      .dest-pill.active span {
        color: #1967d2;
        font-weight: 600;
      }

      /* Content Area */
      #content { flex: 1; overflow-y: auto; padding: 10px 20px; }
      
      /* Toolbar: section filter + bulk links */
      .toolbar {
        display: flex; justify-content: space-between; align-items: flex-start;
        gap: 12px; margin-bottom: 12px; flex-wrap: wrap;
      }
      .filter-chips { display: flex; flex-wrap: wrap; gap: 6px; flex: 1; }
      .chip {
        padding: 4px 12px; border-radius: 14px; border: 1px solid #dadce0; background: white;
        font-size: 12px; color: #5f6368; cursor: pointer; user-select: none; white-space: nowrap;
      }
      .chip:hover { background: #f1f3f4; }
      .chip.active { background: #e8f0fe; border-color: #1a73e8; color: #1967d2; font-weight: 600; }
      .bulk-links { white-space: nowrap; padding-top: 4px; }
      .dest-note { font-size: 12px; color: #5f6368; font-style: italic; }

      /* Sections & Rows */
      .section-card {
        margin-bottom: 18px; border: 1px solid #dadce0; border-radius: 8px;
        background: white; overflow: hidden;
      }
      .section-header {
        padding: 10px 12px; background: #f8f9fa; border-bottom: 1px solid #e8eaed;
        color: #1967d2; font-weight: 600; font-size: 14px;
        display: flex; align-items: center; gap: 4px;
      }
      .section-name { flex: 1; cursor: pointer; font-weight: 600; }
      .section-count { font-size: 11px; color: #5f6368; font-weight: 400; margin-right: 4px; }
      .caret { cursor: pointer; color: #5f6368; font-size: 12px; padding: 0 4px; user-select: none; }
      .section-card.collapsed .caret { transform: rotate(-90deg); display: inline-block; }
      .section-body { padding: 8px; }
      .student-row {
        display: flex; align-items: center; padding: 10px 12px; margin-bottom: 4px;
        background: white; border: 1px solid #dadce0; border-radius: 6px; transition: background 0.1s;
      }
      .student-row:hover { background: #f1f3f4; border-color: #d2e3fc; }
      .none { color: #b0b0b0; }
      
      /* Controls */
      input[type="checkbox"] { transform: scale(1.1); margin-right: 12px; cursor: pointer; }
      label { flex: 1; cursor: pointer; font-size: 14px; display: flex; flex-direction: column; justify-content: center; }
      .email-sub { font-size: 11px; color: #70757a; margin-top: 3px; display: flex; gap: 15px; flex-wrap: wrap; }
      
      /* Badges */
      .badge { display: inline-block; padding: 2px 6px; border-radius: 12px; font-size: 10px; font-weight: bold; margin-left: 8px; }
      .badge-warn { background: #fce8e6; color: #c5221f; }

      /* Footer / Buttons */
      .footer { 
        padding: 15px 20px; background: #fff; border-top: 1px solid #dadce0; 
        display: flex; justify-content: space-between; align-items: center;
      }
      
      .btn { padding: 9px 20px; border-radius: 4px; font-weight: 500; font-size: 14px; cursor: pointer; border: none; }
      .btn-primary { background: #1a73e8; color: white; box-shadow: 0 1px 2px rgba(0,0,0,0.1); }
      .btn-primary:hover { background: #1765cc; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
      .btn-secondary { background: white; color: #5f6368; border: 1px solid #dadce0; margin-right: 10px; }
      .btn-secondary:hover { background: #f8f9fa; color: #202124; }
      
      /* Links */
      .action-link { color: #1a73e8; text-decoration: none; font-size: 12px; margin-right: 15px; cursor: pointer; }
      .action-link:hover { text-decoration: underline; }

      /* Loading Overlay */
      #loading { display: none; position: absolute; top:0; left:0; right:0; bottom:0; background: rgba(255,255,255,0.9); z-index: 10; display:flex; flex-direction: column; align-items: center; justify-content: center; }
      .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #1a73e8; border-radius: 50%; width: 30px; height: 30px; animation: spin 1s linear infinite; margin-bottom: 15px; }
      @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
    
    <!-- LOADER -->
    <div id="loading" style="display:none;">
       <div class="spinner"></div>
       <div style="font-weight:500; color:#555;" id="loading-text">Processing...</div>
    </div>

    <!-- HEADER -->
    <div class="header">
      <h3>${mode === 'email' ? '📧 Email Student Reports' : '📂 Generate Drive Reports'}</h3>
      <div class="subtitle">Select students below to generate their progress reports.</div>
      
      <? if (mode === 'email' && !hasParentEmails) { ?>
        <div class="dest-section">
          <div class="dest-note">✉️ No parent or guardian emails in this sheet &mdash; reports go to students only.</div>
        </div>
      <? } ?>

      <? if (!hasParentEmails) { ?>
        <input type="radio" name="email_dest" value="student" checked hidden>
      <? } ?>

      <? if (mode === 'email' && hasParentEmails) { ?>
        <div class="dest-section">
          <div class="dest-label-title">✉️ Send Emails To:</div>
          <div class="dest-container">
            <label class="dest-pill" id="pill_student">
              <input type="radio" name="email_dest" value="student" onchange="updateDestPills(this)">
              <span>👤 Student Only</span>
            </label>
            <label class="dest-pill" id="pill_parent">
              <input type="radio" name="email_dest" value="parent" onchange="updateDestPills(this)">
              <span>👥 Parent Only</span>
            </label>
            <label class="dest-pill active" id="pill_both">
              <input type="radio" name="email_dest" value="both" checked onchange="updateDestPills(this)">
              <span>✉️ Both</span>
            </label>
          </div>
        </div>
      <? } ?>
    </div>

    <!-- CONTENT -->
    <div id="content">
       <div class="toolbar">
         <div class="filter-chips">
           <span class="chip active" onclick="filterSection(this, 'all')">All sections</span>
           <? for (var g = 0; g < sections.length; g++) { ?>
             <span class="chip" onclick="filterSection(this, '<?= sections[g].id ?>')"><?= sections[g].name ?></span>
           <? } ?>
         </div>
         <div class="bulk-links">
           <a class="action-link" onclick="toggleAll(true)">Select All</a>
           <a class="action-link" onclick="toggleAll(false)">Select None</a>
         </div>
       </div>

       <? for (var g = 0; g < sections.length; g++) { ?>
         <? var sec = sections[g]; ?>
         <div class="section-card" id="card_<?= sec.id ?>">
           <div class="section-header">
             <input type="checkbox" class="sec-master" id="sec_chk_<?= sec.id ?>"
                    onchange="toggleSection(this, '<?= sec.id ?>')" checked>
             <label class="section-name" for="sec_chk_<?= sec.id ?>"><?= sec.name ?></label>
             <span class="section-count"><?= sec.students.length ?> student<?= sec.students.length === 1 ? '' : 's' ?></span>
             <span class="caret" onclick="toggleCollapse('<?= sec.id ?>')">▾</span>
           </div>
           <div class="section-body" id="body_<?= sec.id ?>">
             <? for (var i = 0; i < sec.students.length; i++) { ?>
               <? var stu = sec.students[i]; var uid = sec.id + '_' + i; ?>
               <div class="student-row">
                 <input type="checkbox" id="chk_<?= uid ?>" class="stu-chk <?= sec.id ?>"
                        data-mismatch="<?= stu.isMismatch ?>"
                        data-name="<?= stu.name ?>"
                        data-section="<?= sec.id ?>"
                        value="<?= stu.row ?>" onchange="updateStatus()" checked>
                 <label for="chk_<?= uid ?>">
                   <div><?= stu.name ?>
                     <? if (stu.isMismatch) { ?> <span class="badge badge-warn">Email Mismatch</span> <? } ?>
                   </div>
                   <div class="email-sub">
                     <span>👤 Student: <? if (stu.email) { ?><?= stu.email ?><? } else { ?><i class="none">None</i><? } ?></span>
                     <? if (hasParentEmails) { ?>
                       <span>👥 Parent: <? if (stu.parentEmail) { ?><?= stu.parentEmail ?><? } else { ?><i class="none">None</i><? } ?></span>
                     <? } ?>
                   </div>
                 </label>
               </div>
             <? } ?>
           </div>
         </div>
       <? } ?>
    </div>

    <!-- FOOTER -->
    <div class="footer">
       <span id="status-text" style="font-size:12px; color:#5f6368;">Ready</span>
       <div>
         <button class="btn btn-secondary" onclick="process('preview')">Preview</button>
         <button class="btn btn-primary" onclick="process('${mode}')">
           ${mode === 'email' ? 'Send Emails' : 'Generate Docs'}
         </button>
       </div>
    </div>

    <script>
      function updateDestPills(radio) {
        document.querySelectorAll('.dest-pill').forEach(pill => pill.classList.remove('active'));
        if (radio.checked) {
          radio.closest('.dest-pill').classList.add('active');
        }
      }

      function visibleCards() {
        return Array.from(document.querySelectorAll('.section-card'))
                    .filter(card => card.style.display !== 'none');
      }

      // Select All / None applies to the sections currently shown by the filter.
      function toggleAll(state) {
        visibleCards().forEach(card => {
          card.querySelectorAll('input[type="checkbox"]').forEach(c => { c.checked = state; });
        });
        updateStatus();
      }

      function toggleSection(source, secId) {
        document.querySelectorAll('.' + secId).forEach(c => { c.checked = source.checked; });
        updateStatus();
      }

      function toggleCollapse(secId) {
        const card = document.getElementById('card_' + secId);
        const collapsed = card.classList.toggle('collapsed');
        document.getElementById('body_' + secId).style.display = collapsed ? 'none' : '';
      }

      // View-only filter: hidden sections keep their selections, and the footer
      // count always reports the full selection so nothing is sent by surprise.
      function filterSection(chip, secId) {
        document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
        chip.classList.add('active');
        document.querySelectorAll('.section-card').forEach(card => {
          card.style.display = (secId === 'all' || card.id === 'card_' + secId) ? '' : 'none';
        });
        updateStatus();
      }

      function updateStatus() {
        const checked = Array.from(document.querySelectorAll('.stu-chk:checked'));
        const secs = new Set(checked.map(c => c.getAttribute('data-section')));

        document.querySelectorAll('.section-card').forEach(card => {
          const total = card.querySelectorAll('.stu-chk').length;
          const on = card.querySelectorAll('.stu-chk:checked').length;
          const master = card.querySelector('.sec-master');
          master.checked = on > 0;
          master.indeterminate = on > 0 && on < total;
        });

        const plural = (n, word) => n + ' ' + word + (n === 1 ? '' : 's');
        document.getElementById('status-text').innerText = checked.length === 0
          ? 'No students selected'
          : plural(checked.length, 'student') + ' selected in ' + plural(secs.size, 'section');
      }

      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', updateStatus);
      } else {
        updateStatus();
      }

      function process(action) {
        const checkboxes = document.querySelectorAll('.stu-chk:checked');
        const selected = Array.from(checkboxes).map(c => parseInt(c.value));
        
        if (selected.length === 0) return alert("Please select at least one student.");
        
        // Safety Check for Email
        if (action === 'email') {
           const mismatches = Array.from(checkboxes)
             .filter(c => c.getAttribute('data-mismatch') === 'true')
             .map(c => c.getAttribute('data-name'));
             
           if (mismatches.length > 0) {
              const msg = "⚠️ Warning: " + mismatches.length + " students have email addresses that don't match their names.\\n\\nExample: " + mismatches[0] + "\\n\\nContinue?";
              if (!confirm(msg)) return;
           }
        }

        // Get selected email destination option if in email mode or preview
        let emailDest = 'both';
        const selectedDest = document.querySelector('input[name="email_dest"]:checked');
        if (selectedDest) {
          emailDest = selectedDest.value;
        }

        // UI Updates
        document.getElementById('loading').style.display = 'flex';
        document.getElementById('loading-text').innerText = (action === 'preview') ? 'Generating Preview...' : 'Processing...';

        google.script.run
          .withSuccessHandler((res) => {
             document.getElementById('loading').style.display = 'none';
             if (action === 'preview') showPreview(res);
             else google.script.host.close();
          })
          .withFailureHandler((err) => {
             document.getElementById('loading').style.display = 'none';
             alert('Error: ' + err.message);
          })
          .runReportBatch(action, selected, emailDest);
      }

      function showPreview(htmlContent) {
        // Simple Modal for Preview
        const win = window.open("", "Preview", "width=600,height=600");
        win.document.write(htmlContent);
      }
    </script>
  `);

  template.sections = sections;
  template.mode = mode;
  template.hasParentEmails = hasParentEmails;
  return template.evaluate();
}

/**
 * Shows the tutorial sidebar.
 */
function showTutorialSidebar() {
  const html = HtmlService.createHtmlOutput(buildTutorialHtml())
    .setTitle('Gradebook Tools Guide')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function buildTutorialHtml() {
  return `
    <style>
      body { font-family: 'Segoe UI', Roboto, sans-serif; font-size: 14px; padding: 15px; color: #333; line-height: 1.5; }
      h3 { margin-top: 20px; color: #1a73e8; display: flex; align-items: center; gap: 8px; border-bottom: 1px solid #eee; padding-bottom: 5px; }
      h3:first-of-type { margin-top: 0; }
      .card { background: #f8f9fa; border: 1px solid #ddd; border-radius: 8px; padding: 15px; margin-bottom: 15px; }
      .step { display: flex; gap: 10px; margin-bottom: 8px; align-items: flex-start; }
      .num { background: #1a73e8; color: white; border-radius: 50%; width: 20px; height: 20px; display: flex; align-items: center; justify-content: center; font-size: 12px; flex-shrink: 0; margin-top: 2px; }
      button { background: #1a73e8; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; width: 100%; font-weight: 500; margin-top: 5px; }
      button.secondary { background: white; border: 1px solid #dadce0; color: #1a73e8; }
      button:hover { opacity: 0.9; }
      .section-title { font-weight: 700; color: #5f6368; margin-bottom: 5px; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 15px; }
      .menu-item { margin-bottom: 8px; }
      .menu-name { font-weight: 600; color: #202124; }
      .menu-desc { font-size: 13px; color: #5f6368; margin-top: 2px; }
    </style>
    
    <h3>👋 Gradebook Guide</h3>
    <p>Generate individual progress reports for students or Google Drive archives.</p>

    <div class="card">
        <div style="font-weight:bold; margin-bottom:10px;">🚀 Quick Start</div>
        <div class="step"><div class="num">1</div><div><b>Prepare Data</b>: Ensure your sheet has "Name" and "Email" columns.</div></div>
        <div class="step"><div class="num">2</div><div><b>Select Tool</b>: Choose Email or Drive reports from the menu.</div></div>
        <div class="step"><div class="num">3</div><div><b>Run</b>: Select students and click "Go".</div></div>
        
        <button class="secondary" onclick="google.script.run.generateGradebookTemplate()">📘 Create Demo Sheet</button>
    </div>

    <h3>📖 Menu Reference</h3>
    
    <div class="section-title">Generation Tools</div>
    <div class="menu-item">
        <div class="menu-name">📧 Email Reports (Selector)</div>
        <div class="menu-desc">Opens the student selector logic. Sends individual emails to students with their grades + standard mastery (if applicable). Includes a "Preview" mode.</div>
    </div>
    <div class="menu-item">
        <div class="menu-name">📂 Generate Reports (Drive)</div>
        <div class="menu-desc">Creates a single Google Doc containing reports for all selected students, separated by page breaks. Useful for printing or archiving.</div>
    </div>

    <div class="section-title">Setup</div>
    <div class="menu-item">
        <div class="menu-name">📘 Generate Demo Gradebook</div>
        <div class="menu-desc">Creates a sample sheet with properly formatted headers and dummy data so you can test the script immediately.</div>
    </div>

    <div style="margin-top:20px; font-size:12px; color:#666; text-align:center; border-top: 1px solid #eee; padding-top: 15px;">
        <p style="margin-bottom:5px;">Developed by <a href="https://knuffke.com/support" target="_blank" style="color:#333; text-decoration:none;"><b>David Knuffke</b></a></p>
        <p style="font-size:10px; margin-top:5px;">Made available under a <a href="http://creativecommons.org/licenses/by-nc-sa/4.0/" target="_blank">CC BY-NC-SA 4.0 License</a>.</p>
        <a href="#" onclick="google.script.host.close()">Close Guide</a>
    </div>
  `;
}

/**
 * Opens the Gradebook Setup Checker & Guide Dialog.
 */
function showSetupGuide() {
  const html = HtmlService.createHtmlOutput(buildSetupGuideHtml())
    .setWidth(650)
    .setHeight(680)
    .setTitle('Gradebook Setup Checker & Guide');
  SpreadsheetApp.getUi().showModalDialog(html, '🛠️ Gradebook Setup Checker & Guide');
}

/**
 * Runs structural diagnostics on the active sheet and returns analysis.
 */
function runSetupVerification() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const sheetName = sheet.getName();
    
    const results = {
      sheetName: sheetName,
      nameColFound: false,
      nameColLetter: "",
      emailColFound: false,
      emailColLetter: "",
      parentColFound: false,
      parentColLetter: "",
      studentCount: 0,
      sectionCount: 0,
      sectionNames: [],
      headerRowNumber: 0,
      firstDataRowNumber: 0,
      warnings: [],
      successes: []
    };

    if (sheet.getLastRow() < 2) {
      results.warnings.push("The active sheet is empty or has fewer than 2 rows. It must contain at least 4 rows to represent a proper Gradebook.");
      return results;
    }

    // Use the same scan the Student Selector uses, so the two can never disagree.
    const roster = scanGradebookRoster(sheet);
    const cols = roster.cols;
    const colLetter = (idx) => {
      let label = "";
      for (let n = idx; n >= 0; n = Math.floor(n / 26) - 1) {
        label = String.fromCharCode(65 + (n % 26)) + label;
      }
      return label;
    };

    const headerRowNumber = roster.headerRow + 1;      // 1-indexed for humans
    const firstDataRowNumber = roster.firstDataRow + 1;
    results.headerRowNumber = headerRowNumber;
    results.firstDataRowNumber = firstDataRowNumber;
    results.successes.push(`Roster labels found in Row ${headerRowNumber}; student data read from Row ${firstDataRowNumber} down.`);

    // 1. Check Name Column
    if (!cols.nameFallback) {
      results.nameColFound = true;
      results.nameColLetter = colLetter(cols.name);
      if (cols.name === 1) {
        results.successes.push("Column B correctly designated as 'Name'.");
      } else {
        results.warnings.push(`'Name' column found in Column ${results.nameColLetter} instead of Column B. Keeping student names in Column B is highly recommended.`);
      }
    } else {
      results.warnings.push(`No column containing 'Name' was found in Row ${headerRowNumber}. Assuming Column B. Add a 'Name' header to be sure students are read correctly.`);
    }

    // 2. Student & Parent Email Columns
    if (cols.email !== -1) {
      results.emailColFound = true;
      results.emailColLetter = colLetter(cols.email);
      results.successes.push(`Student 'Email' column found in Column ${results.emailColLetter}.`);
    } else {
      results.warnings.push(`No student 'Email' column was detected in Row ${headerRowNumber}. Add an 'Email' header above your email addresses to allow sending reports to students.`);
    }

    if (cols.parentEmail !== -1) {
      results.parentColFound = true;
      results.parentColLetter = colLetter(cols.parentEmail);
      results.successes.push(`Parent/Guardian 'Parent Email' column found in Column ${results.parentColLetter}.`);
    } else {
      results.warnings.push(`No 'Parent Email' column was found. If you wish to send copies to parents, add a 'Parent Email' or 'Guardian Email' header in Row ${headerRowNumber}.`);
    }

    // 3. Student Rows & Class Sections
    const namedSections = roster.sections.filter(s => s !== "Ungrouped");
    results.studentCount = roster.students.length;
    results.sectionCount = namedSections.length;
    results.sectionNames = namedSections;

    if (results.studentCount > 0) {
      results.successes.push(`Parsed ${results.studentCount} active student rows.`);
    } else {
      results.warnings.push(`No active students detected below Row ${headerRowNumber}. Ensure student names sit under a 'Name' header and that each student has an email or some graded work.`);
    }

    if (results.sectionCount > 0) {
      results.successes.push(`Detected ${results.sectionCount} class sections: ${namedSections.join(', ')}.`);
    } else {
      results.warnings.push("No class sections detected. To group students in the selector, either put the section name (e.g. 'Block 1') in Column A of each student row, or give each class its own heading row with the section name and no grades.");
    }

    return results;
  } catch (e) {
    throw new Error("Verification Error: " + e.message);
  }
}

/**
 * Builds the HTML content for the Setup Checker & Guide Dialog.
 */
function buildSetupGuideHtml() {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        body {
          font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
          margin: 0;
          padding: 0;
          background-color: #f8f9fa;
          color: #3c4043;
          font-size: 14px;
          line-height: 1.5;
        }
        
        .container {
          display: flex;
          flex-direction: column;
          height: 100vh;
          box-sizing: border-box;
        }

        /* Navigation Tabs */
        .tabs {
          display: flex;
          background-color: #ffffff;
          border-bottom: 1px solid #dadce0;
          padding: 10px 20px 0 20px;
          flex-shrink: 0;
        }
        .tab-btn {
          padding: 12px 24px;
          cursor: pointer;
          font-weight: 500;
          font-size: 14px;
          color: #5f6368;
          background: none;
          border: none;
          border-bottom: 3px solid transparent;
          outline: none;
          transition: all 0.2s ease;
          display: flex;
          align-items: center;
          gap: 8px;
        }
        .tab-btn:hover {
          color: #1a73e8;
          background-color: #f8f9fa;
          border-radius: 4px 4px 0 0;
        }
        .tab-btn.active {
          color: #1a73e8;
          border-bottom-color: #1a73e8;
          font-weight: 600;
        }

        /* Panel Content */
        .tab-content {
          flex: 1;
          overflow-y: auto;
          padding: 20px 24px;
          box-sizing: border-box;
        }
        .panel {
          display: none;
        }
        .panel.active {
          display: block;
        }

        h2 {
          margin-top: 0;
          font-size: 18px;
          font-weight: 500;
          color: #202124;
          display: flex;
          align-items: center;
          gap: 8px;
        }
        p {
          margin-top: 0;
          margin-bottom: 15px;
          color: #5f6368;
          font-size: 13.5px;
        }

        /* Grid Table Mockup styling */
        .mockup-card {
          background: white;
          border: 1px solid #dadce0;
          border-radius: 8px;
          padding: 16px;
          margin-bottom: 20px;
          box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }
        .mockup-title {
          font-weight: 600;
          color: #202124;
          font-size: 13px;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          margin-bottom: 12px;
          display: flex;
          align-items: center;
          gap: 6px;
        }
        .grid-mockup {
          width: 100%;
          border-collapse: collapse;
          font-size: 11px;
          font-family: monospace;
          margin-bottom: 10px;
          border: 1px solid #dadce0;
        }
        .grid-mockup th, .grid-mockup td {
          border: 1px solid #dadce0;
          padding: 6px 8px;
          text-align: left;
        }
        .grid-mockup tr.header-cat {
          background-color: #e0e0e0;
          font-weight: bold;
        }
        .grid-mockup tr.header-main {
          background-color: #434343;
          color: white;
          font-weight: bold;
        }
        .grid-mockup tr.header-std {
          background-color: #f3f3f3;
          font-style: italic;
        }
        .grid-mockup tr.section-row {
          background-color: #000000;
          color: white;
          font-weight: bold;
        }
        .grid-mockup tr.student-row {
          background-color: #ffffff;
        }

        .highlight-col {
          border: 1.5px solid #1a73e8 !important;
          background-color: #e8f0fe;
        }

        /* Bullet instruction styling */
        .instruction-list {
          padding-left: 20px;
          margin-bottom: 20px;
        }
        .instruction-list li {
          margin-bottom: 10px;
          color: #3c4043;
        }
        .instruction-list strong {
          color: #202124;
        }

        /* Live Checker elements */
        .status-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          background: #e8f0fe;
          border-radius: 8px;
          padding: 12px 16px;
          margin-bottom: 20px;
          border: 1px solid #d2e3fc;
        }
        .status-title {
          font-weight: 600;
          color: #1967d2;
          display: flex;
          align-items: center;
          gap: 8px;
        }
        .status-btn {
          background: #1a73e8;
          color: white;
          border: none;
          padding: 6px 14px;
          border-radius: 4px;
          font-size: 12px;
          font-weight: 500;
          cursor: pointer;
          transition: background 0.2s;
        }
        .status-btn:hover {
          background: #1765cc;
        }

        .check-item {
          display: flex;
          padding: 12px 16px;
          background: white;
          border: 1px solid #dadce0;
          border-radius: 8px;
          margin-bottom: 12px;
          gap: 16px;
          align-items: flex-start;
          transition: box-shadow 0.2s ease;
        }
        .check-item:hover {
          box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        .check-icon {
          font-size: 20px;
          flex-shrink: 0;
          margin-top: -2px;
        }
        .check-details {
          flex: 1;
        }
        .check-title {
          font-weight: 600;
          color: #202124;
          margin-bottom: 4px;
        }
        .check-desc {
          font-size: 12.5px;
          color: #5f6368;
        }

        .check-success { border-left: 4px solid #34a853; }
        .check-success .check-icon { color: #34a853; }
        .check-warning { border-left: 4px solid #f9ab00; }
        .check-warning .check-icon { color: #f9ab00; }

        /* Loader Overlay */
        #checker-loading {
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          padding: 40px 0;
        }
        .spinner {
          border: 3px solid #f3f3f3;
          border-top: 3px solid #1a73e8;
          border-radius: 50%;
          width: 32px;
          height: 32px;
          animation: spin 1s linear infinite;
          margin-bottom: 16px;
        }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }

        /* Footer buttons */
        .footer {
          padding: 15px 24px;
          background-color: #ffffff;
          border-top: 1px solid #dadce0;
          display: flex;
          justify-content: flex-end;
          gap: 12px;
          flex-shrink: 0;
        }
        .btn {
          padding: 8px 18px;
          border-radius: 4px;
          font-size: 13.5px;
          font-weight: 500;
          cursor: pointer;
          border: 1px solid transparent;
        }
        .btn-secondary {
          background: white;
          color: #5f6368;
          border-color: #dadce0;
        }
        .btn-secondary:hover {
          background: #f8f9fa;
          color: #202124;
        }
        .btn-primary {
          background: #1a73e8;
          color: white;
        }
        .btn-primary:hover {
          background: #1765cc;
        }
      </style>
    </head>
    <body>
      <div class="container">
        
        <!-- Navigation Tabs -->
        <div class="tabs">
          <button class="tab-btn active" onclick="switchTab(event, 'tutorial-panel')">
            📖 1. Sheet Setup Tutorial
          </button>
          <button class="tab-btn" onclick="switchTab(event, 'checker-panel')">
            🔍 2. Live Setup Checker
          </button>
        </div>

        <!-- 1. Setup Tutorial Panel -->
        <div id="tutorial-panel" class="tab-content panel active">
          <h2>How to Set Up Your Gradebook Sheet</h2>
          <p>The Gradebook Reporter scripts look for specific cells, formatting, and structures to generate individual progress reports successfully. Set up your active sheet using the layout below:</p>
          
          <div class="mockup-card">
            <div class="mockup-title">📊 Gradebook Spreadsheet Structure Mockup</div>
            <table class="grid-mockup">
              <thead>
                <tr class="header-cat">
                  <td>Row 1 (Categories)</td>
                  <td></td>
                  <td></td>
                  <td></td>
                  <td>Classwork</td>
                  <td>Classwork</td>
                  <td>Homework</td>
                  <td>Assessments</td>
                </tr>
                <tr class="header-main">
                  <td>Row 2 (Headers)</td>
                  <td>Section</td>
                  <td class="highlight-col">Name</td>
                  <td class="highlight-col">Email</td>
                  <td>Parent Email</td>
                  <td>Assignment 1</td>
                  <td>Assignment 2</td>
                  <td>Exam 1</td>
                </tr>
                <tr class="header-std">
                  <td>Row 3 (Standards)</td>
                  <td></td>
                  <td></td>
                  <td></td>
                  <td></td>
                  <td>Standard 1</td>
                  <td>Standard 1</td>
                  <td>Standard 2</td>
                </tr>
              </thead>
              <tbody>
                <tr class="section-row">
                  <td>Row 4 (Divider)</td>
                  <td colspan="7">Block 1</td>
                </tr>
                <tr class="student-row">
                  <td>Row 5 (Student)</td>
                  <td>Block 1</td>
                  <td class="highlight-col">Potter, Harry</td>
                  <td class="highlight-col">harry@hogwarts.edu</td>
                  <td>james.potter@hogwarts.edu</td>
                  <td>1</td>
                  <td>1</td>
                  <td>95</td>
                </tr>
                <tr class="student-row">
                  <td>Row 6 (Student)</td>
                  <td>Block 1</td>
                  <td class="highlight-col">Granger, Hermione</td>
                  <td class="highlight-col">hermione@hogwarts.edu</td>
                  <td>mr.granger@dentist.com</td>
                  <td>1</td>
                  <td>1</td>
                  <td>100</td>
                </tr>
              </tbody>
            </table>
            <div style="font-size: 11px; color:#70757a; text-align: center;">
              * Highlighted columns (<b>Name</b> and <b>Email</b>) are strictly required for reports to work.
            </div>
          </div>

          <ul class="instruction-list">
            <li>
              <strong>Row 1: Categories Row</strong> — Categorizes assignment columns. Merged cells or filled cells categorizing headers below them (e.g. <i>Classwork, Homework, Assessments</i>).
            </li>
            <li>
              <strong>Row 2: Header Labels Row</strong> — Must contain exact column header names:
              <ul>
                <li>Column A: <strong>Section</strong> (class/block designation).</li>
                <li>Column B: <strong>Name</strong> (entered as <i>"LastName, FirstName"</i> or <i>"FirstName LastName"</i>).</li>
                <li>Any Column: <strong>Email</strong> (student email column).</li>
                <li>Any Column: <strong>Parent Email</strong> (optional, adjacent column containing parent/guardian emails).</li>
                <li>Columns E+: Individual assignment names.</li>
              </ul>
            </li>
            <li>
              <strong>Row 3: Standards / Targets Row</strong> — Used optionally. Used to map standards/learning targets to assignments. If not using standards, keep this row empty but do NOT delete it.
            </li>
            <li>
              <strong>Row 4+: Student Data & Dividers</strong>
              <ul>
                <li><strong>Class Sections</strong>: Students are grouped two ways, and either works. Put the section name (e.g. <i>Block 1</i>) in Column A of every student row, <em>or</em> give each class a heading row that holds only the section name &mdash; no email, no grades. Styling the heading row with a solid background is optional; the script goes by content, not colour.</li>
                <li><strong>Student Rows</strong>: Student details and grade records. A row counts as a student only if it has a name plus an email, a parent email, some graded work, or a <i>Last, First</i> style name &mdash; so course titles and banner rows are treated as headings instead of students. Formula cells and summary averages are automatically skipped.</li>
              </ul>
            </li>
          </ul>
        </div>

        <!-- 2. Live Setup Checker Panel -->
        <div id="checker-panel" class="tab-content panel">
          <h2>Active Sheet Health Check</h2>
          <p>Scan your current active sheet to verify that all structural parts are correctly aligned and formatted.</p>
          
          <div class="status-header">
            <div class="status-title">
              📋 Active Sheet: <span id="sheet-name-label" style="font-weight: bold; color: #202124;">Loading...</span>
            </div>
            <button class="status-btn" onclick="runCheck()">🔄 Refresh Diagnostics</button>
          </div>

          <!-- Loading Indicator -->
          <div id="checker-loading">
            <div class="spinner"></div>
            <div style="color: #5f6368; font-size: 13px;">Analyzing spreadsheet layouts...</div>
          </div>

          <!-- Check Results -->
          <div id="checker-results" style="display:none;"></div>
        </div>

        <!-- Dialog Footer -->
        <div class="footer">
          <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>
          <button class="btn btn-primary" onclick="generateDemo()">📘 Create Demo Sheet</button>
        </div>

      </div>

      <script>
        // Switch between tabs
        function switchTab(evt, panelId) {
          document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
          document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
          
          evt.currentTarget.classList.add('active');
          document.getElementById(panelId).classList.add('active');
          
          if (panelId === 'checker-panel') {
            runCheck();
          }
        }

        // Run structural checks
        function runCheck() {
          document.getElementById('checker-loading').style.display = 'flex';
          document.getElementById('checker-results').style.display = 'none';

          google.script.run
            .withSuccessHandler((res) => {
              document.getElementById('checker-loading').style.display = 'none';
              renderCheckerResults(res);
            })
            .withFailureHandler((err) => {
              document.getElementById('checker-loading').style.display = 'none';
              const resultsDiv = document.getElementById('checker-results');
              resultsDiv.innerHTML = '<div class="check-item check-warning"><span class="check-icon">⚠️</span><div class="check-details"><div class="check-title">Analysis Failed</div><div class="check-desc">' + err.message + '</div></div></div>';
              resultsDiv.style.display = 'block';
            })
            .runSetupVerification();
        }

        // Render diagnostics checklist
        function renderCheckerResults(res) {
          document.getElementById('sheet-name-label').innerText = res.sheetName || 'Unknown Sheet';
          const resultsDiv = document.getElementById('checker-results');
          resultsDiv.innerHTML = '';

          let html = '';

          // 0. Layout Interpretation - shows how the sheet was actually read, so
          //    an unfamiliar layout is visible instead of silently misparsed.
          html += '<div class="check-item check-success">' +
                  '<span class="check-icon">🧭</span>' +
                  '<div class="check-details">' +
                    '<div class="check-title">Layout Detected</div>' +
                    '<div class="check-desc">Roster headers read from Row <b>' + res.headerRowNumber + '</b>; ' +
                      'students read from Row <b>' + res.firstDataRowNumber + '</b> down. ' +
                      'If those row numbers look wrong, everything below will be wrong too.</div>' +
                  '</div>' +
                '</div>';

          // 1. Name Column Check
          if (res.nameColFound) {
            html += '<div class="check-item check-success">' +
                    '<span class="check-icon">✅</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">Student Name Column</div>' +
                      '<div class="check-desc">Found in Column <b>' + res.nameColLetter + '</b>. Student names are correctly located for reporting.</div>' +
                    '</div>' +
                  '</div>';
          } else {
            html += '<div class="check-item check-warning">' +
                    '<span class="check-icon">⚠️</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">Student Name Column Missing</div>' +
                      '<div class="check-desc">We couldn\'t find a column labeled "Name" in Row 2. You need to designate Column B as "Name" for the script to function.</div>' +
                    '</div>' +
                  '</div>';
          }

          // 2. Student Email Check
          if (res.emailColFound) {
            html += '<div class="check-item check-success">' +
                    '<span class="check-icon">✅</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">Student Email Column</div>' +
                      '<div class="check-desc">Found in Column <b>' + res.emailColLetter + '</b>. Student progress reports can be dispatched.</div>' +
                    '</div>' +
                  '</div>';
          } else {
            html += '<div class="check-item check-warning">' +
                    '<span class="check-icon">⚠️</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">Student Email Column Missing</div>' +
                      '<div class="check-desc">No column labeled "Email" was detected in Row 2. Progress reports cannot be sent via email without a student email column.</div>' +
                    '</div>' +
                  '</div>';
          }

          // 3. Parent Email Check
          if (res.parentColFound) {
            html += '<div class="check-item check-success">' +
                    '<span class="check-icon">✅</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">Parent/Guardian Email Column</div>' +
                      '<div class="check-desc">Found in Column <b>' + res.parentColLetter + '</b>. You can now elect to copy reports to parents/guardians.</div>' +
                    '</div>' +
                  '</div>';
          } else {
            html += '<div class="check-item check-warning">' +
                    '<span class="check-icon">ℹ️</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">Parent Email Column (Optional)</div>' +
                      '<div class="check-desc">No column labeled "Parent Email" or "Guardian Email" was found. While standard reports will work, adding one allows you to send copies to parents!</div>' +
                    '</div>' +
                  '</div>';
          }

          // 4. Student Count Check
          if (res.studentCount > 0) {
            html += '<div class="check-item check-success">' +
                    '<span class="check-icon">✅</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">Active Student Rows</div>' +
                      '<div class="check-desc">Correctly loaded <b>' + res.studentCount + '</b> student rows. Calculations, averages, and empty headers are excluded.</div>' +
                    '</div>' +
                  '</div>';
          } else {
            html += '<div class="check-item check-warning">' +
                    '<span class="check-icon">⚠️</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">No Students Detected</div>' +
                      '<div class="check-desc">We couldn\'t find any student rows starting at Row 4. Verify your student names are written starting at row 4, column B.</div>' +
                    '</div>' +
                  '</div>';
          }

          // 5. Class Sections Check
          if (res.sectionCount > 0) {
            html += '<div class="check-item check-success">' +
                    '<span class="check-icon">✅</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">Class Sections</div>' +
                      '<div class="check-desc">Found <b>' + res.sectionCount + '</b> class sections: <b>' + res.sectionNames.join(', ') + '</b>. Each one gets its own group in the selector.</div>' +
                    '</div>' +
                  '</div>';
          } else {
            html += '<div class="check-item check-warning">' +
                    '<span class="check-icon">ℹ️</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">No Class Sections (Optional)</div>' +
                      '<div class="check-desc">All students will appear in one group. To split them by period or block, either put the section name in Column A of each student row, or give each class a heading row containing only the section name.</div>' +
                    '</div>' +
                  '</div>';
          }

          resultsDiv.innerHTML = html;
          resultsDiv.style.display = 'block';
        }

        // Trigger Demo Gradebook Generator
        function generateDemo() {
          google.script.run
            .withSuccessHandler(() => {
              alert("Demo Gradebook sheet successfully created! Switch to the 'Demo Gradebook' sheet tab to see a fully structured sample gradebook.");
              google.script.host.close();
            })
            .withFailureHandler((err) => {
              alert("Error generating demo sheet: " + err.message);
            })
            .generateGradebookTemplate();
        }
      </script>
    </body>
    </html>
  `;
}