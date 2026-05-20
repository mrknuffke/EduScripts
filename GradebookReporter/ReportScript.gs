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

/**
 * Scans the sheet and opens the Student Selector Dialog
 */
function showStudentSelector(mode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getDisplayValues();
  const backgrounds = sheet.getDataRange().getBackgrounds();
  const fontColors = sheet.getDataRange().getFontColors();

  const nameColIndex = 1;
  let emailColIndex = -1;
  let parentEmailColIndex = -1;

  // Find Email & Parent Email Columns in first 2 rows
  for (let r = 0; r < 2; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (!data[r][c]) continue;
      const cellText = data[r][c].toLowerCase();
      if (cellText.includes('parent') || cellText.includes('guardian')) {
        parentEmailColIndex = c;
      } else if (cellText.includes('email') && !cellText.includes('parent') && !cellText.includes('guardian')) {
        emailColIndex = c;
      }
    }
  }
  // Fallback for Student Email: if not found by explicit search but there's another column containing "email"
  if (emailColIndex === -1) {
    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < data[r].length; c++) {
        if (!data[r][c]) continue;
        const cellText = data[r][c].toLowerCase();
        if (cellText.includes('email') && c !== parentEmailColIndex) {
          emailColIndex = c;
          break;
        }
      }
      if (emailColIndex !== -1) break;
    }
  }

  // Collect Students
  let students = [];
  let currentSection = "Ungrouped";

  for (let r = 3; r < data.length; r++) {
    const row = data[r];
    const firstCellBg = backgrounds[r][0];
    const firstCellFont = fontColors[r][0];
    const colA = row[0] ? String(row[0]).trim() : "";
    const name = row[nameColIndex] ? String(row[nameColIndex]).trim() : "";

    // DETECT SECTION HEADER
    if (colA !== "" && (
      firstCellBg === '#000000' ||
      firstCellBg === '#434343' ||
      (firstCellBg !== '#ffffff' && firstCellFont === '#ffffff') ||
      colA.toLowerCase().includes('block') ||
      (name === "" && !colA.includes("Average")))) {
      currentSection = colA;
      continue;
    }

    // FILTER JUNK
    if (!name || name === '' || name === '0' ||
      name.includes('Name:') || name === 'Student' || name === 'Preferred Name' ||
      name.includes('Average') || name.includes('Sum') ||
      colA.includes('Average') || colA.includes('Sum')) {
      continue;
    }

    const email = (emailColIndex > -1 && row[emailColIndex]) ? row[emailColIndex] : "";
    const parentEmail = (parentEmailColIndex > -1 && row[parentEmailColIndex]) ? row[parentEmailColIndex] : "";

    // SAFETY CHECK
    let isMismatch = false;
    if (email.includes('@')) {
      let lastName = name.includes(',') ? name.split(',')[0].trim() : name.split(' ')[0].trim();
      const cleanName = lastName.replace(/[^a-zA-Z]/g, '').toLowerCase();
      const cleanEmail = email.split('@')[0].replace(/[^a-zA-Z]/g, '').toLowerCase();
      if (cleanName.length > 1 && !cleanEmail.includes(cleanName)) isMismatch = true;
    } else if (email !== "") {
      isMismatch = true;
    }

    students.push({
      row: r,
      name: name,
      email: email,
      parentEmail: parentEmail,
      section: currentSection,
      isMismatch: isMismatch
    });
  }

  if (students.length === 0) {
    SpreadsheetApp.getUi().alert("No students found. Check your Gradebook format.");
    return;
  }

  // --- CRITICAL: SORT STUDENTS BY SECTION ---
  students.sort((a, b) => {
    if (a.section === b.section) return a.name.localeCompare(b.name);
    return a.section.localeCompare(b.section);
  });

  // Assign IDs for UI
  let uniqueSections = [...new Set(students.map(s => s.section))];
  let sectionMap = {};
  uniqueSections.forEach((sec, idx) => { sectionMap[sec] = `sec-${idx}`; });
  students.forEach(s => { s.sectionId = sectionMap[s.section]; });

  // Generate and Show UI
  const html = buildStudentSelectorHtml(students, mode);
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
  const nameColIndex = 1;

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
  // Fill-right Categories for merged headers
  const categories = data[categoryRowIndex] ? [...data[categoryRowIndex]] : [];
  for (let i = 1; i < categories.length; i++) {
    if (categories[i] === "" && categories[i - 1] !== "") {
      categories[i] = categories[i - 1];
    }
  }
  let emailColIndex = -1;
  let parentEmailColIndex = -1;

  // Search headers
  for (let i = 0; i < headers.length; i++) {
    if (!headers[i]) continue;
    const text = headers[i].toLowerCase();
    if (text.includes('parent') || text.includes('guardian')) {
      parentEmailColIndex = i;
    } else if (text.includes('email') && !text.includes('parent') && !text.includes('guardian')) {
      emailColIndex = i;
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
  for (let i = nameColIndex + 1; i < headers.length; i++) {
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

  let lastCategory = "Uncategorized";
  columnDefs.forEach(col => {
    if (!col || col.id <= nameColIndex) return;
    if (col.rawCategory && col.rawCategory.trim() !== "") lastCategory = col.rawCategory.trim();
    col.finalCategory = lastCategory;
  });

  let lastSeenHeader = "";
  columnDefs.forEach(col => {
    if (!col || col.id <= nameColIndex) return;
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
    const assessmentKeywords = ['quiz', 'test', 'exam', 'assess', 'wa', 'webassign', 'unit', 'quest', 'lab'];
    const headerScoreKeywords = ['raw', 'score', 'percent', '%', 'letter', 'grade', 'points', 'pts'];
    const summaryKeywords = ['completion', 'missing', 'participation', 'rate'];

    const matchesAssessment = assessmentKeywords.some(k => lowerStandard.includes(k) || lowerCategory.includes(k));
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
      if (!col || idx <= nameColIndex || col.rawHeader === 'Assignment' || col.rawHeader === 'Preferred Name') return;

      const rawLower = col.rawHeader ? col.rawHeader.toLowerCase() : "";
      const finalLower = col.finalName ? col.finalName.toLowerCase() : "";
      if (rawLower.includes("excused") || rawLower.includes("i's and m's") || finalLower.includes("excused") || finalLower.includes("i's and m's") || idx === emailColIndex) return;
      if (!col.finalName && !col.standard) return;

      let value = row[idx];
      let displayValue = value;
      let isIssue = false;
      let isExempt = false;

      if (value) {
        const valStr = String(value).trim();
        const lowerVal = valStr.toLowerCase();
        const rawHeaderLower = col.rawHeader ? col.rawHeader.toLowerCase() : "";

        if ((lowerVal === 'true' || valStr === '1') && !col.isSummaryStat) displayValue = 'Complete';
        else if ((valStr === '0' || lowerVal === 'm' || lowerVal === 'false') && !col.isSummaryStat) { displayValue = 'Missing'; isIssue = true; }
        else if (lowerVal === 'ex') { displayValue = 'Exempt'; isExempt = true; }
        else if (lowerVal === 'i') { displayValue = 'Incomplete'; isIssue = true; }
        // Treat 0.5 or .5 as Incomplete for Activities (AC:) and InfoDocs (ID:)
        else if ((valStr === '0.5' || valStr === '.5') && !col.isSummaryStat) {
          const isActivityOrInfoDoc = rawHeaderLower.startsWith('ac:') || rawHeaderLower.startsWith('id:');
          if (isActivityOrInfoDoc) { displayValue = 'Incomplete'; isIssue = true; }
        }

        if (subjectName === "AP Biology") {
          const nameCheck = col.finalName.toLowerCase().trim();
          if (nameCheck.startsWith("lab")) {
            const numVal = parseFloat(valStr);
            if (!isNaN(numVal) && numVal < 4) isIssue = true;
          }
        }
      } else {
        displayValue = "-";
      }

      let shouldReport = false;
      if (isIssue || isExempt) shouldReport = true;
      if (subjectName === "Chemistry" && col.isSummativeStandard) shouldReport = true;
      if (col.isQuizOrWebAssign && displayValue !== "-") shouldReport = true;
      if (col.isSummaryStat) shouldReport = true;

      if (shouldReport) {
        reportRows.push({
          category: col.finalCategory,
          name: expandAssignmentPrefix(col.finalName),
          value: displayValue,
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

    const hasActualMissingWork = reportRows.some(item =>
      !item.isQuizOrWebAssign && !item.isSummaryStat && (
        item.value === 'Missing' ||
        item.value === 'Incomplete' || (subjectName === "AP Biology" && item.name.toLowerCase().startsWith("lab") && parseFloat(item.value) < 4)
      )
    );

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
      const studentEmail = (emailColIndex > -1) ? row[emailColIndex] : "No Student Email";
      const parentEmail = (parentEmailColIndex > -1) ? row[parentEmailColIndex] : "No Parent Email";
      
      const showStudent = (emailDest === 'student' || emailDest === 'both');
      const showParent = (emailDest === 'parent' || emailDest === 'both') && parentEmail !== "";

      if (showStudent) {
        const htmlBody = generateHtmlReport(titlePrefix, studentName, reportRows, isStudentInTrouble, subjectName, coolMessages, false);
        previewHtml += `<div class="preview-box" style="margin-bottom: 40px; border-bottom: 4px solid #ccc; padding-bottom: 40px; background-color: #fcfcfc; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                          <div class="preview-header" style="background-color: #1a73e8; color: white; padding: 8px 12px; font-weight: bold; border-radius: 4px 4px 0 0; margin-bottom: 15px;">STUDENT PREVIEW ${previewCount + 1}: ${studentName} (${studentEmail})</div>
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

  html += `<p style="font-size: 10px; color: #888; text-align: center; margin-top: 30px;">Generated by Gradebook Tools</p>`;
  html += `</div>`;
  return html;
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
function buildStudentSelectorHtml(students, mode) {
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
      
      /* Sections & Rows */
      .section-header { 
        margin-top: 20px; margin-bottom: 8px; padding-bottom: 5px;
        border-bottom: 2px solid #e8f0fe; color: #1967d2; font-weight: 600; font-size: 14px;
        display: flex; align-items: center;
      }
      .student-row { 
        display: flex; align-items: center; padding: 10px 12px; margin-bottom: 4px;
        background: white; border: 1px solid #dadce0; border-radius: 6px; transition: background 0.1s;
      }
      .student-row:hover { background: #f1f3f4; border-color: #d2e3fc; }
      
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
      
      <? if (mode === 'email') { ?>
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
       <div style="margin-bottom: 10px; display: flex; justify-content: flex-end;">
         <a class="action-link" onclick="toggleAll(true)">Select All</a>
         <a class="action-link" onclick="toggleAll(false)">Select None</a>
       </div>

       <? var lastSec = ""; ?>
       <? for (var i = 0; i < students.length; i++) { ?>
         <? if (students[i].section !== lastSec) { ?>
           <div class="section-header">
             <input type="checkbox" id="sec_chk_<?= students[i].sectionId ?>" 
                    onchange="toggleSection(this, '<?= students[i].sectionId ?>')" checked>
             <label for="sec_chk_<?= students[i].sectionId ?>" style="cursor: pointer;">
               <?= students[i].section ?>
             </label>
           </div>
           <? lastSec = students[i].section; ?>
         <? } ?>
         
         <div class="student-row">
           <input type="checkbox" id="chk_<?= i ?>" class="stu-chk <?= students[i].sectionId ?>" 
                  data-mismatch="<?= students[i].isMismatch ?>"
                  data-name="<?= students[i].name ?>"
                  value="<?= students[i].row ?>" checked>
           <label for="chk_<?= i ?>">
             <div><?= students[i].name ?> 
               <? if (students[i].isMismatch) { ?> <span class="badge badge-warn">Email Mismatch</span> <? } ?>
             </div>
             <div class="email-sub">
               <span>👤 Student: <?= students[i].email || '<i style="color:#b0b0b0;">None</i>' ?></span>
               <span>👥 Parent: <?= students[i].parentEmail || '<i style="color:#b0b0b0;">None</i>' ?></span>
             </div>
           </label>
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

      function toggleAll(state) {
        document.querySelectorAll('input[type="checkbox"]').forEach(c => c.checked = state);
      }

      function toggleSection(source, secId) {
        document.querySelectorAll('.' + secId).forEach(c => c.checked = source.checked);
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

  template.students = students;
  template.mode = mode;
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
    const data = sheet.getDataRange().getDisplayValues();
    
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
      warnings: [],
      successes: []
    };

    if (data.length < 2) {
      results.warnings.push("The active sheet is empty or has fewer than 2 rows. It must contain at least 4 rows to represent a proper Gradebook.");
      return results;
    }

    // Header checks inside Row 2
    const headers = data[1]; // Row 2 is 0-indexed index 1
    
    // 1. Check Name Column
    if (headers.length > 1 && headers[1] && headers[1].toLowerCase().includes("name")) {
      results.nameColFound = true;
      results.nameColLetter = "B";
      results.successes.push("Column B correctly designated as 'Name'.");
    } else {
      let foundNameIdx = headers.findIndex(h => h && h.toLowerCase().includes("name"));
      if (foundNameIdx !== -1) {
        results.nameColFound = true;
        results.nameColLetter = String.fromCharCode(65 + foundNameIdx);
        results.warnings.push(`'Name' column found in Column ${results.nameColLetter} instead of Column B. Keeping student names in Column B is highly recommended.`);
      } else {
        results.warnings.push("No column containing 'Name' was found in Row 2. You need a column named 'Name' (usually Column B) to identify students.");
      }
    }

    // 2. Search Email & Parent Email Columns
    let emailColIdx = -1;
    let parentEmailColIdx = -1;
    for (let c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      const cellText = headers[c].toLowerCase();
      if (cellText.includes('parent') || cellText.includes('guardian')) {
        parentEmailColIdx = c;
      } else if (cellText.includes('email') && !cellText.includes('parent') && !cellText.includes('guardian')) {
        emailColIdx = c;
      }
    }

    if (emailColIdx !== -1) {
      results.emailColFound = true;
      results.emailColLetter = String.fromCharCode(65 + emailColIdx);
      results.successes.push(`Student 'Email' column found in Column ${results.emailColLetter}.`);
    } else {
      results.warnings.push("No student 'Email' column was detected in Row 2. Add an 'Email' column to allow sending reports to students.");
    }

    if (parentEmailColIdx !== -1) {
      results.parentColFound = true;
      results.parentColLetter = String.fromCharCode(65 + parentEmailColIdx);
      results.successes.push(`Parent/Guardian 'Parent Email' column found in Column ${results.parentColLetter}.`);
    } else {
      results.warnings.push("No 'Parent Email' column was found. If you wish to send copies to parents, add a column named 'Parent Email' or 'Guardian Email' in Row 2.");
    }

    // 3. Scan backgrounds/data for Student Rows & Section Dividers
    const nameColIndex = 1;
    const backgrounds = sheet.getDataRange().getBackgrounds();
    const fontColors = sheet.getDataRange().getFontColors();
    let currentSection = "Ungrouped";

    for (let r = 3; r < data.length; r++) {
      const row = data[r];
      const firstCellBg = backgrounds[r][0];
      const firstCellFont = fontColors[r][0];
      const colA = row[0] ? String(row[0]).trim() : "";
      const name = row[nameColIndex] ? String(row[nameColIndex]).trim() : "";

      // DETECT SECTION HEADER
      if (colA !== "" && (
        firstCellBg === '#000000' ||
        firstCellBg === '#434343' ||
        (firstCellBg !== '#ffffff' && firstCellFont === '#ffffff') ||
        colA.toLowerCase().includes('block') ||
        (name === "" && !colA.includes("Average")))) {
        results.sectionCount++;
        currentSection = colA;
        continue;
      }

      // FILTER JUNK
      if (!name || name === '' || name === '0' ||
        name.includes('Name:') || name === 'Student' || name === 'Preferred Name' ||
        name.includes('Average') || name.includes('Sum') ||
        colA.includes('Average') || colA.includes('Sum')) {
        continue;
      }

      results.studentCount++;
    }

    if (results.studentCount > 0) {
      results.successes.push(`Parsed ${results.studentCount} active student rows starting at Row 4.`);
    } else {
      results.warnings.push("No active students detected below Row 3. Ensure student names are entered starting in Row 4, and Column B is designated as 'Name'.");
    }

    if (results.sectionCount > 0) {
      results.successes.push(`Detected ${results.sectionCount} class section dividers in Column A.`);
    } else {
      results.warnings.push("No section headers detected. To group students by class section, style a row with a solid background and input the section name (e.g. 'Block 1') in Column A.");
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
                <li><strong>Section Dividers</strong>: To group students into class periods, style a full row with a solid background (e.g., black or dark gray) and place the block name (e.g. <i>Block 1</i>) in Column A.</li>
                <li><strong>Student Rows</strong>: Student details and grade records. Formula cells or summary averages are automatically skipped.</li>
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

          // 5. Section Dividers Check
          if (res.sectionCount > 0) {
            html += '<div class="check-item check-success">' +
                    '<span class="check-icon">✅</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">Section Dividers</div>' +
                      '<div class="check-desc">Found <b>' + res.sectionCount + '</b> class dividers in Column A. Students will be categorized by class/block in the selector.</div>' +
                    '</div>' +
                  '</div>';
          } else {
            html += '<div class="check-item check-warning">' +
                    '<span class="check-icon">ℹ️</span>' +
                    '<div class="check-details">' +
                      '<div class="check-title">No Section Dividers (Optional)</div>' +
                      '<div class="check-desc">No group headers were found in Column A. To group students by period or class block, fill a row with a solid background and write the name in Column A.</div>' +
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