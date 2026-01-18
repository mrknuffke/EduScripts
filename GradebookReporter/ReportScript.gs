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

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gradebook Tools')
    .addItem('ℹ️ Help & Tutorial', 'showTutorialSidebar')
    .addSeparator()
    .addItem('📧 Email Reports (Selector)', 'openEmailSelector')
    .addItem('📂 Generate Reports (Drive)', 'openDriveSelector')
    .addSeparator()
    .addItem('📘 Generate Demo Gradebook', 'generateGradebookTemplate')
    .addToUi();
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
  const headers = ["Section", "Name", "Email", "Assignment 1", "Assignment 2", "Assignment 3", "Summative Exam"];
  const categories = ["", "", "", "Classwork", "Classwork", "Homework", "Assessments"];
  const standards = ["", "", "", "Standard 1", "Standard 1", "Standard 2", "Standard 3"];

  sheet.getRange(1, 1, 1, headers.length).setValues([categories]).setFontWeight("bold").setBackground("#e0e0e0");
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#434343").setFontColor("white");
  sheet.getRange(3, 1, 1, headers.length).setValues([standards]).setFontStyle("italic").setBackground("#f3f3f3");

  // Dummy Data
  const data = [
    ["Block 1", "Potter, Harry", "harry@hogwarts.edu", "1", "1", "0", "95"],
    ["Block 1", "Granger, Hermione", "hermione@hogwarts.edu", "1", "1", "1", "100"],
    ["Block 1", "Weasley, Ron", "ron@hogwarts.edu", "0", "1", "Missing", "85"],
    ["Block 2", "Malfoy, Draco", "draco@hogwarts.edu", "1", "Exempt", "1", "90"],
    ["Block 2", "Lovegood, Luna", "luna@hogwarts.edu", "1", "1", "1", "92"]
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

  // Find Email Column
  for (let r = 0; r < 2; r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (data[r][c] && data[r][c].toLowerCase().includes('email')) {
        emailColIndex = c;
        break;
      }
    }
    if (emailColIndex !== -1) break;
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
function runReportBatch(mode, rowIndices) {
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

    return processGradebook(sheet, reportTitle, subjectName, mode, rowIndices);
  } catch (e) {
    throw new Error("Line " + e.lineNumber + ": " + e.message);
  }
}

// --- REPORT GENERATION LOGIC ---
function processGradebook(sheet, titlePrefix, subjectName, mode, targetRows) {
  const data = sheet.getDataRange().getDisplayValues();
  const backgrounds = sheet.getDataRange().getBackgrounds();
  const fontColors = sheet.getDataRange().getFontColors();
  const notes = sheet.getDataRange().getNotes();

  const headerRowIndex = 1;
  const categoryRowIndex = 0;
  const standardsRowIndex = 2;
  const nameColIndex = 1;

  const coolMessages = [
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
    "You are a homework ninja. Silent, deadly, effective.", "Boom. Done. Everything submitted.", "You are officially on top of your game."
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
  const categories = data[categoryRowIndex];
  let emailColIndex = -1;

  for (let i = 0; i < headers.length; i++) { if (headers[i] && headers[i].toLowerCase().includes('email')) { emailColIndex = i; break; } }
  if (emailColIndex === -1 && categories) { for (let i = 0; i < categories.length; i++) { if (categories[i] && categories[i].toLowerCase().includes('email')) { emailColIndex = i; break; } } }

  let cutoffColIndex = headers.length;
  const headerBgColors = backgrounds[headerRowIndex];
  for (let i = nameColIndex + 1; i < headers.length; i++) {
    if (headerBgColors[i] === '#000000') { cutoffColIndex = i; break; }
  }

  const standards = data[standardsRowIndex];
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
      isSummativeStandard: false
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

      if (value) {
        const valStr = String(value).trim();
        const lowerVal = valStr.toLowerCase();
        if (lowerVal === 'true' || valStr === '1') displayValue = 'Complete';
        else if (valStr === '0' || lowerVal === 'm') { displayValue = 'Missing'; isIssue = true; }
        else if (lowerVal === 'ex') displayValue = 'Exempt';
        else if (lowerVal === 'i') { displayValue = 'Incomplete'; isIssue = true; }

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
      if (isIssue) shouldReport = true;
      if (subjectName === "Chemistry" && col.isSummativeStandard) shouldReport = true;

      if (shouldReport) {
        reportRows.push({
          category: col.finalCategory,
          name: col.finalName,
          value: displayValue,
          note: rowNotes[idx],
          bgColor: col.bgColor,
          fontColor: col.fontColor,
          rowBg: rowBgColors[idx],
          rowFont: rowFontColors[idx]
        });
      }
    });

    const hasActualMissingWork = reportRows.some(item =>
      item.value === 'Missing' ||
      item.value === 'Incomplete' || (subjectName === "AP Biology" && item.name.toLowerCase().startsWith("lab") && parseFloat(item.value) < 4)
    );

    // --- ACTION HANDLERS ---

    if (mode === 'drive') {
      if (processedCount > 0) docBody.appendPageBreak();
      renderToDoc(docBody, titlePrefix, studentName, reportRows, hasActualMissingWork, subjectName, coolMessages);
      processedCount++;
      if (processedCount % 5 === 0) { doc.saveAndClose(); doc = DocumentApp.openById(docId); docBody = doc.getBody(); }
    }

    else if (mode === 'preview') {
      if (previewCount >= 3) continue;
      const studentEmail = (emailColIndex > -1) ? row[emailColIndex] : "No Email";
      const htmlBody = generateHtmlReport(titlePrefix, studentName, reportRows, hasActualMissingWork, subjectName, coolMessages);

      previewHtml += `<div class="preview-box">
                        <div class="preview-header">PREVIEW ${previewCount + 1}: ${studentName} (${studentEmail})</div>
                        ${htmlBody}
                      </div>`;
      previewCount++;
    }

    else if (mode === 'email') {
      const studentEmail = (emailColIndex > -1) ? row[emailColIndex] : "";
      if (studentEmail && studentEmail.includes('@')) {
        const htmlBody = generateHtmlReport(titlePrefix, studentName, reportRows, hasActualMissingWork, subjectName, coolMessages);
        try {
          MailApp.sendEmail({
            to: studentEmail,
            replyTo: 'dknuffke@sas.edu.sg',
            subject: `${titlePrefix} - ${studentName}`,
            htmlBody: htmlBody
          });
          processedCount++;
        } catch (e) { Logger.log(`Email error ${studentName}: ${e.message}`); }
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
    const pSub = body.appendParagraph("\nStatus: No missing assignments. Nothing is owed at this time.");
    pSub.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(10).setForegroundColor('#555555');

    if (subjectName === "Chemistry" && rows.length > 0) {
      body.appendParagraph("\nSummative Standard Mastery:\n");
      printGroupedTableDoc(body, rows);
    }
  } else {
    printGroupedTableDoc(body, rows);
  }
}

function printGroupedTableDoc(body, rows) {
  const groups = {};
  const order = [];
  rows.forEach(r => {
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
    header.appendTableCell("Score").setBackgroundColor('#EFEFEF').setBold(true).setFontSize(9);

    groups[cat].forEach(item => {
      const tr = table.appendTableRow();
      const c1 = tr.appendTableCell(item.name);
      c1.setBackgroundColor(item.bgColor !== '#ffffff' ? item.bgColor : '#ffffff');
      c1.getChild(0).asParagraph().setFontSize(9).setBold(true).setForegroundColor(item.fontColor);
      c1.setPaddingTop(2).setPaddingBottom(2);
      const c2 = tr.appendTableCell(item.value);
      c2.setBackgroundColor(item.rowBg !== '#ffffff' ? item.rowBg : '#ffffff');
      c2.getChild(0).asParagraph().setFontSize(9).setForegroundColor(item.rowFont);
      c2.setPaddingTop(2).setPaddingBottom(2);
    });
  });
}

function generateHtmlReport(title, studentName, rows, hasMissing, subjectName, messages) {
  let html = `<div style="font-family: Arial, sans-serif; color: #333; max-width: 600px;">`;
  html += `<h2 style="text-align: center; color: #222;">${title}</h2>`;
  html += `<h3 style="text-align: center; color: #555;">${studentName}</h3>`;

  if (!hasMissing) {
    const msg = messages[Math.floor(Math.random() * messages.length)];
    html += `<div style="text-align: center; margin: 20px 0; padding: 15px; background-color: #e8f5e9; border-radius: 5px;">`;
    html += `<h3 style="color: #2E7D32; margin: 0;">${msg}</h3>`;
    html += `<p style="color: #555; font-size: 12px; margin-top: 5px;">Status: No missing assignments. Nothing is owed at this time.</p></div>`;

    if (subjectName === "Chemistry" && rows.length > 0) {
      html += `<h4 style="margin-top: 20px;">Summative Standard Mastery:</h4>`;
      html += generateHtmlTables(rows);
    }
  } else {
    html += generateHtmlTables(rows);
  }
  html += `<p style="font-size: 10px; color: #888; text-align: center; margin-top: 30px;">Generated by Gradebook Tools</p>`;
  html += `</div>`;
  return html;
}

function generateHtmlTables(rows) {
  const groups = {};
  const order = [];
  rows.forEach(r => {
    const c = r.category || "General";
    if (!groups[c]) { groups[c] = []; order.push(c); }
    groups[c].push(r);
  });

  let html = "";
  order.forEach(cat => {
    html += `<h4 style="margin-bottom: 5px; color: #444; border-bottom: 1px solid #ccc; padding-bottom: 3px;">${cat}</h4>`;
    html += `<table style="width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 15px;">`;
    html += `<tr style="background-color: #EFEFEF;">
              <th style="text-align: left; padding: 5px; border: 1px solid #ccc;">Assignment</th>
              <th style="text-align: left; padding: 5px; border: 1px solid #ccc;">Score</th>
             </tr>`;
    groups[cat].forEach(item => {
      const bgStyle = item.bgColor !== '#ffffff' ? `background-color: ${item.bgColor};` : '';
      const fontStyle = `color: ${item.fontColor}; font-weight: bold;`;
      const rowBgStyle = item.rowBg !== '#ffffff' ? `background-color: ${item.rowBg};` : '';
      const rowFontStyle = `color: ${item.rowFont};`;
      html += `<tr>
                <td style="padding: 5px; border: 1px solid #ccc; ${bgStyle} ${fontStyle}">${item.name}</td>
                <td style="padding: 5px; border: 1px solid #ccc; ${rowBgStyle} ${rowFontStyle}">${item.value}</td>
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

      /* Content Area */
      #content { flex: 1; overflow-y: auto; padding: 10px 20px; }
      
      /* Sections & Rows */
      .section-header { 
        margin-top: 20px; margin-bottom: 8px; padding-bottom: 5px;
        border-bottom: 2px solid #e8f0fe; color: #1967d2; font-weight: 600; font-size: 14px;
        display: flex; align-items: center;
      }
      .student-row { 
        display: flex; align-items: center; padding: 8px 12px; margin-bottom: 2px;
        background: white; border: 1px solid #dadce0; border-radius: 4px; transition: background 0.1s;
      }
      .student-row:hover { background: #f1f3f4; border-color: #d2e3fc; }
      
      /* Controls */
      input[type="checkbox"] { transform: scale(1.1); margin-right: 12px; cursor: pointer; }
      label { flex: 1; cursor: pointer; font-size: 14px; display: flex; flex-direction: column; justify-content: center; }
      .email-sub { font-size: 11px; color: #70757a; margin-top: 2px; }
      
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
             <div class="email-sub"><?= students[i].email ?></div>
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
          .runReportBatch(action, selected);
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