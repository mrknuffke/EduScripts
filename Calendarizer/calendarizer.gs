/**
 * @OnlyCurrentDoc
 */

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

// --- CONFIGURATION MANAGEMENT ---
const DEFAULT_CONFIG = {
  // Start month of the school year (0 = Jan, 6 = Jul)
  startMonth: 6,

  // Keywords to identify holidays in the notes column
  holidayKeywords: [
    'holiday', 'break', 'conference', 'inservice', 'start',
    'psat', 'day', 'confs', 'egg'
  ],

  // Styles and Colors
  styles: {
    fontFamily: 'Roboto, sans-serif',
    borderColor: '#dadce0', // Subtle Google Grey
    header: { background: '#f1f3f4', fontColor: '#5f6368' },
    day: { weekday: '#ffffff' },
    weekend: { background: '#5f6368', fontColor: '#ffffff' }, // Darker grey with white text
    emptyDay: { background: '#5f6368' },

    // Modern Pastel Palette (Month 1 -> Month 12)
    // High contrast pairs: Light background, Darker font accent
    monthColors: [
      { background: '#e1bee7', font: '#4a148c' }, // Purple
      { background: '#d1c4e9', font: '#311b92' }, // Deep Purple
      { background: '#c5cae9', font: '#1a237e' }, // Indigo
      { background: '#bbdefb', font: '#0d47a1' }, // Blue
      { background: '#b3e5fc', font: '#01579b' }, // Light Blue
      { background: '#b2ebf2', font: '#006064' }, // Cyan
      { background: '#b2dfdb', font: '#004d40' }, // Teal
      { background: '#c8e6c9', font: '#1b5e20' }, // Green
      { background: '#dcedc8', font: '#33691e' }, // Light Green
      { background: '#fff9c4', font: '#f57f17' }, // Yellow
      { background: '#ffecb3', font: '#ff6f00' }, // Amber
      { background: '#ffe0b2', font: '#e65100' }, // Orange
    ]
  }
};

/**
 * Loads configuration from Document Properties, merging with defaults.
 */
function loadConfig() {
  const props = PropertiesService.getDocumentProperties();
  const saved = props.getProperty('CALENDAR_CONFIG');
  if (saved) {
    const parsed = JSON.parse(saved);
    // Merge styles carefully to ensure new properties like fontFamily exist
    const merged = { ...DEFAULT_CONFIG, ...parsed };
    merged.styles = { ...DEFAULT_CONFIG.styles, ...(parsed.styles || {}) };
    return merged;
  }
  return DEFAULT_CONFIG;
}

/**
 * Saves configuration to Document Properties.
 */
function saveConfig(newConfig) {
  PropertiesService.getDocumentProperties().setProperty('CALENDAR_CONFIG', JSON.stringify(newConfig));
}


// --- CUSTOM MENU ---
/**
 * Creates a custom menu in the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendar Tools')
    .addItem('ℹ️ Help & Tutorial', 'showTutorialSidebar')
    .addItem('📄 Create Pacing Template', 'createPacingTemplate')
    .addItem('⚙️ Configuration', 'configureSettings')
    .addSeparator()
    .addItem('Create Wall Calendar View', 'createWallCalendar')
    .addItem('Create Lateral Calendar View', 'createLateralCalendar')
    .addToUi();
}

/**
 * Creates a new Pacing Chart Template with intelligent formulas.
 */
function createPacingTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const date = new Date();
  const currentYear = date.getFullYear();
  const defaultName = `Draft ${currentYear}-${currentYear + 1}`;

  const prompt = ui.prompt(
    'Create Pacing Template',
    `Enter a name for the new sheet (e.g., "${defaultName}"):`,
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const sheetName = prompt.getResponseText().trim();
  if (!sheetName) {
    ui.alert("Sheet name cannot be empty.");
    return;
  }

  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    const result = ui.alert(
      'Sheet Already Exists',
      `The sheet "${sheetName}" already exists.\n\nDo you want to delete it and create a new template? (This cannot be undone)`,
      ui.ButtonSet.YES_NO
    );

    if (result === ui.Button.YES) {
      ss.deleteSheet(sheet);
    } else {
      return;
    }
  }

  sheet = ss.insertSheet(sheetName);

  // 1. Set Headers
  const headers = [
    ["Class #", "Date 1", "Date 2", "Day", "Notes", "CLASS NAME 1", "HW/Notes", "CLASS NAME 2", "HW/Notes"]
  ];
  sheet.getRange("A1:I1").setValues(headers)
    .setFontWeight("bold")
    .setBackground("#0C343D") // Dark Teal
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // 2. Set Initial Row (Manual Start)
  // We'll calculate a likely start date (e.g., closest Monday to Sept 1st)
  // But for safety, just put a placeholder date.
  const row2 = [
    1,
    new Date(currentYear, 8, 1), // Sept 1st
    '=IF(TEXT(B2, "ddd") = "Fri", B2 + 3, IF(WEEKDAY(B2) = 7, B2 + 2, B2 + 1))', // Date 2 formula for row 2
    "A/B",
    "Start of Year"
  ];

  // 3. Set Daisy Chain Formulas (Row 3) - These are the ones users drag down
  // Note: We use R1C1 notation or simple A1 relative references.
  // B3 (Date 1): Look at C2 (Prev Date 2)
  const formulaDate1 = '=IF(TEXT(C2, "ddd") = "Fri", C2 + 3, IF(WEEKDAY(C2) = 7, C2 + 2, C2 + 1))';

  // C3 (Date 2): Look at B3 (Curr Date 1)
  const formulaDate2 = '=IF(TEXT(B3, "ddd") = "Fri", B3 + 3, IF(WEEKDAY(B3) = 7, B3 + 2, B3 + 1))';

  // D3 (Toggle): Look at D2
  const formulaBlock = '=IF(OR(D2="A/B", D2="A", D2="B"), "C/D", IF(OR(D2="C/D", D2="C", D2="D"), "A/B", ""))';

  // A3 (Class increment)
  const formulaClass = '=A2+1';

  const row3 = [
    formulaClass,
    formulaDate1,
    formulaDate2,
    formulaBlock,
    "" // Notes empty
  ];

  // Write Row 2
  sheet.getRange(2, 1, 1, 5).setValues([row2]);

  // Write Row 3 and fill down
  const rangeRow3 = sheet.getRange(3, 1, 1, 5);
  rangeRow3.setValues([row3]);

  // Fill down formulas for ~20 rows to get them started
  rangeRow3.autoFill(sheet.getRange(3, 1, 20, 5), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // 4. Formatting
  sheet.setColumnWidths(1, 1, 60); // Class #
  sheet.setColumnWidths(2, 2, 100); // Dates
  sheet.setColumnWidths(4, 1, 80); // Day
  sheet.setColumnWidths(5, 1, 200); // Notes
  sheet.setColumnWidths(6, 4, 150); // Class columns

  // Date formatting
  sheet.getRange("B2:C100").setNumberFormat("ddd, d. m");

  SpreadsheetApp.getUi().alert(`Created "${sheetName}" with template formulas.\n\nTry changing the start date in B2 to see the chain update!`);
}

// --- DATA PARSING LOGIC ---
/**
 * Helper function to convert month abbreviations to a month index (0-11).
 */
function getMonthIndex(monthAbbr) {
  // ... rest of file calls from line 109 ...

  // WE NEED TO SKIP TO THE TUTORIAL FUNCTION AT THE END TO UPDATE IT
  // Since I can't skip content in a single replace block unless I include everything between,
  // I will structure this edit to ONLY replace onOpen and ADD the new function right after it.
  // Then I will do a SECOND edit for the tutorial HTML.
  // WAITING - ACTUALLY I CAN DO IT ALL IF I AM CAREFUL OR I WILL JUST DO TWO EDITS.
  // I WILL DO TWO EDITS TO BE SAFE. THIS TOOL CALL WILL JUST ADD THE FUNCTION AND UPDATE MENU.
}


// --- DATA PARSING LOGIC ---
/**
 * Helper function to convert month abbreviations to a month index (0-11).
 */
function getMonthIndex(monthAbbr) {
  if (!monthAbbr) return undefined;
  const months = { 'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'may': 4, 'jun': 5, 'jul': 6, 'aug': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11 };
  return months[monthAbbr.substring(0, 3).toLowerCase()];
}

/**
 * Parses the pacing chart data.
 */
function parsePacingData(sheet, startYear, config) {
  const data = sheet.getDataRange().getValues();
  const eventMap = {};
  const timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  let currentYear = startYear;

  // Helper to initialize day object
  const initDay = () => ({ block: '', holiday: '', classInfo: [], notes: [] });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const classNumber = row[0];
    const date1 = row[1] ? new Date(row[1]) : null;
    const date2 = row[2] ? new Date(row[2]) : null;
    const dayType = row[3] || '';
    const notes = row[4] || '';

    // Adjust year if date passes threshold (simple heuristic)
    if (date1 && date1.getMonth() < config.startMonth) {
      currentYear = startYear + 1;
    }

    if (classNumber && !isNaN(classNumber)) {
      if (date1 && date1.getTime()) {
        const key1 = Utilities.formatDate(date1, timeZone, "yyyy-MM-dd");
        if (!eventMap[key1]) eventMap[key1] = initDay();
        eventMap[key1].block = (dayType === 'A/B') ? 'A' : (dayType === 'C/D') ? 'C' : dayType;
        eventMap[key1].classInfo.push(`Class ${classNumber}`);
      }
      if (date2 && date2.getTime() && (!date1 || date1.getTime() !== date2.getTime())) {
        const key2 = Utilities.formatDate(date2, timeZone, "yyyy-MM-dd");
        if (!eventMap[key2]) eventMap[key2] = initDay();
        eventMap[key2].block = (dayType === 'A/B') ? 'B' : (dayType === 'C/D') ? 'D' : dayType;
        eventMap[key2].classInfo.push(`Class ${classNumber}`);
      }
    } else if (!classNumber && dayType && date1 && date1.getTime()) {
      const key1 = Utilities.formatDate(date1, timeZone, "yyyy-MM-dd");
      if (!eventMap[key1]) eventMap[key1] = initDay();
      eventMap[key1].holiday = dayType;
    }

    if (!notes) continue;
    const noteLines = notes.split('\n');
    for (const line of noteLines) {
      const crossMonthRegex = /(\d{1,2})\s+([a-zA-Z]+)\s*-\s*(\d{1,2})\s+([a-zA-Z]+)/;
      const rangeRegex = /(\d{1,2})-(\d{1,2})\s+([a-zA-Z]+)/;
      const singleDayRegex = /(\d{1,2})\s+([a-zA-Z]+)/;
      let match;
      const description = line.includes(':') ? line.split(':').slice(1).join(':').trim() : line;

      if ((match = line.match(crossMonthRegex))) {
        const startDay = parseInt(match[1]), startMonth = getMonthIndex(match[2]);
        const endDay = parseInt(match[3]), endMonth = getMonthIndex(match[4]);
        if (startMonth === undefined || endMonth === undefined) continue;
        let startYear = currentYear, endYear = (startMonth > endMonth) ? currentYear + 1 : currentYear;
        const startDate = new Date(startYear, startMonth, startDay);
        const endDate = new Date(endYear, endMonth, endDay);
        for (let d = new Date(startDate.getTime()); d <= endDate; d.setDate(d.getDate() + 1)) {
          const key = Utilities.formatDate(d, timeZone, "yyyy-MM-dd");
          if (!eventMap[key]) eventMap[key] = initDay();
          eventMap[key].holiday = description;
        }
      } else if ((match = line.match(rangeRegex))) {
        const startDay = parseInt(match[1]), endDay = parseInt(match[2]), monthIndex = getMonthIndex(match[3]);
        if (monthIndex === undefined) continue;
        for (let day = startDay; day <= endDay; day++) {
          const eventDate = new Date(currentYear, monthIndex, day);
          const key = Utilities.formatDate(eventDate, timeZone, "yyyy-MM-dd");
          if (!eventMap[key]) eventMap[key] = initDay();
          eventMap[key].holiday = description;
        }
      } else if ((match = line.match(singleDayRegex))) {
        const day = parseInt(match[1]), monthIndex = getMonthIndex(match[2]);
        if (monthIndex === undefined) continue;
        const eventDate = new Date(currentYear, monthIndex, day);
        const key = Utilities.formatDate(eventDate, timeZone, "yyyy-MM-dd");
        if (!eventMap[key]) eventMap[key] = initDay();

        if (config.holidayKeywords.some(k => description.toLowerCase().includes(k))) {
          eventMap[key].holiday = description;
        } else {
          eventMap[key].notes.push(description);
        }
      }
    }
  }
  return eventMap;
}


// --- RICH TEXT HELPER FUNCTIONS ---
/**
 * Shared helper to build a RichTextValue with bolded segments.
 */
function createRichText(textParts, boldSegments) {
  const fullText = textParts.join('');
  if (!fullText) {
    return SpreadsheetApp.newRichTextValue().setText('').build();
  }

  const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
  const builder = SpreadsheetApp.newRichTextValue().setText(fullText);

  boldSegments.forEach(segment => {
    // Ensure bounds are valid
    if (segment.start >= 0 && segment.end <= fullText.length) {
      builder.setTextStyle(segment.start, segment.end, boldStyle);
    }
  });

  return builder.build();
}

/**
 * Builds a RichTextValue for a cell in the Wall Calendar. Bolds holidays and notes.
 */
function buildWallCalendarRichText(dayData, dayNumber, dayOfWeekStr) {
  const textParts = [];
  const boldSegments = [];
  let currentLength = 0;

  const firstLineNormal = `${dayNumber} ${dayOfWeekStr} `;
  textParts.push(firstLineNormal);
  currentLength += firstLineNormal.length;

  if (dayData) {
    if (dayData.block) {
      const blockText = `(${dayData.block}) `;
      textParts.push(blockText);
      currentLength += blockText.length;
    }
    if (dayData.holiday) {
      const holidayText = `- ${dayData.holiday}`;
      boldSegments.push({ start: currentLength, end: currentLength + holidayText.length });
      textParts.push(holidayText);
      currentLength += holidayText.length;
    }

    dayData.classInfo.forEach(info => {
      const classText = `\n${info}`;
      textParts.push(classText);
      currentLength += classText.length;
    });
    dayData.notes.forEach(note => {
      const noteText = `\n${note}`;
      boldSegments.push({ start: currentLength, end: currentLength + noteText.length });
      textParts.push(noteText);
      currentLength += noteText.length;
    });
  }

  // Add an empty line at the end so clicking the cell places the cursor on a new line
  textParts.push('\n');

  return createRichText(textParts, boldSegments);
}

/**
 * Builds a RichTextValue for a cell in the Lateral Calendar. Bolds holidays and notes.
 */
function buildLateralCalendarRichText(dayData, dayOfWeekStr) {
  const textParts = [dayOfWeekStr];
  const boldSegments = [];
  let currentLength = dayOfWeekStr.length;

  if (dayData) {
    if (dayData.block) {
      const blockText = `     (${dayData.block})`;
      textParts.push(blockText);
      currentLength += blockText.length;
    }
    if (dayData.holiday) {
      const holidayText = `\n${dayData.holiday}`;
      boldSegments.push({ start: currentLength, end: currentLength + holidayText.length });
      textParts.push(holidayText);
      currentLength += holidayText.length;
    }
    dayData.classInfo.forEach(info => {
      const classText = `\n${info}`;
      textParts.push(classText);
      currentLength += classText.length;
    });
    dayData.notes.forEach(note => {
      const noteText = `\n${note}`;
      boldSegments.push({ start: currentLength, end: currentLength + noteText.length });
      textParts.push(noteText);
      currentLength += noteText.length;
    });
  }

  // Add an empty line at the end so clicking the cell places the cursor on a new line
  textParts.push('\n');

  return createRichText(textParts, boldSegments);
}


// --- WALL CALENDAR DRAWING FUNCTION ---
/**
 * Draws a single month's calendar grid onto a sheet using Rich Text.
 * Now optimized to batch RichText writes.
 */
function drawMonth(sheet, year, month, startRow, eventData, monthCounter, config) {
  const timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const monthName = Utilities.formatDate(new Date(year, month), timeZone, "MMMM yyyy");

  const currentMonthStyle = config.styles.monthColors[monthCounter % 12];
  const borderCol = config.styles.borderColor || '#999';

  // Set consistent row heights for headers
  sheet.setRowHeight(startRow, 50);
  sheet.setRowHeight(startRow + 1, 30);

  // Apply month header styling.
  sheet.getRange(startRow, 1, 1, 7).merge().setValue(monthName)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontWeight('bold')
    .setFontSize(22)
    .setFontFamily(config.styles.fontFamily)
    .setBackground(currentMonthStyle.background)
    .setFontColor(currentMonthStyle.font)
    .setBorder(true, true, true, true, null, null, borderCol, SpreadsheetApp.BorderStyle.SOLID);

  // Apply weekday header styling.
  const weekDays = ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'];
  sheet.getRange(startRow + 1, 1, 1, 7).setValues([weekDays])
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontFamily(config.styles.fontFamily)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(config.styles.header.background)
    .setFontColor(config.styles.header.fontColor)
    .setBorder(null, true, true, true, true, null, borderCol, SpreadsheetApp.BorderStyle.SOLID);

  const gridStartRow = startRow + 2;
  const backgroundColors = [];
  const fontColors = [];
  const richTextGrid = []; // 2D array for batched writing

  let currentDay = 1;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const firstDayOfWeek = new Date(year, month, 1).getDay();

  for (let week = 0; week < 6; week++) {
    const currentRow = gridStartRow + week;
    const rowBackgrounds = [];
    const rowFontColors = [];
    const rowRichTexts = [];

    sheet.setRowHeight(currentRow, 120);

    for (let day = 0; day < 7; day++) {
      // Determine if cell is part of the month
      const isDayCell = !((week === 0 && day < firstDayOfWeek) || currentDay > daysInMonth);

      if (!isDayCell) {
        rowBackgrounds.push(config.styles.emptyDay.background);
        rowFontColors.push('#000000');
        rowRichTexts.push(SpreadsheetApp.newRichTextValue().setText('').build());
      } else {
        // Background Logic
        // Use a subtle alternative color for weekends if desired, else check config
        if (day === 0 || day === 6) {
          rowBackgrounds.push(config.styles.weekend ? config.styles.weekend.background : config.styles.emptyDay.background);
          rowFontColors.push(config.styles.weekend ? config.styles.weekend.fontColor : '#000000');
        } else {
          rowBackgrounds.push(config.styles.day.weekday);
          rowFontColors.push('#000000');
        }

        // Content Logic
        const currentDate = new Date(year, month, currentDay);
        const dayOfWeekStr = Utilities.formatDate(currentDate, timeZone, "E");
        const dateKey = Utilities.formatDate(currentDate, timeZone, "yyyy-MM-dd");
        const dayData = eventData[dateKey];

        rowRichTexts.push(buildWallCalendarRichText(dayData, currentDay, dayOfWeekStr));

        currentDay++;
      }
    }
    backgroundColors.push(rowBackgrounds);
    fontColors.push(rowFontColors);
    richTextGrid.push(rowRichTexts);

    if (currentDay > daysInMonth) break;
  }

  const numWeeks = backgroundColors.length;
  if (numWeeks > 0) {
    const range = sheet.getRange(gridStartRow, 1, numWeeks, 7);
    range.setBackgrounds(backgroundColors);
    range.setFontColors(fontColors);
    range.setRichTextValues(richTextGrid); // BATCHED WRITE

    range.setVerticalAlignment('top').setWrap(true).setFontSize(10).setFontFamily(config.styles.fontFamily);
    range.setBorder(true, true, true, true, true, true, borderCol, SpreadsheetApp.BorderStyle.SOLID);
  }

  return gridStartRow + numWeeks;
}

// --- LATERAL CALENDAR DRAWING FUNCTION ---
/**
 * Draws the lateral calendar layout using the new color scheme while keeping the data.
 */
function drawLateralCalendarLayout(sheet, eventData, startYear, config) {
  const timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const schoolYearStart = new Date(startYear, config.startMonth, 1);
  const borderCol = config.styles.borderColor || '#999';

  // --- HEADER SETUP ---
  const dayNumberHeader = ['MONTH'];
  for (let day = 1; day <= 31; day++) {
    dayNumberHeader.push(day);
  }

  const headerRange = sheet.getRange(1, 1, 1, 32);
  headerRange.setValues([dayNumberHeader])
    .setBackground(config.styles.header.background)
    .setFontColor(config.styles.header.fontColor)
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontFamily(config.styles.fontFamily)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, borderCol, SpreadsheetApp.BorderStyle.SOLID);

  // --- CALENDAR GRID GENERATION ---

  const allBackgrounds = [];
  const allFontColors = [];
  const allRichTexts = [];
  const allMonthNames = [];
  const startRow = 2;
  const totalMonths = 12;

  for (let i = 0; i < totalMonths; i++) {
    const currentDate = new Date(schoolYearStart.getFullYear(), schoolYearStart.getMonth() + i, 1);
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth();

    const monthAbbr = Utilities.formatDate(currentDate, timeZone, "MMM").toUpperCase();
    allMonthNames.push([monthAbbr]);

    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const currentMonthStyle = config.styles.monthColors[i % 12];

    const rowRichTexts = [];
    const rowBackgrounds = [];
    const rowFontColors = [];

    // Note: The grid is 31 columns wide (Days 1-31)
    for (let day = 1; day <= 31; day++) {
      if (day > daysInMonth) {
        rowBackgrounds.push(config.styles.emptyDay.background);
        rowFontColors.push('#000000');
        rowRichTexts.push(SpreadsheetApp.newRichTextValue().setText('').build());
      } else {
        const date = new Date(year, month, day);
        const dayOfWeek = date.getDay(); // 0 = Sunday, 6 = Saturday

        if (dayOfWeek === 0 || dayOfWeek === 6) { // Weekend
          rowBackgrounds.push(config.styles.weekend ? config.styles.weekend.background : config.styles.emptyDay.background);
          rowFontColors.push(config.styles.weekend ? config.styles.weekend.fontColor : '#000000');
        } else { // Weekday
          rowBackgrounds.push(config.styles.day.weekday);
          rowFontColors.push('#000000');
        }

        const dayOfWeekStr = Utilities.formatDate(date, timeZone, "E");
        const dateKey = Utilities.formatDate(date, timeZone, "yyyy-MM-dd");
        const dayData = eventData[dateKey];
        rowRichTexts.push(buildLateralCalendarRichText(dayData, dayOfWeekStr));
      }
    }
    allBackgrounds.push(rowBackgrounds);
    allFontColors.push(rowFontColors);
    allRichTexts.push(rowRichTexts);
  }

  // --- BATCH WRITE ---
  // 0. Global formatting (Defaults)
  // Apply defaults first so specific column formatting can override them.
  sheet.getRange(startRow, 1, totalMonths, 32)
    .setVerticalAlignment('top')
    .setWrap(true)
    .setFontSize(9)
    .setFontFamily(config.styles.fontFamily)
    .setBorder(true, true, true, true, true, true, borderCol, SpreadsheetApp.BorderStyle.SOLID);

  // 1. Month Names Column (Col 1)
  const monthColRange = sheet.getRange(startRow, 1, totalMonths, 1);
  monthColRange.setValues(allMonthNames);

  // Style month names
  const monthColBackgrounds = [];
  const monthColFontColors = [];

  for (let i = 0; i < totalMonths; i++) {
    const style = config.styles.monthColors[i % 12];
    monthColBackgrounds.push([style.background]); // Colored background for the sidebar
    monthColFontColors.push([style.font]); // Font color
  }
  monthColRange.setBackgrounds(monthColBackgrounds).setFontColors(monthColFontColors);
  monthColRange.setFontWeight('bold').setFontSize(48).setFontFamily(config.styles.fontFamily)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, true, borderCol, SpreadsheetApp.BorderStyle.SOLID);

  // 2. Calendar Grid (Cols 2-32)
  const gridRange = sheet.getRange(startRow, 2, totalMonths, 31);
  gridRange.setBackgrounds(allBackgrounds);
  gridRange.setFontColors(allFontColors);
  gridRange.setRichTextValues(allRichTexts);

  // 3. Row formatting
  for (let i = 0; i < totalMonths; i++) {
    sheet.setRowHeight(startRow + i, 115);
  }

  // --- FINAL FORMATTING ---
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidths(2, 31, 120);
}


// --- MAIN FUNCTIONS ---
/**
 * Main function to generate the wall calendar view.
 */
function createWallCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheet = ss.getActiveSheet();
  const sourceSheetName = sourceSheet.getName();
  const targetSheetName = "Wall Calendar View";

  const config = loadConfig();

  const startYearMatch = sourceSheetName.match(/^(\d{4})-\d{4}$/);
  if (!startYearMatch) {
    ui.alert('Invalid Sheet Name', 'The active sheet name must be in the format "YYYY-YYYY" (e.g., "2025-2026").', ui.ButtonSet.OK);
    return;
  }
  const startYear = parseInt(startYearMatch[1], 10);

  const eventData = parsePacingData(sourceSheet, startYear, config);

  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    ss.deleteSheet(targetSheet);
  }
  targetSheet = ss.insertSheet(targetSheetName);

  // Adjusted column widths for cleaner fit
  targetSheet.setColumnWidths(1, 7, 140);

  let currentRowOnSheet = 1;
  // Use a predictable loop to avoid date object mutation issues.
  for (let i = 0; i < 12; i++) {
    // School year starts in CONFIG.startMonth
    const currentMonthDate = new Date(startYear, config.startMonth + i, 1);
    const year = currentMonthDate.getFullYear();
    const month = currentMonthDate.getMonth();

    currentRowOnSheet = drawMonth(targetSheet, year, month, currentRowOnSheet, eventData, i, config);

    // Only add a blank spacer row if it's not the last month.
    if (i < 11) {
      currentRowOnSheet += 1;
    }
  }

  // Hide extra columns for neatness
  const maxCols = targetSheet.getMaxColumns();
  if (maxCols > 7) {
    targetSheet.deleteColumns(8, maxCols - 7);
  }

  // After generating content, remove any surplus empty rows at the bottom.
  const lastRow = targetSheet.getLastRow();
  const maxRows = targetSheet.getMaxRows();
  if (maxRows > lastRow) {
    targetSheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }

  ui.alert('Success! Your "Wall Calendar View" sheet has been created.');
}

/**
 * Main function to generate the lateral calendar view.
 */
function createLateralCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheet = ss.getActiveSheet();
  const sourceSheetName = sourceSheet.getName();
  const targetSheetName = "Lateral Calendar View";

  const config = loadConfig();

  const startYearMatch = sourceSheetName.match(/^(\d{4})-\d{4}$/);
  if (!startYearMatch) {
    ui.alert('Invalid Sheet Name', 'The active sheet name must be in the format "YYYY-YYYY" (e.g., "2025-2026").', ui.ButtonSet.OK);
    return;
  }
  const startYear = parseInt(startYearMatch[1], 10);

  const eventData = parsePacingData(sourceSheet, startYear, config);

  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    targetSheet.clear();
    targetSheet.setFrozenRows(0);
    targetSheet.setFrozenColumns(0);
  } else {
    targetSheet = ss.insertSheet(targetSheetName);
  }

  drawLateralCalendarLayout(targetSheet, eventData, startYear, config);

  const maxRows = targetSheet.getMaxRows();
  const maxCols = targetSheet.getMaxColumns();

  if (maxRows > 13) {
    targetSheet.deleteRows(14, maxRows - 13);
  }
  if (maxCols > 32) {
    targetSheet.deleteColumns(33, maxCols - 32);
  }

  ui.alert('Success! Your "Lateral Calendar View" sheet has been created.\n\nFor best printing results, please use Landscape orientation, set the scale to "Fit to height", and ensure Frozen Rows/Columns are turned OFF.');
}


// --- SETTINGS DIALOG ---
/**
 * Opens the settings modal.
 */
function configureSettings() {
  const config = loadConfig();
  const html = HtmlService.createHtmlOutput(buildSettingsHtml(config))
    .setWidth(600)
    .setHeight(650)
    .setTitle('Calendarizer Configuration');
  SpreadsheetApp.getUi().showModalDialog(html, 'Calendarizer Configuration');
}

function buildSettingsHtml(config) {
  // Pass config safely to client-side
  const safeConfig = JSON.stringify(config);

  return `
    <style>
      body { font-family: 'Segoe UI', Roboto, sans-serif; padding: 20px; color: #333; background-color: #f9f9f9; }
      h3 { margin-top: 0; color: #1a73e8; font-size: 16px; border-bottom: 2px solid #e0e0e0; padding-bottom: 8px; margin-bottom: 15px; }
      
      .section { background: white; border: 1px solid #ddd; border-radius: 8px; padding: 15px; margin-bottom: 15px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
      
      label { display: block; font-weight: 500; font-size: 13px; margin-bottom: 5px; color: #555; }
      select, input[type="text"] { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; box-sizing: border-box; }
      .desc { font-size: 12px; color: #666; margin-top: 4px; margin-bottom: 10px; line-height: 1.4; }
      
      .color-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; }
      .color-item { text-align: center; }
      .color-item label { font-size: 11px; margin-bottom: 2px; }
      input[type="color"] { width: 100%; height: 30px; border: none; cursor: pointer; background: none; }
      
      .btn-container { text-align: right; margin-top: 20px; padding-top: 15px; border-top: 1px solid #ddd; }
      button { padding: 10px 20px; border-radius: 4px; font-weight: 600; cursor: pointer; border: none; font-size: 14px; }
      .btn-save { background: #1a73e8; color: white; }
      .btn-save:hover { background: #1557b0; }
      .btn-cancel { background: transparent; color: #666; margin-right: 10px; }
      .btn-cancel:hover { text-decoration: underline; }
    </style>

    <div class="section">
      <h3>📅 General Settings</h3>
      <label>School Year Start Month</label>
      <select id="startMonth">
        <option value="0">January</option>
        <option value="5">June</option>
        <option value="6">July</option>
        <option value="7">August</option>
        <option value="8">September</option>
      </select>
      <div class="desc">Determines which month comes first in the calendar year cycle.</div>
    </div>

    <div class="section">
      <h3>🔍 Parsing Rules</h3>
      <label>Holiday Keywords (Comma Separated)</label>
      <input type="text" id="keywords" placeholder="holiday, break, no school...">
      <div class="desc">Rows in your sheet containing these words in the "Notes" column will be treated as full holidays (bolded, no class info).</div>
    </div>

    <div class="section">
      <h3>🎨 Month Colors</h3>
      <div class="desc">Customize the background color for each month in the cycle (Month 1 = Start Month).</div>
      <div class="color-grid" id="colorGrid"></div>
    </div>

    <div class="btn-container">
      <button class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
      <button class="btn-save" id="saveBtn" onclick="save()">Save Settings</button>
    </div>

    <script>
      const CONFIG = ${safeConfig};
      
      // Init
      document.getElementById('startMonth').value = CONFIG.startMonth;
      document.getElementById('keywords').value = CONFIG.holidayKeywords.join(', ');
      
      // Build Color Inputs
      const monthNames = ["Month 1", "Month 2", "Month 3", "Month 4", "Month 5", "Month 6", "Month 7", "Month 8", "Month 9", "Month 10", "Month 11", "Month 12"];
      const grid = document.getElementById('colorGrid');
      
      CONFIG.styles.monthColors.forEach((style, idx) => {
         const div = document.createElement('div');
         div.className = 'color-item';
         div.innerHTML = \`
           <label>\${monthNames[idx]}</label>
           <input type="color" id="color_\${idx}" value="\${style.background}">
         \`;
         grid.appendChild(div);
      });

      function save() {
        const btn = document.getElementById('saveBtn');
        btn.textContent = 'Saving...';
        btn.disabled = true;

        const newConfig = JSON.parse(JSON.stringify(CONFIG)); // Clone
        
        // Update Values
        newConfig.startMonth = parseInt(document.getElementById('startMonth').value);
        newConfig.holidayKeywords = document.getElementById('keywords').value.split(',').map(s => s.trim()).filter(s => s !== '');
        
        // Update Colors
        for (let i = 0; i < 12; i++) {
           const col = document.getElementById('color_' + i).value;
           newConfig.styles.monthColors[i].background = col;
        }

        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .withFailureHandler((err) => { alert('Error: ' + err); btn.disabled = false; btn.textContent = 'Save Settings'; })
          .saveConfig(newConfig);
      }
    </script>
  `;
}


// --- TUTORIAL SIDEBAR ---
/**
 * Shows the tutorial sidebar.
 */
function showTutorialSidebar() {
  const html = HtmlService.createHtmlOutput(buildTutorialHtml())
    .setTitle('Calendarizer Guide')
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
      .section-title { font-weight: 700; color: #5f6368; margin-bottom: 5px; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 15px; }
      .menu-item { margin-bottom: 8px; }
      .menu-name { font-weight: 600; color: #202124; }
      .menu-desc { font-size: 13px; color: #5f6368; margin-top: 2px; }
      code { background: #eee; padding: 2px 4px; border-radius: 4px; font-family: monospace; font-size: 12px; }
      .tip { background-color: #e8f0fe; border-left: 3px solid #1a73e8; padding: 8px; margin: 10px 0; font-size: 13px; }
    </style>
    
    <h3>🗓️ Calendarizer Guide</h3>
    <p>Visualize your pacing chart as a Wall Calendar or Lateral Calendar.</p>

    <h3>⚡ How the Sheet Works</h3>
    <div class="card">
      <div style="font-weight:bold; margin-bottom:10px;">The "Daisy Chain" Logic</div>
      <p style="margin-bottom:10px;">This sheet uses formulas to automatically calculate school days. Each cell depends on the one before it:</p>
      
      <div class="menu-item">
        <div class="menu-name">1. Date 1 (Column B)</div>
        <div class="menu-desc">
           Looks at <b>Date 2</b> of the <i>previous row</i> and finds the next weekday (skips Sat/Sun).
        </div>
      </div>
      
       <div class="menu-item">
        <div class="menu-name">2. Date 2 (Column C)</div>
        <div class="menu-desc">
           Looks at <b>Date 1</b> of the <i>current row</i> and finds the next weekday.
        </div>
      </div>
      
       <div class="menu-item">
        <div class="menu-name">3. Block Type (Column D)</div>
        <div class="menu-desc">
           Automatically toggles between <b>A/B</b> and <b>C/D</b> based on the previous row.
        </div>
      </div>

       <div class="tip">
         <b>Why this matters:</b> If you change one date at the top, the rest of the year updates automatically!
       </div>
    </div>

    <div class="card">
       <div style="font-weight:bold; margin-bottom:10px;">🛑 When to Manual Override</div>
       <p>The "Chain" works great until it hits a holiday or break. You must <b>manually type</b> over the formulas when:</p>
       <ul style="padding-left: 20px; margin-top: 5px;">
         <li><b>Start of Year/Semester:</b> Type the first date manually to start the chain.</li>
         <li><b>Holidays/Breaks:</b> If a week is skipped, type the correct "Date 1" in the next row to jump the gap.</li>
         <li><b>Resets:</b> If the A/B cycle needs to reset, type the new block letter manually in Column D.</li>
       </ul>
       <p style="margin-top:10px; font-style:italic; font-size:12px;">
         <b>Note:</b> Once you manually type a date, you can drag the formulas from the cell <i>below</i> it back up to restart the automatic chain.
       </p>
    </div>

    <div class="card">
        <div style="font-weight:bold; margin-bottom:10px;">🚀 Quick Start</div>
        <div class="step"><div class="num">1</div><div><b>Template</b>: Use <b>Calendar Tools > Create Pacing Template</b> to get started.</div></div>
        <div class="step"><div class="num">2</div><div><b>Format</b>: Ensure columns match:
          <ul style="margin:5px 0 0 -20px;">
            <li><b>A</b>: Class Number</li>
            <li><b>B-C</b>: Start/End Dates</li>
            <li><b>D</b>: Block/Day Type (A, B, or Holiday)</li>
            <li><b>E</b>: Notes (Holidays)</li>
          </ul>
        </div></div>
        <div class="step"><div class="num">3</div><div><b>Run</b>: Select a view from the menu.</div></div>
    </div>

    <h3>📖 Menu Reference</h3>
    
    <div class="menu-item">
        <div class="menu-name">Create Wall Calendar View</div>
        <div class="menu-desc">Generates a traditional monthly calendar grid (July to June). Best for printing by month.</div>
    </div>
    <div class="menu-item">
        <div class="menu-name">Create Lateral Calendar View</div>
        <div class="menu-desc">Creates a compact linear view relative to the school year logic. Good for seeing the "flow" of the year.</div>
    </div>
    <div class="menu-item">
        <div class="menu-name">📄 Create Pacing Template</div>
        <div class="menu-desc">Creates a new sheet with smart formulas pre-filled for the current school year.</div>
    </div>
    <div class="menu-item">
        <div class="menu-name">⚙️ Configuration</div>
        <div class="menu-desc">Change the School Year Start Month, customize Holiday Keywords, and edit Calendar Colors.</div>
    </div>

    <div style="margin-top:20px; font-size:12px; color:#666; text-align:center; border-top: 1px solid #eee; padding-top: 15px;">
        <p style="margin-bottom:5px;">Developed by <a href="https://knuffke.com/support" target="_blank" style="color:#333; text-decoration:none;"><b>David Knuffke</b></a></p>
        <p style="font-size:10px; margin-top:5px;">Made available under a <a href="http://creativecommons.org/licenses/by-nc-sa/4.0/" target="_blank">CC BY-NC-SA 4.0 License</a>.</p>
        <a href="#" onclick="google.script.host.close()">Close Guide</a>
    </div>
  `;
}
