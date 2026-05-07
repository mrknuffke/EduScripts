function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Department Tools")
    .addItem("How to Use / Tutorial", "showTutorial")
    .addSeparator()
    .addItem("1. Department Setup Wizard", "openSetupWizard")
    .addItem("2. Create Initial Schedule Grid", "buildMasterSchedule")
    .addItem("3. Course Entry Editor", "courseEntryEditor")
    .addItem("4. Color and Format Courses", "colorAndFormatCourses")
    .addSeparator()
    .addItem("5. Generate C/D Days (Day Swapper)", "daySwapper")
    .addSeparator()
    .addItem("Format All Schedule Sheets", "colorAndFormatAllSchedules")
    .addItem("Refresh Courses from Schedule", "refreshConfigCourses")
    .addItem("Validate Config Sheet", "validateConfig")
    .addToUi();
}

function showTutorial() {
  var html = `
    <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; color: #333; line-height: 1.5; }
          h2 { color: #2E7D32; border-bottom: 2px solid #2E7D32; padding-bottom: 5px; }
          h3 { color: #1565C0; margin-top: 20px; }
          ol { padding-left: 20px; }
          li { margin-bottom: 10px; }
          .highlight { background-color: #FFF9C4; padding: 2px 4px; border-radius: 3px; }
        </style>
      </head>
      <body>
        <h2>Department Schedule Tools Tutorial</h2>
        <p>Welcome! This toolset helps you generate and format a master schedule from raw exported data.</p>
        
        <h3>Step 1: Setup Your Department</h3>
        <p>Run <b>Department Setup Wizard</b>. This will create a native <b>Department Config</b> tab in your spreadsheet, pre-filled with courses from your current sheet.</p>
        <p><i>You can type your preferred hex colors and abbreviations directly into this sheet at any time!</i></p>

        <h3>Step 2: Import Data & Build Master</h3>
        <ol>
          <li>Copy over your teacher info from the provided master sheet. <i>(Make sure to keep the top rows with the organizing information!)</i></li>
          <li>Make sure you are viewing that specific sheet.</li>
          <li>Run <b>Create Initial Schedule Grid</b>. This will reorganize the raw rows and generate <b>TWO new sheets</b>: one for Semester 1 and one for Semester 2, perfectly mapped out across A, B, C, and D days.</li>
        </ol>

        <h3>Step 3: Clean Up Text & Formatting</h3>
        <p>You can format your generated schedules using the remaining tools. <i>(Note: You will need to run these tools once while viewing the Semester 1 sheet, and again while viewing the Semester 2 sheet!)</i></p>
        <ul>
          <li><b>Course Entry Editor:</b> Select which lines of text to keep (e.g., stripping out block times if you don't need them) and automatically apply your chosen course abbreviations.</li>
          <li><b>Color and Format Courses:</b> Automatically apply the background colors, font colors, and text alignment you defined in the Department Config sheet.</li>
        </ul>
        
        <p style="margin-top: 30px; text-align: center;">
          <button style="padding: 8px 16px; cursor: pointer;" onclick="google.script.host.close()">Close</button>
        </p>
      </body>
    </html>
  `;
  var htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(500)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Tutorial & Help');
}

// ==========================================
// CONFIG MANAGEMENT (Native Sheet)
// ==========================================

const CONFIG_SHEET_NAME = "Department Config";
const DATA_START_ROW = 6;
const DATA_COL_START = "C";
const DATA_COL_END = "U";

function getConfigData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!sheet) {
    return {
      departmentName: "Default Department",
      fontFamily: "Comfortaa",
      horizontalAlignment: "center",
      verticalAlignment: "middle",
      wrapText: true,
      courses: {} // format: { "Full Name": { abbr: "Abbr", bg: "#hex", font: "#hex" } }
    };
  }
  
  var data = sheet.getDataRange().getValues();
  
  var config = {
    departmentName: "Default Department",
    fontFamily: "Comfortaa",
    horizontalAlignment: "center",
    verticalAlignment: "middle",
    wrapText: true,
    boldCourseNames: true,
    teacherBg: "#FFFFFF", teacherFont: "#000000",
    semesterBg: "#FFFFFF", semesterFont: "#000000",
    dayBg: "#000000", dayFont: "#FFFFFF",
    blockBg: "#FFFFFF", blockFont: "#000000",
    courses: {}
  };
  
  // Safe extraction with fallbacks
  for (var i = 0; i < data.length; i++) {
    var label1 = String(data[i][0]);
    var val1 = data[i][1];
    var label2 = String(data[i][2]);
    var val2 = data[i][3];
    
    if (label1.indexOf("Department Name") !== -1 && val1 !== "") config.departmentName = val1;
    if (label1.indexOf("Font Family") !== -1 && val1 !== "") config.fontFamily = val1;
    if (label1.indexOf("Horiz Align") !== -1 && val1 !== "") config.horizontalAlignment = val1;
    if (label1.indexOf("Vert Align") !== -1 && val1 !== "") config.verticalAlignment = val1;
    if (label1.indexOf("Wrap Text") !== -1 && val1 !== "") config.wrapText = (val1 === true || val1 === "true");
    if (label1.indexOf("Bold Course Names") !== -1 && val1 !== "") config.boldCourseNames = (val1 === true || val1 === "true");
    
    if (label1.indexOf("Teacher Names") !== -1 && val1 !== "") config.teacherBg = val1;
    if (label2.indexOf("Teacher Names Font") !== -1 && val2 !== "") config.teacherFont = val2;
    
    if (label1.indexOf("Semester Label") !== -1 && val1 !== "") config.semesterBg = val1;
    if (label2.indexOf("Semester Label Font") !== -1 && val2 !== "") config.semesterFont = val2;
    
    if (label1.indexOf("Day Headers") !== -1 && val1 !== "") config.dayBg = val1;
    if (label2.indexOf("Day Headers Font") !== -1 && val2 !== "") config.dayFont = val2;
    
    if (label1.indexOf("Block Headers") !== -1 && val1 !== "") config.blockBg = val1;
    if (label2.indexOf("Block Headers Font") !== -1 && val2 !== "") config.blockFont = val2;
  }
  
  var courseStartIndex = 7;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === "Course Title") {
      courseStartIndex = i + 1;
      break;
    }
  }
  
  for (var i = courseStartIndex; i < data.length; i++) {
    var title = data[i][0];
    if (title) {
      config.courses[title] = {
        abbr: data[i][1] || title,
        bg: data[i][2] || "#D9D9D9",
        font: data[i][3] || "#000000"
      };
    }
  }
  return config;
}

function saveConfigData(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  }
  
  var output = [
    ["Department Name:", config.departmentName, "", ""],
    ["Font Family:", config.fontFamily, "", ""],
    ["Horiz Align:", config.horizontalAlignment, "", ""],
    ["Vert Align:", config.verticalAlignment, "", ""],
    ["Wrap Text:", config.wrapText, "", ""],
    ["Bold Course Names:", config.boldCourseNames, "", ""],
    ["Teacher Names Background:", config.teacherBg, "Teacher Names Font:", config.teacherFont],
    ["Semester Label Background:", config.semesterBg, "Semester Label Font:", config.semesterFont],
    ["Day Headers Background:", config.dayBg, "Day Headers Font:", config.dayFont],
    ["Block Headers Background:", config.blockBg, "Block Headers Font:", config.blockFont],
    ["", "", "", ""],
    ["Course Title", "Abbreviation", "Background Hex", "Font Hex"]
  ];
  
  for (var title in config.courses) {
    var c = config.courses[title];
    output.push([title, c.abbr, c.bg, c.font]);
  }
  
  sheet.getRange(1, 1, output.length, 4).setValues(output);
  sheet.getRange(output.length > 8 ? 8 : output.length, 1, 1, 4).setFontWeight("bold");
}

function getUniqueCoursesFromActiveSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var courses = new Set();
  
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (j === 0) continue; // Skip column A entirely (Teacher names)
      if (i < 2) continue; // Skip top header rows
      
      var cell = String(data[i][j]).trim();
      if (cell) {
        var firstLine = cell.split('\n')[0].trim();
        var lookupTitle = firstLine.replace(/[\*\u200B-\u200D\uFEFF]/g, '').trim();
        var lowerTitle = lookupTitle.toLowerCase();
        
        var isBlock = lowerTitle.indexOf("block") !== -1;
        var isSemester = lowerTitle.indexOf("semester") !== -1;
        var isDay = lowerTitle === "a day" || lowerTitle === "b day" || lowerTitle === "c day" || lowerTitle === "d day" || lowerTitle === "a" || lowerTitle === "b";
        var isAdvisory = lowerTitle.indexOf("advisory") !== -1;
        var isTeacherHeader = lowerTitle === "teacher";
        
        // Better heuristic to skip obvious non-course cells
        if (lookupTitle && lookupTitle.length > 2 && !isBlock && !isSemester && !isDay && !isAdvisory && !isTeacherHeader) {
             courses.add(lookupTitle);
        }
      }
    }
  }
  return Array.from(courses).sort();
}

// Sample course template (Science department). Used as placeholder data when
// no courses can be scanned from the active sheet during initial setup.
// Replace or extend this object to provide defaults for other departments.
const SCIENCE_DEFAULTS = {
  "Zoology": { bg: "#D9EAD3", font: "#000000" },
  "Marine Biology": { bg: "#D9EAD3", font: "#000000" },
  "Biotechnology": { abbr: "Biotech", bg: "#D9EAD3", font: "#000000" },
  "Physiology": { bg: "#D9EAD3", font: "#000000" },
  "Forensic Science": { abbr: "Forensics", bg: "#FAEFEF", font: "#000000" },
  "Physical Science": { bg: "#E7F2FC", font: "#000000" },
  "Chemistry": { bg: "#EA9999", font: "#000000" },
  "Accelerated Chemistry": { abbr: "Accel.Chem", bg: "#E06666", font: "#000000" },
  "AP Chemistry": { bg: "#990000", font: "#FFFFFF" },
  "Physics": { bg: "#9FC5E8", font: "#000000" },
  "Accelerated Physics": { abbr: "Accel.Phys", bg: "#6D9EEB", font: "#000000" },
  "AP Physics": { bg: "#3D78D8", font: "#FFFFFF" },
  "AP Physics C": { bg: "#3D78D8", font: "#FFFFFF" },
  "AP Physics 2": { bg: "#3D78D8", font: "#FFFFFF" },
  "Biology": { bg: "#93C47D", font: "#000000" },
  "Accelerated Biology": { abbr: "Accel.Bio", bg: "#6AA850", font: "#000000" },
  "AP Biology": { bg: "#38761D", font: "#FFFFFF" },
  "Environmental Science": { bg: "#FFF2CC", font: "#000000" },
  "AT Enviro Sci & Field Research": { abbr: "ATES", bg: "#FFD966", font: "#000000" },
  "AT Science: Modeling & Simulation": { abbr: "AT Sci.Mod.", bg: "#E7F2FC", font: "#000000" },
  "Robotics Science": { abbr: "Robotic Sci.", bg: "#D9D9D9", font: "#000000" }
};

// ==========================================
// NATIVE SHEET SETUP (Formerly Wizard)
// ==========================================

function openSetupWizard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  // If the config sheet already exists, just activate it and warn them
  if (sheet) {
    sheet.activate();
    SpreadsheetApp.getUi().alert("Setup Complete", "The 'Department Config' sheet already exists. You can edit your colors and abbreviations directly on this sheet.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Otherwise, create it and populate with defaults and scanned courses
  var scannedCourses = getUniqueCoursesFromActiveSheet();
  
  var newConfig = {
    departmentName: "My Department",
    fontFamily: "Comfortaa",
    horizontalAlignment: "center",
    verticalAlignment: "middle",
    wrapText: true,
    courses: {}
  };
  
  var handled = new Set();
  
  // 1. Add scanned courses with Science defaults if matched
  for (var i=0; i<scannedCourses.length; i++) {
    var c = scannedCourses[i];
    handled.add(c);
    
    if (SCIENCE_DEFAULTS[c]) {
      var sd = SCIENCE_DEFAULTS[c];
      newConfig.courses[c] = { abbr: sd.abbr || c, bg: sd.bg, font: sd.font };
    } else {
      newConfig.courses[c] = { abbr: c, bg: '#D9D9D9', font: '#000000' };
    }
  }
  
  // 2. If nothing was scanned, dump the whole Science template so they have examples
  if (scannedCourses.length === 0) {
    for (var title in SCIENCE_DEFAULTS) {
      var sd = SCIENCE_DEFAULTS[title];
      newConfig.courses[title] = { abbr: sd.abbr || title, bg: sd.bg, font: sd.font };
    }
  }
  
  // Save to build the sheet
  saveConfigData(newConfig);
  
  // Format the new sheet nicely
  sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setFrozenRows(7);
  
  sheet.activate();
  SpreadsheetApp.getUi().alert("Setup Initialized", "A new 'Department Config' sheet has been created.\n\nPlease type your preferred hex color codes and abbreviations directly into this sheet. The tools will automatically read from it!", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================
// COURSE ENTRY EDITOR
// ==========================================

function courseEntryEditor() {
  var html = `
    <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 15px; color: #333; }
          .option { margin-bottom: 12px; }
          .header { font-weight: bold; margin-bottom: 15px; font-size: 1.1em; }
          .btn { background-color: #4CAF50; color: white; padding: 8px 12px; border: none; border-radius: 4px; cursor: pointer; margin-right: 8px; font-size: 14px; }
          .btn-cancel { background-color: #f44336; }
          .btn:hover { opacity: 0.9; }
          label { cursor: pointer; }
        </style>
      </head>
      <body>
        <div class="header">Select information to keep:</div>
        <div class="option"><label><input type="checkbox" id="keepLine1"> Line 2 (Section & Students)</label></div>
        <div class="option"><label><input type="checkbox" id="keepLine2" checked> Line 3 (Room)</label></div>
        <div class="option"><label><input type="checkbox" id="keepLine3"> Line 4 (Block/Day)</label></div>
        <div class="option"><label><input type="checkbox" id="keepLine4" checked> Line 5 (Term)</label></div>
        <hr style="margin: 15px 0;">
        <div class="option"><label><input type="checkbox" id="applyAll"> Apply to all schedule sheets</label></div>
        <div style="margin-top: 20px;">
          <button class="btn" onclick="apply()">Apply</button>
          <button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>
        </div>
        <script>
          function apply() {
            var options = {
              keepLine1: document.getElementById('keepLine1').checked,
              keepLine2: document.getElementById('keepLine2').checked,
              keepLine3: document.getElementById('keepLine3').checked,
              keepLine4: document.getElementById('keepLine4').checked,
              applyAll: document.getElementById('applyAll').checked
            };
            document.querySelectorAll('button').forEach(b => b.disabled = true);
            document.querySelector('.header').innerText = 'Processing... Please wait.';
            google.script.run.withSuccessHandler(function() {
              google.script.host.close();
            }).processCourseEntries(options);
          }
        </script>
      </body>
    </html>
  `;
  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(320).setHeight(310);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Course Entry Editor');
}

function processCourseEntries(options) {
  if (options.applyAll) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets().filter(function(s) {
      return s.getName().indexOf("Schedule") !== -1 && s.getName() !== CONFIG_SHEET_NAME;
    });
    if (sheets.length === 0) {
      SpreadsheetApp.getUi().alert("No schedule sheets found. Run 'Create Initial Schedule Grid' first.");
      return;
    }
    var totalProcessed = 0;
    for (var k = 0; k < sheets.length; k++) {
      totalProcessed += processCourseEntriesOnSheet(sheets[k], options);
    }
    ss.toast("Processed " + totalProcessed + " cells across " + sheets.length + " schedule sheets.", "Success", 5);
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("Please switch to a schedule sheet before running this tool.");
    return;
  }
  var count = processCourseEntriesOnSheet(sheet, options);
  SpreadsheetApp.getActiveSpreadsheet().toast("Processed " + count + " course cells on \"" + sheet.getName() + "\".", "Success", 4);
}

function processCourseEntriesOnSheet(sheet, options) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return 0;
  
  var range = sheet.getRange(DATA_COL_START + DATA_START_ROW + ":" + DATA_COL_END + lastRow);
  var data = range.getValues();
  var richTextValues = range.getRichTextValues();

  var config = getConfigData();
  var processed = 0;

  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      var cellText = data[i][j].trim();
      if (!cellText) continue;

      var lines = cellText.split('\n');

      if (lines.length > 0) {
        var rawTitle = lines[0].trim();
        var lookupTitle = rawTitle.replace(/[\*\u200B-\u200D\uFEFF]/g, '').trim();
        
        if (config.courses[lookupTitle] && config.courses[lookupTitle].abbr) {
          lines[0] = config.courses[lookupTitle].abbr;
        } else {
          lines[0] = rawTitle;
        }
      }

      var newLines = [lines[0]];
      var roomStartIndex = -1;
      var roomLength = 0;

      if (lines.length > 1 && options.keepLine1) newLines.push(lines[1]);
      if (lines.length > 2 && options.keepLine2) {
        roomStartIndex = newLines.join('\n').length + 1;
        roomLength = lines[2].length;
        newLines.push(lines[2]);
      }
      if (lines.length > 3 && options.keepLine3) newLines.push(lines[3]);
      if (lines.length > 4 && options.keepLine4) newLines.push(lines[4]);

      var textValueBuilder = SpreadsheetApp.newRichTextValue().setText(newLines.join('\n'));
      textValueBuilder = textValueBuilder.setTextStyle(0, lines[0].length, SpreadsheetApp.newTextStyle().setBold(true).build());

      if (roomStartIndex !== -1) {
        textValueBuilder = textValueBuilder.setTextStyle(roomStartIndex, roomStartIndex + roomLength, SpreadsheetApp.newTextStyle().setBold(true).build());
      }

      richTextValues[i][j] = textValueBuilder.build();
      processed++;
    }
  }
  
  range.setRichTextValues(richTextValues);
  return processed;
}

// ==========================================
// COLOR AND FORMAT COURSES
// ==========================================

function colorAndFormatCourses() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("Please switch to a schedule sheet before running this tool.");
    return;
  }
  var config = getConfigData();
  var result = colorAndFormatSheet(sheet, config);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Formatted " + result.formatted + " cells (" + result.unmatched + " unmatched) using " + config.departmentName + " settings!",
    "Success", 5
  );
}

function colorAndFormatAllSchedules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfigData();
  var sheets = ss.getSheets().filter(function(s) {
    return s.getName().indexOf("Schedule") !== -1 && s.getName() !== CONFIG_SHEET_NAME;
  });
  if (sheets.length === 0) {
    SpreadsheetApp.getUi().alert("No schedule sheets found. Run 'Create Initial Schedule Grid' first.");
    return;
  }
  var totalFormatted = 0, totalUnmatched = 0;
  for (var i = 0; i < sheets.length; i++) {
    var result = colorAndFormatSheet(sheets[i], config);
    totalFormatted += result.formatted;
    totalUnmatched += result.unmatched;
  }
  ss.toast(
    "Formatted " + totalFormatted + " cells across " + sheets.length + " sheets (" + totalUnmatched + " unmatched).",
    "Success", 5
  );
}

function colorAndFormatSheet(sheet, config) {
  var lastRow = sheet.getLastRow();
  var result = { formatted: 0, unmatched: 0 };
  if (lastRow < DATA_START_ROW) return result;

  var range = sheet.getRange(DATA_COL_START + DATA_START_ROW + ":" + DATA_COL_END + lastRow);
  var data = range.getValues();

  var defaultStyle = {
    background: "#D9D9D9",
    font: "#000000"
  };

  var backgrounds = range.getBackgrounds();
  var fontColors = range.getFontColors();
  var horizontalAlignments = range.getHorizontalAlignments();
  var verticalAlignments = range.getVerticalAlignments();
  var wraps = range.getWraps();
  var fontFamilies = range.getFontFamilies();
  var richTextValues = range.getRichTextValues();

  // Create reverse mapping from abbreviation to title, to support formatting after abbreviations applied
  var abbrToTitle = {};
  for(var t in config.courses) {
      if(config.courses[t].abbr) {
          abbrToTitle[config.courses[t].abbr] = t;
      }
  }

  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      var cellContent = data[i][j].trim();

      if (cellContent) {
        var courseTitle = cellContent.split('\n')[0].trim();
        var lookupTitle = courseTitle.replace(/[\*\u200B-\u200D\uFEFF]/g, '').trim();

        // Check if config has it. If not, maybe it's already an abbreviation?
        var cData = config.courses[lookupTitle];
        if (!cData && abbrToTitle[lookupTitle]) {
           cData = config.courses[abbrToTitle[lookupTitle]];
        }

        if (cData) {
          backgrounds[i][j] = cData.bg || defaultStyle.background;
          fontColors[i][j] = cData.font || defaultStyle.font;
          result.formatted++;
        } else {
          backgrounds[i][j] = defaultStyle.background;
          fontColors[i][j] = defaultStyle.font;
          result.unmatched++;
        }

        horizontalAlignments[i][j] = config.horizontalAlignment;
        verticalAlignments[i][j] = config.verticalAlignment;
        wraps[i][j] = config.wrapText;
        fontFamilies[i][j] = config.fontFamily;
        
        // Build rich text to apply formatting to the first line (bolding and coloring)
        var fontColor = fontColors[i][j];
        var baseStyle = SpreadsheetApp.newTextStyle()
            .setForegroundColor(fontColor)
            .setFontFamily(config.fontFamily)
            .build();
            
        var titleStyle = SpreadsheetApp.newTextStyle()
            .setForegroundColor(fontColor)
            .setFontFamily(config.fontFamily)
            .setBold(config.boldCourseNames)
            .build();
            
        var builder = SpreadsheetApp.newRichTextValue()
            .setText(cellContent)
            .setTextStyle(0, cellContent.length, baseStyle);
            
        if (courseTitle.length > 0) {
            builder.setTextStyle(0, courseTitle.length, titleStyle);
        }
        
        richTextValues[i][j] = builder.build();
      }
    }
  }

  range.setBackgrounds(backgrounds);
  range.setHorizontalAlignments(horizontalAlignments);
  range.setVerticalAlignments(verticalAlignments);
  range.setWraps(wraps);
  // We apply the rich text values to enforce the bolding and font colors without conflict
  range.setRichTextValues(richTextValues);
  
  // Format Structural Elements
  sheet.getRange("A2:U2").setBackground(config.semesterBg).setFontColor(config.semesterFont).setFontFamily(config.fontFamily);
  sheet.getRange("A4:U4").setBackground(config.dayBg).setFontColor(config.dayFont).setFontFamily(config.fontFamily);
  sheet.getRange("A5:U5").setBackground(config.blockBg).setFontColor(config.blockFont).setFontFamily(config.fontFamily);
  if (lastRow >= DATA_START_ROW) {
    sheet.getRange("A" + DATA_START_ROW + ":A" + lastRow).setBackground(config.teacherBg).setFontColor(config.teacherFont).setFontFamily(config.fontFamily);
  }
  
  return result;
}

// ==========================================
// BUILD MASTER SCHEDULE
// ==========================================

function buildMasterSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  if (data.length < 3 || data[1][0] !== "Teacher") {
    SpreadsheetApp.getUi().alert("Please run this tool while viewing the sheet that contains your pasted CSV data.");
    return;
  }
  
  // Auto-detect S1 and S2 columns based on Row 0 headers
  var headerRow = data[0];
  var s1Start = -1;
  var s2Start = -1;
  
  for (var c = 2; c < headerRow.length; c++) {
    var txt = String(headerRow[c]).toLowerCase();
    if (txt.indexOf("semester 1") !== -1 && s1Start === -1) s1Start = c;
    if (txt.indexOf("semester 2") !== -1 && s2Start === -1) s2Start = c;
  }
  
  // Fallbacks if headers changed
  if (s1Start === -1) s1Start = 2;
  if (s2Start === -1) s2Start = (headerRow.length > 7 && String(headerRow[6]) === "") ? 7 : 6;
  
  var scheduleData = {};
  var teachers = [];
  
  for (var i = 2; i < data.length; i++) {
    var row = data[i];
    var rawTeacher = row[0];
    if (!rawTeacher) continue;
    
    var teacher = String(rawTeacher).split(" (")[0].trim();
    var day = String(row[1]).trim();
    
    if (teachers.indexOf(teacher) === -1) {
      teachers.push(teacher);
      scheduleData[teacher] = {
        S1: { "A": ["", "", "", ""], "B": ["", "", "", ""], "C": ["", "", "", ""], "D": ["", "", "", ""] },
        S2: { "A": ["", "", "", ""], "B": ["", "", "", ""], "C": ["", "", "", ""], "D": ["", "", "", ""] }
      };
    }
    
    if (day === "A" || day === "B" || day === "C" || day === "D") {
      scheduleData[teacher].S1[day] = [
        row[s1Start] || "", row[s1Start+1] || "", row[s1Start+2] || "", row[s1Start+3] || ""
      ];
      // Prevent out of bounds if rows are truncated
      scheduleData[teacher].S2[day] = [
        row[s2Start] || "", row[s2Start+1] || "", row[s2Start+2] || "", row[s2Start+3] || ""
      ];
    }
  }
  
  teachers.sort();
  
  var config = getConfigData();
  
  function buildSemesterSheet(semesterName, semKey) {
    var newSheetName = config.departmentName + " " + semesterName + " Schedule";
    var newSheet = ss.getSheetByName(newSheetName);
    if (!newSheet) {
      newSheet = ss.insertSheet(newSheetName);
    } else {
      newSheet.clear();
    }
    
    var emptyRow = ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""];
    var finalOutput = [];
    finalOutput.push(emptyRow); 
    
    var semRow = ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""];
    semRow[2] = config.departmentName + " " + semesterName;
    finalOutput.push(semRow); 
    
    finalOutput.push(emptyRow); 
    
    finalOutput.push([
      "", "", 
      "A Day", "", "", "", "", 
      "B Day", "", "", "", "", 
      "C Day", "", "", "", "", 
      "D Day", "", "", ""
    ]);
    
    finalOutput.push([
      "Teacher", "", 
      "Block 1", "Block 2", "Block 3", "Block 4", "", 
      "Block 1", "Block 2", "Block 3", "Block 4", "", 
      "Block 1", "Block 2", "Block 3", "Block 4", "", 
      "Block 1", "Block 2", "Block 3", "Block 4"
    ]);
    
    for (var i = 0; i < teachers.length; i++) {
      var t = teachers[i];
      var dA = scheduleData[t][semKey]["A"];
      var dB = scheduleData[t][semKey]["B"];
      var dC = scheduleData[t][semKey]["C"];
      var dD = scheduleData[t][semKey]["D"];
      
      finalOutput.push([
        t, "", 
        dA[0], dA[1], dA[2], dA[3], "", 
        dB[0], dB[1], dB[2], dB[3], "", 
        dC[0], dC[1], dC[2], dC[3], "", 
        dD[0], dD[1], dD[2], dD[3]
      ]);
    }
    
    newSheet.getRange(1, 1, finalOutput.length, finalOutput[0].length).setValues(finalOutput);
    
    newSheet.getRange(2, 3, 1, 19).merge();
    newSheet.getRange(2, 3).setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
    
    newSheet.getRange(4, 1, 2, finalOutput[0].length).setFontWeight("bold").setHorizontalAlignment("center");
    newSheet.getRange(4, 3, 1, 4).merge();  // A Day
    newSheet.getRange(4, 8, 1, 4).merge();  // B Day
    newSheet.getRange(4, 13, 1, 4).merge(); // C Day
    newSheet.getRange(4, 18, 1, 4).merge(); // D Day
    
    newSheet.getRange(1, 1, finalOutput.length, finalOutput[0].length).setVerticalAlignment("middle");
    
    newSheet.setColumnWidth(1, 180); 
    newSheet.setColumnWidth(2, 20);  
    newSheet.setColumnWidth(7, 20);  
    newSheet.setColumnWidth(12, 20); 
    newSheet.setColumnWidth(17, 20); 
    
    var blockCols = [3, 4, 5, 6, 8, 9, 10, 11, 13, 14, 15, 16, 18, 19, 20, 21];
    for (var c = 0; c < blockCols.length; c++) {
      var colIndex = blockCols[c];
      newSheet.setColumnWidth(colIndex, 150);
      newSheet.getRange(6, colIndex, teachers.length, 1).setWrap(true);
    }
    
    newSheet.setFrozenRows(5);
    newSheet.setFrozenColumns(1);
  }
  
  buildSemesterSheet("Semester 1", "S1");
  buildSemesterSheet("Semester 2", "S2");
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Master Schedules built successfully!", "Success", 5);
}

// ==========================================
// DAY SWAPPER (GENERATE C & D DAYS)
// ==========================================

function daySwapper() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() === CONFIG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("Please switch to a schedule sheet before running this tool.");
    return;
  }
  // A Day to C Day
  copyColumns(sheet, "C", "D", "O", "P", DATA_START_ROW); // A1, A2 -> C3, C4
  copyColumns(sheet, "E", "F", "M", "N", DATA_START_ROW); // A3, A4 -> C1, C2
  // B Day to D Day
  copyColumns(sheet, "H", "I", "T", "U", DATA_START_ROW); // B1, B2 -> D3, D4
  copyColumns(sheet, "J", "K", "R", "S", DATA_START_ROW); // B3, B4 -> D1, D2
  
  SpreadsheetApp.getActiveSpreadsheet().toast("C and D days successfully generated from A and B days!", "Success", 4);
}

function copyColumns(sheet, srcCol1, srcCol2, destCol1, destCol2, startRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;
  
  var srcRange1 = sheet.getRange(srcCol1 + startRow + ":" + srcCol1 + lastRow);
  var srcRange2 = sheet.getRange(srcCol2 + startRow + ":" + srcCol2 + lastRow);
  var destRange1 = sheet.getRange(destCol1 + startRow + ":" + destCol1 + lastRow);
  var destRange2 = sheet.getRange(destCol2 + startRow + ":" + destCol2 + lastRow);

  destRange1.setRichTextValues(srcRange1.getRichTextValues());
  destRange2.setRichTextValues(srcRange2.getRichTextValues());
  destRange1.setBackgrounds(srcRange1.getBackgrounds());
  destRange2.setBackgrounds(srcRange2.getBackgrounds());
  destRange1.setFontColors(srcRange1.getFontColors());
  destRange2.setFontColors(srcRange2.getFontColors());
  destRange1.setHorizontalAlignments(srcRange1.getHorizontalAlignments());
  destRange1.setVerticalAlignments(srcRange1.getVerticalAlignments());
  destRange2.setHorizontalAlignments(srcRange2.getHorizontalAlignments());
  destRange2.setVerticalAlignments(srcRange2.getVerticalAlignments());
  destRange1.setWraps(srcRange1.getWraps());
  destRange2.setWraps(srcRange2.getWraps());
  destRange1.setFontFamilies(srcRange1.getFontFamilies());
  destRange2.setFontFamilies(srcRange2.getFontFamilies());
}

// ==========================================
// REFRESH COURSES FROM SCHEDULE
// ==========================================

function refreshConfigCourses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!configSheet) {
    SpreadsheetApp.getUi().alert("No config sheet found. Run 'Department Setup Wizard' first.");
    return;
  }
  
  var config = getConfigData();
  var existingTitles = new Set(Object.keys(config.courses));
  
  // Build a set of abbreviations so we can also recognize already-abbreviated names
  var existingAbbrs = new Set();
  for (var t in config.courses) {
    if (config.courses[t].abbr) existingAbbrs.add(config.courses[t].abbr);
  }
  
  // Scan all schedule sheets for course names
  var sheets = ss.getSheets().filter(function(s) {
    return s.getName().indexOf("Schedule") !== -1 && s.getName() !== CONFIG_SHEET_NAME;
  });
  
  var newCourses = [];
  for (var s = 0; s < sheets.length; s++) {
    var data = sheets[s].getDataRange().getValues();
    for (var i = DATA_START_ROW - 1; i < data.length; i++) {
      for (var j = 2; j < data[i].length; j++) {
        var cell = String(data[i][j]).trim();
        if (!cell) continue;
        var firstLine = cell.split('\n')[0].trim();
        var cleanTitle = firstLine.replace(/[\*\u200B-\u200D\uFEFF]/g, '').trim();
        if (cleanTitle && cleanTitle.length > 2 && !existingTitles.has(cleanTitle) && !existingAbbrs.has(cleanTitle)) {
          var lower = cleanTitle.toLowerCase();
          if (lower.indexOf("block") === -1 && lower.indexOf("semester") === -1 && lower.indexOf("advisory") === -1 &&
              lower !== "a day" && lower !== "b day" && lower !== "c day" && lower !== "d day" &&
              lower !== "teacher" && lower !== "a" && lower !== "b") {
            newCourses.push(cleanTitle);
            existingTitles.add(cleanTitle);
          }
        }
      }
    }
  }
  
  if (newCourses.length === 0) {
    SpreadsheetApp.getUi().alert("All courses in your schedule sheets are already in the config. Nothing to add!");
    return;
  }
  
  newCourses.sort();
  var lastRow = configSheet.getLastRow();
  var output = [];
  for (var i = 0; i < newCourses.length; i++) {
    output.push([newCourses[i], newCourses[i], "#D9D9D9", "#000000"]);
  }
  configSheet.getRange(lastRow + 1, 1, output.length, 4).setValues(output);
  
  configSheet.activate();
  SpreadsheetApp.getActiveSpreadsheet().toast("Added " + newCourses.length + " new course(s) to config. Don't forget to set their colors and abbreviations!", "Courses Refreshed", 6);
}

// ==========================================
// VALIDATE CONFIG SHEET
// ==========================================

function validateConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!configSheet) {
    SpreadsheetApp.getUi().alert("No config sheet found. Run 'Department Setup Wizard' first.");
    return;
  }
  
  var config = getConfigData();
  var warnings = [];
  
  // Collect all course names from schedule sheets for cross-reference
  var scheduleCourses = new Set();
  var sheets = ss.getSheets().filter(function(s) {
    return s.getName().indexOf("Schedule") !== -1 && s.getName() !== CONFIG_SHEET_NAME;
  });
  for (var s = 0; s < sheets.length; s++) {
    var data = sheets[s].getDataRange().getValues();
    for (var i = DATA_START_ROW - 1; i < data.length; i++) {
      for (var j = 2; j < data[i].length; j++) {
        var cell = String(data[i][j]).trim();
        if (cell) {
          var firstLine = cell.split('\n')[0].trim().replace(/[\*\u200B-\u200D\uFEFF]/g, '').trim();
          if (firstLine) scheduleCourses.add(firstLine);
        }
      }
    }
  }
  
  // Validate each config course entry
  var hexPattern = /^#[0-9A-Fa-f]{6}$/;
  var configData = configSheet.getDataRange().getValues();
  var courseStartIndex = -1;
  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0] === "Course Title") {
      courseStartIndex = i + 1;
      break;
    }
  }
  
  if (courseStartIndex !== -1) {
    for (var i = courseStartIndex; i < configData.length; i++) {
      var rowNum = i + 1;
      var title = configData[i][0];
      var bg = String(configData[i][2]).trim();
      var font = String(configData[i][3]).trim();
      
      if (!title) continue;
      
      // Check hex codes
      if (bg && !hexPattern.test(bg)) {
        warnings.push("Row " + rowNum + ": Background \"" + bg + "\" is not a valid hex color (e.g. #FF0000).");
      }
      if (font && !hexPattern.test(font)) {
        warnings.push("Row " + rowNum + ": Font color \"" + font + "\" is not a valid hex color (e.g. #000000).");
      }
      
      // Check if course exists in any schedule (by title or abbreviation)
      var abbr = configData[i][1];
      if (!scheduleCourses.has(title) && !scheduleCourses.has(abbr)) {
        warnings.push("Row " + rowNum + ": \"" + title + "\" not found in any schedule sheet. It may be unused or misspelled.");
      }
    }
  }
  
  // Also check structural color fields
  for (var i = 0; i < configData.length; i++) {
    if (courseStartIndex !== -1 && i >= courseStartIndex) break;
    for (var j = 1; j < configData[i].length; j += 2) {
      var val = String(configData[i][j]).trim();
      if (val && val.indexOf("#") === 0 && !hexPattern.test(val)) {
        warnings.push("Row " + (i + 1) + ", Col " + (j + 1) + ": \"" + val + "\" is not a valid hex color.");
      }
    }
  }
  
  if (warnings.length === 0) {
    SpreadsheetApp.getUi().alert("Config Validation", "✅ Everything looks good! No issues found.", SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    var msg = "Found " + warnings.length + " potential issue(s):\n\n" + warnings.join("\n\n");
    SpreadsheetApp.getUi().alert("Config Validation", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
