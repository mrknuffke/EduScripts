// ===========================================================================
//  Gradebook roster-scanning fixtures.
//  Each fixture is {values, backgrounds, fontColors} for one sheet layout.
//  BLACK marks a divider banner cell (solid fill + white text).
// ===========================================================================
var WHITE = '#ffffff', BLACK = '#000000';

function sheetFrom(values, bannerRows) {
  var width = values.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
  var bg = values.map(function (row, r) {
    return Array.apply(null, { length: width }).map(function () {
      return bannerRows.indexOf(r) > -1 ? BLACK : WHITE;
    });
  });
  var fc = values.map(function (row, r) {
    return Array.apply(null, { length: width }).map(function () {
      return bannerRows.indexOf(r) > -1 ? '#ffffff' : '#000000';
    });
  });
  return {
    getDataRange: function () {
      return { getDisplayValues: function () { return values; } };
    },
    getRange: function (row, col, numRows) {
      var c = col - 1;
      return {
        getBackgrounds: function () {
          var o = []; for (var r = row - 1; r < row - 1 + numRows; r++) o.push([bg[r] ? bg[r][c] : WHITE]); return o;
        },
        getFontColors: function () {
          var o = []; for (var r = row - 1; r < row - 1 + numRows; r++) o.push([fc[r] ? fc[r][c] : '#000000']); return o;
        }
      };
    },
    getLastRow: function () { return values.length; }
  };
}

// --- 1. AP Bio Sem 1: checkbox col A, roster labels row 2, black dividers
//        in the Name column that also carry a headcount.
var AP_BIO = sheetFrom([
  ["", "Category:", "Admin", "", "Completion", "", "", "Topic Questions"],
  ["Initial Meeting", "Name:", "Preferred Name:", "Email:", "AC: Complete Pre-Test", "AC: Water Stations", "ID: 1.2 Water", "Est.Count"],
  ["", "", "", "", "", "", "", ""],
  ["FALSE", "AP Bio 1 (B2/D1)", "", "", "", "", "", "16"],
  ["TRUE",  "Lee, Isaiah",  "Isaiah", "lee781513@sas.edu.sg", "1", "1", "1", "1"],
  ["FALSE", "Lee, Kayla",   "Kayla",  "lee46496@sas.edu.sg",  "1", "1", "1", "1"],
  ["TRUE",  "Mahajan, Aarav", "Aarav", "mahajan4681@sas.edu.sg", "m", "i", "1", "1"],
  ["",      "Average/Sum:", "", "", "17", "20", "12", "0"],
  ["FALSE", "AP Bio 3 (A4/C2)", "", "", "", "", "", "18"],
  ["TRUE",  "Abdul Wahid, Abdul Muhsin", "Muhsin", "abdulwahid808236@sas.edu.sg", "1", "1", "1", "1"],
  ["FALSE", "Bhattacharyya, Aarna", "", "bhattachar771987@sas.edu.sg", "1", "1", "1", "1"],
  ["FALSE", "Zhang, Shuhan", "Shuhan", "zhang441826@sas.edu.sg", "1", "i", "1", "1"],
  ["",      "Average/Sum:", "", "", "15", "18", "14", "0"]
], [3, 8]);

// --- 2. Chem Sem 1: assignment names in row 2, roster labels in row 4,
//        checkbox completion columns, merged divider label in column A.
var CHEM = sheetFrom([
  ["", "", "", "Admin", "Formative Work", "", "", "i's and m's", "Completion Percentage"],
  ["Initial Appoinment?", "", "", "Assigment", "AC: Journal Submitted", "AC: 1.3- The Table Has a Pattern", "AC: 1.5 Reactivity & Bonding", "", ""],
  ["", "", "", "Standards", "", "", "", "", ""],
  ["", "Name", "Preferred Name", "Email", "Date", "", "", "", ""],
  ["Chem 2 (A2/C4)", "", "", "23", "", "", "", "", ""],
  ["TRUE",  "Bong, Abigail", "", "bong49520@sas.edu.sg", "TRUE", "TRUE", "FALSE", "1", "67"],
  ["FALSE", "Chen, Ellie", "", "chen804009@sas.edu.sg", "TRUE", "TRUE", "FALSE", "1", "67"],
  ["TRUE",  "Guilfoile, Luca", "", "guilfoile47663@sas.edu.sg", "TRUE", "TRUE", "TRUE", "0", "100"],
  ["FALSE", "He, Zhuting", "Henry", "he48640@sas.edu.sg", "FALSE", "FALSE", "FALSE", "3", "0"],
  ["FALSE", "Yu, Chengxuan", "", "yu48719@sas.edu.sg", "TRUE", "FALSE", "FALSE", "2", "33"],
  ["", "", "", "", "", "", "", "", ""],
  ["", "Average:", "", "", "", "", "", "", ""]
], [4]);

// --- 2b. Chem Sem 1 after the leading checkbox column was removed: roster is
//         now A/B/C, assignments start at D, banner label merged in column A.
var CHEM_V2 = sheetFrom([
  ["", "", "Admin", "Formative Work", "", "", "i's and m's", "Completion Percentage"],
  ["", "", "", "AC: Initial Appointment", "AC: Journal Submitted", "AC: 1.3- The Table Has a Pattern", "", ""],
  ["", "", "", "DCI.1 Structure & Properties of Matter", "", "", "", ""],
  ["Name", "Preferred Name", "Email", "", "", "", "", ""],
  ["A2/C4", "", "23", "", "", "", "", ""],
  ["Bong, Abigail", "", "bong49520@sas.edu.sg", "TRUE", "TRUE", "TRUE", "1", "67"],
  ["He, Yanting", "Eric", "he48639@sas.edu.sg", "TRUE", "TRUE", "TRUE", "1", "67"],
  ["He, Zhuting", "Henry", "he48640@sas.edu.sg", "FALSE", "FALSE", "FALSE", "3", "0"],
  ["Yu, Chengxuan", "", "yu48719@sas.edu.sg", "FALSE", "TRUE", "FALSE", "2", "33"],
  ["", "", "", "", "", "", "", ""],
  ["Average:", "", "", "", "", "", "", ""]
], [4]);

// --- 3. The demo gradebook this script generates: Section value per row.
var DEMO = sheetFrom([
  ["", "", "", "", "Classwork", "Classwork", "Homework", "Assessments"],
  ["Section", "Name", "Email", "Parent Email", "Assignment 1", "Assignment 2", "Assignment 3", "Summative Exam"],
  ["", "", "", "", "Standard 1", "Standard 1", "Standard 2", "Standard 3"],
  ["Block 1", "Potter, Harry", "harry@hogwarts.edu", "james.potter@hogwarts.edu", "1", "1", "0", "95"],
  ["Block 1", "Granger, Hermione", "hermione@hogwarts.edu", "mr.granger@londondentist.com", "1", "1", "1", "100"],
  ["Block 2", "Malfoy, Draco", "draco@hogwarts.edu", "lucius.malfoy@hogwarts.edu", "1", "Exempt", "1", "90"],
  ["Block 2", "Lovegood, Luna", "luna@hogwarts.edu", "xenophilius.lovegood@quibbler.org", "1", "1", "1", "92"]
], []);

// --- 4. Unstyled heading rows (no fill at all) + a name containing "Sum".
var PLAIN = sheetFrom([
  ["", "", "", "Tests"],
  ["", "Name", "Email", "Exam 1"],
  ["", "", "", ""],
  ["", "Period 1", "", ""],
  ["", "Sumner, Kate", "ksumner@school.edu", "88"],
  ["", "Average", "", "88"],
  ["", "Period 2", "", ""],
  ["", "Okafor, Ada", "aokafor@school.edu", "91"]
], []);

// --- 5. No sections, no parent column, wide sheet.
var FLAT = sheetFrom([
  ["", "", "", "Tests"],
  ["", "Name", "Email", "Exam 1"],
  ["", "", "", ""],
  ["", "Solo, Han", "hsolo@falcon.com", "77"],
  ["", "Organa, Leia", "lorgana@alderaan.gov", "99"]
], []);
