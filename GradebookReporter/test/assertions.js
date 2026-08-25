// ============================== assertions =================================
var out = [], failures = 0;
function check(label, actual, expected) {
  var a = JSON.stringify(actual), e = JSON.stringify(expected), ok = a === e;
  if (!ok) { failures++; out.push("  FAIL " + label + "\n        got:      " + a + "\n        expected: " + e); }
  else out.push("  PASS " + label);
}
function scan(title, sheet) {
  var roster = scanGradebookRoster(sheet);
  var groups = groupStudentsBySection(roster.students, roster.sections);
  out.push("\n== " + title + " ==");
  out.push("  header row " + (roster.headerRow + 1) + ", data from row " + (roster.firstDataRow + 1) +
           ", cols " + JSON.stringify(roster.cols));
  out.push("  " + groups.map(function (g) {
    return g.name + " (" + g.students.length + "): " + g.students.map(function (s) { return s.name; }).join(", ");
  }).join("\n  "));
  return { roster: roster, groups: groups, names: groups.map(function (g) { return g.name; }) };
}

var a = scan("1. AP Bio Sem 1", AP_BIO);
check("AP Bio: roster header is row 2", a.roster.headerRow + 1, 2);
check("AP Bio: email column found (D)", a.roster.cols.email, 3);
check("AP Bio: column A not used as Section", a.roster.cols.section, -1);
check("AP Bio: no TRUE/FALSE sections", a.names, ["AP Bio 1 (B2/D1)", "AP Bio 3 (A4/C2)"]);
check("AP Bio: 6 students", a.roster.students.length, 6);
check("AP Bio: headcount on banner did not create a student",
  a.roster.students.filter(function (s) { return s.name.indexOf("AP Bio") > -1; }).length, 0);
check("AP Bio: Average/Sum rows dropped",
  a.roster.students.filter(function (s) { return s.name.indexOf("Average") > -1; }).length, 0);
check("AP Bio: split 3 / 3", a.groups.map(function (g) { return g.students.length; }), [3, 3]);
check("AP Bio: emails captured", a.groups[0].students[0].email, "lee781513@sas.edu.sg");

var c = scan("2. Chem Sem 1", CHEM);
check("Chem: roster header is row 4", c.roster.headerRow + 1, 4);
check("Chem: data starts row 5", c.roster.firstDataRow + 1, 5);
check("Chem: email column found (D)", c.roster.cols.email, 3);
check("Chem: no student called 'Name'",
  c.roster.students.filter(function (s) { return s.name === "Name"; }).length, 0);
check("Chem: no Ungrouped bucket", c.names.indexOf("Ungrouped"), -1);
check("Chem: one section from merged banner", c.names, ["Chem 2 (A2/C4)"]);
check("Chem: 5 students", c.roster.students.length, 5);
check("Chem: emails resolved, not None", c.roster.students.map(function (s) { return s.email !== ""; }), [true, true, true, true, true]);
check("Chem: checkbox cols excluded from Section", c.roster.cols.section, -1);
check("Chem: Average row dropped",
  c.roster.students.filter(function (s) { return s.name.indexOf("Average") > -1; }).length, 0);

var c2 = scan("2b. Chem Sem 1 (roster shifted to columns A/B/C)", CHEM_V2);
check("ChemV2: roster header is row 4", c2.roster.headerRow + 1, 4);
check("ChemV2: Name is column A", c2.roster.cols.name, 0);
check("ChemV2: Preferred Name is column B", c2.roster.cols.preferred, 1);
check("ChemV2: Email is column C", c2.roster.cols.email, 2);
check("ChemV2: assignments begin after column C", c2.roster.cols.lastRosterCol, 2);
check("ChemV2: 4 students", c2.roster.students.length, 4);
check("ChemV2: every student has a name", c2.roster.students.every(function (s) { return s.name !== ""; }), true);
check("ChemV2: every student has an email", c2.roster.students.every(function (s) { return s.email.indexOf("@") > -1; }), true);
check("ChemV2: banner became the section", c2.names, ["A2/C4"]);
check("ChemV2: Average row dropped",
  c2.roster.students.filter(function (s) { return s.name.indexOf("Average") > -1; }).length, 0);

// Report generation resolves the same columns from values it already holds.
var rc = resolveRosterColumns(CHEM_V2.getDataRange().getDisplayValues());
check("ChemV2 report path: name column", rc.name, 0);
check("ChemV2 report path: email column", rc.email, 2);
check("ChemV2 report path: last roster column", rc.lastRosterCol, 2);
var rcBio = resolveRosterColumns(AP_BIO.getDataRange().getDisplayValues());
check("AP Bio report path: name column B", rcBio.name, 1);
check("AP Bio report path: email column D", rcBio.email, 3);
check("AP Bio report path: last roster column D", rcBio.lastRosterCol, 3);
var rcDemo = resolveRosterColumns(DEMO.getDataRange().getDisplayValues());
check("Demo report path: last roster column D", rcDemo.lastRosterCol, 3);

var d = scan("3. Demo gradebook (per-row Section column)", DEMO);
check("Demo: Section column is A", d.roster.cols.section, 0);
check("Demo: 4 students", d.roster.students.length, 4);
check("Demo: two blocks", d.names, ["Block 1", "Block 2"]);
check("Demo: parent emails read", d.groups[0].students[0].parentEmail !== "", true);

var pl = scan("4. Unstyled heading rows", PLAIN);
check("Plain: sections without any fill styling", pl.names, ["Period 1", "Period 2"]);
check("Plain: 'Sumner' survives the summary filter", pl.groups[0].students[0].name, "Sumner, Kate");
check("Plain: bare 'Average' row dropped", pl.roster.students.length, 2);

var f = scan("5. No sections at all", FLAT);
check("Flat: single Ungrouped bucket", f.names, ["Ungrouped"]);
check("Flat: 2 students", f.roster.students.length, 2);

out.push("\n" + (failures === 0 ? "ALL " + (out.filter(function (l) { return l.indexOf("  PASS") === 0; }).length) + " CHECKS PASSED"
                                : failures + " CHECK(S) FAILED"));
console.log(out.join("\n"));
