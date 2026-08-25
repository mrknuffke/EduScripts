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
