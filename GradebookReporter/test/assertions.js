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

// =================== selector dialog rendering ==============================
out.push("\n== 6. Selector dialog ==");

var withParents = renderSelector(DIALOG_SECTIONS, 'email', true);
check("dialog: two section cards", (withParents.match(/class="section-card"/g) || []).length, 2);
check("dialog: three student rows", (withParents.match(/class="stu-chk /g) || []).length, 3);
check("dialog: section names rendered",
  withParents.indexOf('AP Biology - Block 3') > -1 && withParents.indexOf('AP Biology - Block 5') > -1, true);
check("dialog: counts pluralised",
  withParents.indexOf('2 students') > -1 && withParents.indexOf('1 student<') > -1, true);
check("dialog: filter chips are All + one per section", (withParents.match(/class="chip/g) || []).length, 3);
check("dialog: sheet row indices preserved",
  withParents.indexOf('value="4"') > -1 && withParents.indexOf('value="8"') > -1, true);
check("dialog: mismatch badge rendered once", (withParents.match(/class="badge badge-warn"/g) || []).length, 1);
check("dialog: missing address shows None", withParents.indexOf('<i class="none">None</i>') > -1, true);
check("dialog: apostrophe in name escaped", withParents.indexOf("O&#39;Brien") > -1, true);
check("dialog: every student tagged with its section", (withParents.match(/data-section="sec-/g) || []).length, 3);

// With parent addresses present, all three destinations are offered.
check("parents present: three destination pills", (withParents.match(/class="dest-pill/g) || []).length, 3);
check("parents present: Parent Only offered", withParents.indexOf('Parent Only') > -1, true);
check("parents present: parent line shown per student", (withParents.match(/Parent:/g) || []).length, 3);
check("parents present: no forced default", withParents.indexOf('value="student" checked hidden'), -1);

// Universally blank parent column: the choice disappears rather than misleading.
var noParents = renderSelector(DIALOG_SECTIONS, 'email', false);
check("no parents: destination pills removed", (noParents.match(/class="dest-pill/g) || []).length, 0);
check("no parents: Parent Only not offered", noParents.indexOf('Parent Only'), -1);
check("no parents: Both not offered", noParents.indexOf('>✉️ Both<'), -1);
check("no parents: reason shown", noParents.indexOf('reports go to students only') > -1, true);
check("no parents: per-student parent line dropped", noParents.indexOf('Parent:'), -1);
check("no parents: student-only destination forced", noParents.indexOf('value="student" checked hidden') > -1, true);
check("no parents: students still listed", (noParents.match(/class="stu-chk /g) || []).length, 3);
check("no parents: sections still grouped", (noParents.match(/class="section-card"/g) || []).length, 2);

var drive = renderSelector(DIALOG_SECTIONS, 'drive', false);
check("drive mode: no destination section", (drive.match(/class="dest-pill/g) || []).length, 0);
check("drive mode: student-only still forced for preview",
  drive.indexOf('value="student" checked hidden') > -1, true);

// =================== completion encouragement ===============================
out.push("\n== 7. Completion encouragement ==");

function tierOf(name, value, isParent) {
  var note = buildEncouragementNote(statRows(name, value), !!isParent);
  return note ? note.tier : null;
}

// Reading the percentage, in either direction and either scale.
check("reads Completion Percentage", findCompletionPercent(statRows("Completion Percentage", "67")), 67);
check("reads a 0-1 fraction", findCompletionPercent(statRows("Completion Percentage", "0.67")), 67);
check("treats 1 as 100%", findCompletionPercent(statRows("Completion Percentage", "1")), 100);
check("strips a percent sign", findCompletionPercent(statRows("Completion Rate", "45%")), 45);
check("inverts % Incomplete", findCompletionPercent(statRows("% Incomplete", "30%")), 70);
check("ignores a raw missing COUNT", findCompletionPercent(statRows("Missing Assignments", "3")), null);
check("no completion stat at all", findCompletionPercent([{ name: "Lab 1", value: "1", isSummaryStat: false }]), null);

// Tier boundaries.
check("100% produces no note", tierOf("Completion Percentage", "100"), null);
check("99% is minor", tierOf("Completion Percentage", "99"), "minor");
check("80% is minor (boundary)", tierOf("Completion Percentage", "80"), "minor");
check("79.9% is moderate", tierOf("Completion Percentage", "79.9"), "moderate");
check("50% is moderate (boundary)", tierOf("Completion Percentage", "50"), "moderate");
check("49.9% is urgent", tierOf("Completion Percentage", "49.9"), "urgent");
check("0% is urgent", tierOf("Completion Percentage", "0"), "urgent");

// Every tier offers a way to reach out and never forecloses hope.
["99", "65", "20"].forEach(function (v) {
  var note = buildEncouragementNote(statRows("Completion Percentage", v), false);
  check(v + "%: three concrete steps", note.steps.length, 3);
  check(v + "%: invites contact",
    /email me|book a time|book an appointment|reply to this email|get in touch|reach out/i.test(
      note.steps.join(" ") + " " + note.closing), true);
  check(v + "%: shows the actual percentage", note.intro.indexOf(v + "%") > -1, true);
  check(v + "%: no despairing language",
    /too late|no hope|hopeless|give up|beyond saving|failed|failure/i.test(
      note.heading + " " + note.intro + " " + note.steps.join(" ") + " " + note.closing), false);
  check(v + "%: no blaming language",
    /lazy|careless|excuse|disappoint|unacceptable|should have|your own fault/i.test(
      note.heading + " " + note.intro + " " + note.steps.join(" ") + " " + note.closing), false);
});

// Parent wording differs from student wording.
var studentNote = buildEncouragementNote(statRows("Completion Percentage", "65"), false);
var parentNote = buildEncouragementNote(statRows("Completion Percentage", "65"), true);
check("parent note addresses the student in third person",
  parentNote.intro.indexOf("Your student's") > -1, true);
check("student note addresses the student directly",
  studentNote.intro.indexOf("Your completion") > -1, true);
check("same tier for both audiences", parentNote.tier, studentNote.tier);

// The configured reply-to address is surfaced when one is set.
STUBBED_REPLY_TO = "";
check("no reply-to configured: no address in closing",
  /reach me at/.test(buildEncouragementNote(statRows("Completion Percentage", "65"), false).closing), false);
STUBBED_REPLY_TO = "dknuffke@sas.edu.sg";
check("reply-to configured: address offered",
  buildEncouragementNote(statRows("Completion Percentage", "65"), false).closing.indexOf("dknuffke@sas.edu.sg") > -1, true);
STUBBED_REPLY_TO = "";

// HTML rendering.
var urgentHtml = generateHtmlEncouragement(statRows("Completion Percentage", "20"), false);
check("html: renders a block", urgentHtml.indexOf("<div") === 0, true);
check("html: three list items", (urgentHtml.match(/<li /g) || []).length, 3);
check("html: urgent uses the strongest accent", urgentHtml.indexOf("#c62828") > -1, true);
check("html: minor uses the calm accent",
  generateHtmlEncouragement(statRows("Completion Percentage", "90"), false).indexOf("#1a73e8") > -1, true);
check("html: moderate uses the warning accent",
  generateHtmlEncouragement(statRows("Completion Percentage", "65"), false).indexOf("#ef6c00") > -1, true);
check("html: nothing at 100%", generateHtmlEncouragement(statRows("Completion Percentage", "100"), false), "");
check("html: nothing when no completion stat is present",
  generateHtmlEncouragement([{ name: "Lab 1", value: "1", isSummaryStat: false }], false), "");

// =================== assignment categories ==================================
out.push("\n== 8. Assignment categories ==");

// Chem Sem 1 row 1: "Admin" labels the roster block (Name/Preferred/Email in
// A-C), "Formative Work" starts at column E, and column D has no label at all.
var chemCats = resolveCategoryRow(["", "", "Admin", "", "Formative Work", "", ""], 2);
check("roster label cleared from its own columns", chemCats.slice(0, 3), ["", "", ""]);
check("roster label does not leak into first assignment", chemCats[3], "");
check("real category still fills rightwards", chemCats.slice(4), ["Formative Work", "Formative Work", "Formative Work"]);

// A category that does start on the first assignment column still applies.
var spanning = resolveCategoryRow(["", "", "Admin", "Formative Work", "", "", ""], 2);
check("category on the first assignment column is kept", spanning[3], "Formative Work");
check("and fills across its span", spanning.slice(3), ["Formative Work", "Formative Work", "Formative Work", "Formative Work"]);

// Several categories in a row each own their span.
var multi = resolveCategoryRow(["Section", "", "", "", "Classwork", "", "Homework", ""], 3);
check("multiple categories keep their own spans",
  multi.slice(4), ["Classwork", "Classwork", "Homework", "Homework"]);
check("Section header cleared with the roster block", multi.slice(0, 4), ["", "", "", ""]);

// Degenerate inputs.
check("missing category row yields nothing", resolveCategoryRow(null, 2).length, 0);
check("blank category row stays blank", resolveCategoryRow(["", "", "", ""], 1), ["", "", "", ""]);
check("non-string cells tolerated", resolveCategoryRow([null, undefined, "Labs", ""], 1), ["", "", "Labs", "Labs"]);

// =================== assessment column detection ============================
out.push("\n== 9. Assessment columns ==");

// The Chem sheet's scored formative column: nothing recognisable in the
// standards row or category, so only the header can classify it.
check("scored 'Formative 1.1' recognised by header",
  matchesAssessmentColumn("Formative 1.1", "D", "Formative Work"), true);
check("'Summative 2.3' recognised by header",
  matchesAssessmentColumn("Summative 2.3", "", ""), true);
check("'Unit 1 Quiz' recognised by header",
  matchesAssessmentColumn("Unit 1 Quiz", "", ""), true);

// Still matched via the standards row or category, as before.
check("matched by category", matchesAssessmentColumn("1.1", "", "Topic Quest Labs"), true);
check("matched by standards row", matchesAssessmentColumn("1.1", "Lab Skills", ""), true);
check("short 'wa' still matches in a category", matchesAssessmentColumn("1.1", "", "WA 3"), true);

// The two-letter keyword must not fire on ordinary headers.
check("'Water Stations' header is not an assessment",
  matchesAssessmentColumn("AC: Water Stations", "", ""), false);
check("'ID: 1.2 Water' header is not an assessment",
  matchesAssessmentColumn("ID: 1.2 Water", "", ""), false);
check("'AC: Macromolecules' is not an assessment",
  matchesAssessmentColumn("AC: Macromolecules", "", ""), false);
check("'Initial Appointment' is not an assessment",
  matchesAssessmentColumn("AC: Initial Appointment", "", ""), false);
check("'Journal Submitted' is not an assessment",
  matchesAssessmentColumn("AC: Journal Submitted", "", ""), false);
check("empty everything is not an assessment",
  matchesAssessmentColumn("", "", ""), false);
check("null inputs tolerated",
  matchesAssessmentColumn(null, null, null), false);

// Documented consequence: a header containing "test" now reports even when done.
check("'Complete Pre-Test' header DOES match (see README note)",
  matchesAssessmentColumn("AC: Complete Pre-Test", "", ""), true);

// =================== 0-4 rubric scoring =====================================
out.push("\n== 10. Rubric scoring ==");

check("0 -> Not Yet Evident", rubricLabelFor("0"), "Not Yet Evident");
check("1 -> Emerging", rubricLabelFor("1"), "Emerging");
check("2 -> Developing", rubricLabelFor("2"), "Developing");
check("3 -> Meeting", rubricLabelFor("3"), "Meeting");
check("4 -> Meeting with Distinction", rubricLabelFor("4"), "Meeting with Distinction");
check("numeric input works too", rubricLabelFor(3), "Meeting");

// Blank is not zero: "no score yet" must not read as "Not Yet Evident".
check("empty string is not a zero", rubricLabelFor(""), null);
check("null is not a zero", rubricLabelFor(null), null);
check("undefined is not a zero", rubricLabelFor(undefined), null);
check("whitespace is not a zero", rubricLabelFor("   "), null);

// Anything off the 0-4 scale passes through untouched.
check("half marks pass through", rubricLabelFor("0.5"), null);
check("3.5 passes through", rubricLabelFor("3.5"), null);
check("5 is off the scale", rubricLabelFor("5"), null);
check("a percentage passes through", rubricLabelFor("95"), null);
check("checkbox text passes through", rubricLabelFor("TRUE"), null);
check("'m' passes through", rubricLabelFor("m"), null);
check("'Exempt' passes through", rubricLabelFor("Exempt"), null);

// Which columns use the rubric.
check("Chem formative column", isRubricScoredColumn("Chemistry", "Formative 1.1", "", "Formative 1.1", false), true);
check("Chem summative standard", isRubricScoredColumn("Chemistry", "1.3", "", "1.3", true), true);
check("Chem summative by header", isRubricScoredColumn("Chemistry", "Summative 2.1", "", "Summative 2.1", false), true);
check("Chem formative by category", isRubricScoredColumn("Chemistry", "1.5 Bonding", "Formative Work", "1.5 Bonding", false), true);
check("XL Chemistry included", isRubricScoredColumn("XL Chemistry", "Formative 1.1", "", "Formative 1.1", false), true);
check("AP Bio lab by name", isRubricScoredColumn("AP Biology", "Lab 3", "", "Lab 3", false), true);
check("AP Bio lab by category", isRubricScoredColumn("AP Biology", "Topic Quest 2", "Labs", "Topic Quest 2", false), true);

// Columns that must NOT be converted.
check("Chem checkbox activity is not rubric",
  isRubricScoredColumn("Chemistry", "AC: Journal Submitted", "Admin", "AC: Journal Submitted", false), false);
check("AP Bio completion column is not rubric",
  isRubricScoredColumn("AP Biology", "AC: Water Stations", "Completion", "AC: Water Stations", false), false);
check("AP Bio topic question is not rubric",
  isRubricScoredColumn("AP Biology", "ID: 1.2 Water", "Topic Questions", "ID: 1.2 Water", false), false);
check("unknown subject never uses the rubric",
  isRubricScoredColumn("Grade", "Formative 1.1", "Formative Work", "Formative 1.1", false), false);

// The combination that motivated ordering the rubric first.
check("a rubric 1 is Emerging, not Complete", rubricLabelFor("1"), "Emerging");
check("a rubric 0 is Not Yet Evident, not Missing", rubricLabelFor("0"), "Not Yet Evident");

out.push("\n" + (failures === 0 ? "ALL " + (out.filter(function (l) { return l.indexOf("  PASS") === 0; }).length) + " CHECKS PASSED"
                                : failures + " CHECK(S) FAILED"));
console.log(out.join("\n"));
