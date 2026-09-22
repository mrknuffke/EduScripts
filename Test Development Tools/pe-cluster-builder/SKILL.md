---
name: pe-cluster-builder
description: 'Build a single-standard assessment cluster: one NGSS Performance Expectation (or one standard from another framework), one phenomenon-anchored stimulus, roughly 40 minutes of selected- and constructed-response items, every item anchored to an evidence statement feature, scored 0/1 per criterion and rolled up to a 1-4 SBG level per reporting bucket. Delivers a teacher review draft first, then student paper, answer key and scoring sheet (.docx), and optionally a Google Form plus auto-scoring spreadsheet via Apps Script. Use whenever a teacher asks for a PE assessment, item cluster, task cluster, single-PE assessment, standards-based quiz on one PE, a Form-based PE check, or wants to convert an existing assessment into this format, in any science discipline (life, physical, earth and space, engineering) or grade band, or in a non-NGSS framework (AP, IB, state standards). Also use for re-running checks on, revising, or re-scoring an existing cluster built with this skill.'
---

# PE Cluster Builder

Builds one **assessment cluster**: a stand-alone instrument on exactly one standard, given as
soon as that standard has been taught. The architecture is discipline-neutral. Everything that
varies by course (the reporting buckets, the scope list, the item floor, fonts, paper size, local
context) lives in a **course profile**, so the same skill serves chemistry, biology, physics, earth
science or a non-NGSS course without edits to this file.

## Files in this skill

| File | Read it | Governs |
|---|---|---|
| `references/design-rules.md` | Before drafting any item | Coverage derivation, anchoring, cross-scoring, compound and OR-pathway criteria, Distinction probes, context, item architecture, language |
| `references/scoring.md` | Before writing criteria text, and again before building | Gates, Distinction rules, standing blocks, grading reflection, workbook analysis |
| `references/screening-tool.md` | Step 2 (as design inputs) and Step 5 (as a report) | The eleven-item 3D screen |
| `references/content-schema.md` | Before writing `content.json` | The single source of truth every output is generated from |
| `references/build.md` | Step 7 | .docx geometry and components, the Form route, validation |
| `references/domain-adaptation.md` | Step 1, whenever the course is not the one the profile was written for, or the framework is not NGSS | Discipline-specific traps, non-NGSS frameworks, non-science use |
| `assets/course-profile-template.md` | Step 0 when no profile exists | The fill-in profile |
| `assets/example-content.json` | When unsure what a finished spec looks like | A complete worked cluster (HS-LS2-3) |
| `assets/form-builder-template.gs` | Never edited by hand; `make_form_script.py` fills it | The Apps Script Form + scoring workbook builder |
| `scripts/check_cluster.py` | After every change to `content.json` | Automated structural, item and consistency checks |
| `scripts/build_docs.js` | Step 6 (draft) and Step 7 (final) | Generates every .docx from `content.json` |
| `scripts/make_form_script.py` | Step 7, Form route | Generates the ready-to-paste .gs from `content.json` |

## Non-negotiables

These hold on every build and are why the instrument is defensible in a grade appeal.

1. **One standard.** Pressure toward a second PE means the teacher wants a different instrument.
   Say so; do not stretch the cluster.
2. **The evidence statement sets coverage.** Every in-scope observable feature is assessed at least
   once and at most twice. The bucket set never sets the criterion count.
3. **Scope comes from the course's taught-skill list, not the PE text.** A feature with no matching
   code on the scope list is out, however clearly the PE licenses it.
4. **Every criterion is anchored.** It names one evidence-statement feature, or it is a declared
   Distinction probe that names the feature it extends. Nothing else is scored.
5. **Draft before build.** Nothing operational is generated until a teacher or PLC has reviewed
   the draft document.
6. **One source of truth.** Every output is generated from one `content.json`. Never hand-edit a
   .docx or the .gs to change content; change the JSON and regenerate.
7. **Ask one question at a time** during setup and design. Never batch questions.

## Workflow

### Step 0. Load or build the course profile

Check for `references/course-profile.md` (a filled copy of the template). If it exists, load it
and confirm the course name with the user in one line. If it does not, or the user is building for
a different course than the profile describes, run the setup interview from
`assets/course-profile-template.md`, **one question at a time**, in the order the template lists.
At the end, write the filled profile out as a file and tell the user they can drop it into the
skill's `references/` folder so the interview never runs again.

The profile answers: framework, reporting taxonomy and bucket map, scope list and its codes, item
floor per bucket, Distinction mode, per-assessment vs per-period levels, SBG labels, paper size,
fonts, accent colour, spelling, locale for context anchoring, default delivery mode.

### Step 1. Establish the standard, the evidence, and the scope

Ask, one at a time, only for what the profile does not already hold:

1. **Which PE (or standard)?**
2. **The evidence statement.** Ask for the PDF or text. For NGSS these are published by Achieve
   (nextgenscience.org, "Evidence Statements", organised by grade band). Do not reconstruct one
   from memory. If the file is missing or unreadable, stop and say so in one line. For non-NGSS
   frameworks, read `references/domain-adaptation.md` for what stands in for the evidence
   statement.
3. **The scope list for this PE**, from the profile's source (skill registry, learning targets,
   unit review sheet). Write the filtered list back in the conversation with its codes.
4. **Terms deliberately not yet introduced.** "Are there terms adjacent to this content that the
   unit has held back, which a complete answer would normally reach for?" Flag each and confirm the
   substitute reasoning path.
5. **Delivery mode:** paper, or Google Form plus scoring workbook. Ask now; the item text differs.

Then enumerate every observable feature row (Components/Relationships/Connections, or
Explanation/Evidence/Reasoning/Revision, whatever the statement uses), filter by scope, and write
out both lists: features in, features out and why. This list is the spine of the build.

Read `references/domain-adaptation.md` for the discipline now, not later. It names the traps that
cost the most rework (quantitative collisions in physics, teleological distractors in biology,
real-data attribution in earth science, problem-not-phenomenon in engineering).

### Step 2. Design conversation

Read `references/design-rules.md` and `references/screening-tool.md`. Settle with the user, one
decision at a time:

- The context: adjacent to the unit storyline, not a repeat of it, locally anchored per the profile.
  "Someone else's investigation" is the most reliable generator.
- The bucket set: every bucket this PE maps to under the profile's taxonomy. A mapped bucket with
  no criteria is a design failure.
- Which features will be assessed twice, and how the two exposures differ.
- Where cross-scoring is needed to populate a bucket the features do not feed directly (usually
  the CCC one).
- Where the data comes from, whether it is constructed, and its attribution.

Keep a running decision log (template at the end of this file).

### Step 3. Write `content.json`

Read `references/content-schema.md`. Write the full spec: meta, scope, features, buckets,
stimulus, glossary, and every criterion with its item text, options, keys, misconceptions, and
what-earns-the-point. **Assign key positions for every selected-response item before writing any
options**, then write to that plan.

### Step 4. Run the checker

```bash
python3 scripts/check_cluster.py content.json
```

Fix every error. Warnings are either fixed or carried into the draft's open decisions with a
default. Re-run after every edit to option text, including edits that were not about the options.
Recompute every derived number in the stimulus by script, not by eye.

### Step 5. Screening report

Screen the spec against all eleven items in `references/screening-tool.md`. Report in the
conversation as a table (item, rating, evidence, change offered). It is a report, not a gate. Ask
which changes to make, apply the accepted ones in `content.json`, re-run the checker.

### Step 6. The review draft

```bash
node scripts/build_docs.js content.json out/ --draft
```

Produces `<slug>-DRAFT-review.docx`: header, context and stimulus as students see it, every item in
student order with key, misconceptions, criterion and skill code, the evidence-statement coverage
table, the Distinction probes, the criteria table in bucket order, the gate table, the grading
reflection, the screening report, and numbered open decisions (each a question with a default).
Deliver it and stop. Tell the user what to look at first (the coverage table).

When the review comes back: apply the changes to `content.json`, state what changed, re-run the
checker, and only then build.

### Step 7. Build

Read `references/build.md`.

- **Paper:** `node scripts/build_docs.js content.json out/ --final` gives Student, Answer Key and
  Scoring Sheet.
- **Form:** also run `python3 scripts/make_form_script.py content.json out/`. Deliver the .gs,
  every stimulus PNG named in `images`, and the scoring sheet .docx. Tell the user the upload and
  run steps in `references/build.md`.

Validate every .docx, render to images, and look at them before delivering.

### Step 8. Delivery report

In the conversation, not in any document: files produced, criterion counts per bucket, gate table,
checker output (errors zero, remaining warnings listed), screening result with declined changes as
known gaps, and anything the teacher must do before students sit it (upload images, set the access
code, attach a data citation).

## When the user brings an existing assessment

Converting an existing single-PE assessment or quiz: read it, map each existing item to a feature
(or to nothing), list uncovered features and unanchored items, and present that gap analysis before
drafting. Items that map to nothing come off or are rewritten; they are not kept as warm-ups.

Porting a finished cluster to another PE or course: the JSON structure carries over, the content
does not. Start again at Step 1 for the new standard. Reusing the old context is fine only if it is
genuinely adjacent to the new unit.

## Format rules for every output

Curly quotes and apostrophes. No em-dashes. Unicode subscripts and superscripts in every formula,
including student-facing prose (CO₂, 6.02 × 10²³, m s⁻²). No LaTeX, MathJax or dollar-sign
notation anywhere. Zero revision residue in student- or teacher-facing text ("updated", "revised",
"new this year"); rationale for changes lives in the conversation only. `check_cluster.py` enforces
most of this on the JSON.

## Decision log template

Keep this in the conversation and update it as decisions land. Do not build until every line is
settled.

```
Course profile loaded: ___
Standard: ___   Framework: ___
Evidence statement source: ___
Features in scope (n): ___   Features out of scope, with reason: ___
Held-back terms and substitute reasoning: ___
Delivery mode: paper / form
Buckets scored (from the profile map): ___
Item floor per bucket: ___   Levels scope: per-assessment / per-period
Distinction mode: probe / perfect / none   (MWD requires top gate: yes / no)
Context, and how it is adjacent without repeating the unit: ___
Data: constructed / real (source ___)
Features assessed twice, and how the exposures differ: ___
Cross-scored criteria (and any tagged to 3+ buckets): ___
Compound / OR-pathway criteria: ___
Key position plan: ___
Checker: errors ___ warnings ___
Screening report: accepted ___ declined ___
Draft reviewed by ___ on ___; changes applied: ___
```
