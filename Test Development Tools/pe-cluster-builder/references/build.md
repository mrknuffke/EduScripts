# Build

How outputs are generated from `content.json`, what they look like, and how to verify them.

---

## 1. Commands

```bash
# structural checks: run after every edit
python3 scripts/check_cluster.py content.json

# review draft (Step 6)
node scripts/build_docs.js content.json out/ --draft

# final paper set (Step 7): Student, Answer Key, Scoring Sheet
node scripts/build_docs.js content.json out/ --final

# Form route (Step 7): ready-to-paste Apps Script
python3 scripts/make_form_script.py content.json out/
```

`build_docs.js` needs the `docx` npm package (preinstalled in Claude's container; otherwise
`npm install docx`). It refuses to build `--final` while the checker reports errors. Stimulus PNGs
named in `stimulus.blocks[].image` are embedded if they sit next to `content.json`.

---

## 2. Page and type

Set from `meta.format`:

| Setting | A4 (default) | Letter |
|---|---|---|
| Page (DXA) | 11906 × 16838 | 12240 × 15840 |
| Margins | 720 sides and top, 1008 bottom | same |
| Content width | 10466 | 10800 |

Every table uses the full content width with explicit `columnWidths` and per-cell widths in DXA
(percentage widths break in Google Docs). Shading uses `ShadingType.CLEAR`.

Fonts default to Garamond body 12 pt and Montserrat Medium headings. Neither is installed in
Claude's build container, so a rendered preview substitutes and page counts are approximate.
Verify fonts by inspecting the XML, not the preview.

---

## 3. Components

- **Title block:** cluster title, standard code, then the Name / Block / Date underscore line.
- **Info box:** single-cell shaded table: what the assessment asks, resources, time.
- **Stimulus:** context paragraphs, then each block (table or text) with title and source or
  "illustrative" note, then the glossary box.
- **Selected response:** stem, then options lettered A to D in the spec's order.
- **Constructed response:** stem with sentence cap, then answer lines.
- **Answer lines** are table rows with a bottom border only, row height 550 twips at-least. Never
  bottom-bordered paragraphs: they break on Google Docs import.
- **Draw boxes and data grids** carry `cantSplit` so they never break across a page.

### Answer-line calibration

A full-width line holds about 12 handwritten words; a response sentence is about 18 words. Lines =
sentence cap × 18 ÷ 12, rounded up, plus one line of slack.

| Sentence cap | Minimum lines |
|---|---|
| 2 | 4 |
| 3 | 6 |
| 4 | 7 |
| 5 | 9 |
| 6 | 10 |

The checker warns when `lines` is below this rule for the item's `sentence_cap`.

---

## 4. Documents

| File | Mode | Audience | Carries |
|---|---|---|---|
| `<slug>-DRAFT-review.docx` | `--draft` | teacher or PLC | everything, including keys, misconceptions, coverage table, gates, screening, open decisions |
| `<slug>-Student.docx` | `--final` | students | title block, info box, stimulus, glossary, items. No key markers of any kind |
| `<slug>-Answer-Key.docx` | `--final` | markers | every item: key or full what-earns-the-point, criterion ID, bucket, follow-through note |
| `<slug>-Scoring-Sheet.docx` | `--final` | markers | criteria in bucket order with cue phrases, probes, merged rollup and gates, standing blocks |

Open decisions and the screening report appear only in the draft.

---

## 5. Verify before delivering

```bash
python3 /mnt/skills/public/docx/scripts/office/validate.py out/<file>.docx
python3 /mnt/skills/public/docx/scripts/office/soffice.py --headless --convert-to pdf out/<file>.docx
pdftoppm -png -r 80 out/<file>.pdf out/page
```

Look at every page image. Check: no key markers on the student paper, answer space where the
sentence caps say it should be, tables not overflowing, scoring sheet within two pages. If the
docx skill is not present at that path, skip validation and say so in the delivery report.

Grep every output for revision residue ("updated", "revised", "new this year", "v2") and for
straight quotes or em-dashes that slipped in from a figure caption.

---

## 6. Form route

`make_form_script.py` writes `<slug>-form-builder.gs`: the template in
`assets/form-builder-template.gs` with the spec block generated from `content.json`. Never edit the
spec block by hand.

What the teacher does (put these steps in the delivery report):

1. Upload every stimulus PNG to Google Drive with the exact file names listed in the script's
   `CONFIG.images`. A missing image falls back to the text version, so the build degrades rather
   than fails.
2. Go to script.google.com, New project, paste the whole .gs file, Save.
3. Run `buildAssessment`. Authorise when prompted (it needs Forms, Sheets and Drive).
4. The execution log prints the Form edit link, the live link and the scoring workbook link.
5. Set the access code in the Form's validation, or edit `CONFIG.accessCode` and rebuild. Change
   the code to close the Form.
6. Grade release is manual by default and cannot be changed from the script.

Re-running creates a new Form and workbook; it never modifies an existing one.

Section 1 of every Form is fixed: Name, Block, Teacher (dropdown from `form.teachers`), Access code
(regex-validated, case-insensitive). The access code is a soft gate against casual sharing, not
security. Teacher carries through to the workbook so grading and analysis filter by section.

The script validates the spec before creating anything (bucket counts, exclusivity, anchoring,
probes, key distribution, option lengths, roster and code placeholders) and throws on errors. That
duplicates `check_cluster.py` deliberately: the teacher may edit the .gs later.
