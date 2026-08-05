# Project Brief: Learn My Students

A local, zero-install flashcard tool that turns PowerSchool photo-roster PDF exports into a Leitner-box drill for learning student names and faces.

This document is written to be handed to a coding agent. It specifies behaviour, data formats, and acceptance criteria. Where a design decision was made deliberately, the reasoning is included so the agent does not "helpfully" undo it.

Note on quoting conventions: prose in this document uses curly quotes. All code, JSON, and filenames use straight quotes and must be reproduced literally.

---

## 1. Purpose and users

**Primary user:** a high school teacher with roughly five class sections and roughly one hundred students, who needs to learn names and faces fast at the start of a semester. Replaces a deprecated in-house web app.

**Secondary users:** colleagues at the same school who receive the tool as a file and set it up themselves from a README with no support of any kind.

**Hard consequence of that second point:** every design decision resolves in favour of "works when double-clicked" over "is architecturally tidy." There is no install step, no package manager, no terminal, no build tooling on the user's machine, no account, and no server.

**Deployment context:** macOS, current Safari and Chrome. Distribution is by handing someone a file (AirDrop, Drive, email). A public GitHub repository is possible later but is explicitly not assumed.

---

## 2. Non-goals

Do not build these. Each was considered and cut.

- No accounts, sync, cloud storage, or telemetry of any kind.
- No server, no backend, no build step required of the end user.
- No face detection, face recognition, or ML. The PDF already tells us where every face is.
- No OCR. The source PDFs have a real text layer (verified — see section 4).
- No spaced-repetition sophistication beyond Leitner. No SM-2, no FSRS, no ease factors.
- No mobile or responsive layout beyond "does not break on a laptop screen." This is a desktop tool.
- No multi-user, no sharing decks between teachers. Decks contain photographs of minors.
- No editing of card content in v1 (see section 8.4 for the deliberate exception and the forward-compatibility hook).

---

## 3. Architecture

**One self-contained HTML file.** All JavaScript, CSS, and vendored library code inlined. The user double-clicks it and it opens in their browser at a `file://` origin. No network access at runtime, ever.

**Do not depend on a CDN.** The tool must work on a plane, on school wifi that blocks unfamiliar domains, and in five years when a CDN URL has changed. Vendor `pdf.js` into the file.

**Repository layout** (for maintenance; the end user only ever receives the built file):

```
name-learner.html          <- the built artifact, committed, this is what you hand people
src/
  index.template.html
  app.js
  app.css
  vendor/
    pdf.min.js             <- vendored, pinned version, checked in
    pdf.worker.min.js      <- vendored, pinned version, checked in
build.py                   <- stdlib only, no dependencies, inlines src/ into name-learner.html
README.md
.gitignore
```

`build.py` must run on a stock macOS Python 3 with no `pip install`. It does string substitution of the source files into the template. That is all it needs to do.

### 3.1 pdf.js worker constraint

`pdf.js` normally loads its worker from a separate file, which fails at a `file://` origin. Two viable approaches, in order of preference:

1. Inline the worker source into a `Blob`, create an object URL from it, and set `pdfjsLib.GlobalWorkerOptions.workerSrc` to that URL. Verify this is permitted at `file://` in both Safari and Chrome.
2. If (1) fails in either browser, disable the worker entirely (`pdfjsLib.GlobalWorkerOptions.workerSrc = ''` with `disableWorker`, or the equivalent for the pinned version) and parse on the main thread. Performance is irrelevant here — these are two-page PDFs — so blocking the main thread for a second is acceptable.

**The agent must actually test this on macOS Safari and Chrome before declaring the import path done.** This is the single most likely place the "just double-click it" promise breaks, and it will fail silently with an unhelpful console error if it fails.

---

## 4. Source data (verified)

A real export was inspected. Findings, which the implementation can rely on but should not assume are universal:

- Producer: `PowerSchool 26.6.0`, report title `Photos: 12/page`, PDF 1.4, A4 (595 × 842 pt).
- Twelve students per page in a 4 × 3 grid, with a partial final row when the class size is not a multiple of twelve.
- A real embedded text layer. Fonts are embedded CID TrueType with Identity-H encoding and unicode maps, so text extraction is clean. **No OCR is required.**
- One embedded JPEG per student, RGB, 200 px wide, height varying between roughly 200 and 301 px depending on the source photo's aspect ratio.
- Each photo is followed directly beneath by two text lines: the student's full name, then their student ID number.
- A page header reading the school name, positioned well above the grid.

Sample geometry from page 1 (points, PDF coordinate space as reported by PyMuPDF with a top-left origin):

| Element | Rect |
| --- | --- |
| Photo, row 1 col 1 | 82.3, 100.8, 151.7, 205.2 |
| Name line beneath it | 91.0, 208.0, 139.8, 217.9 |
| ID line beneath that | 104.0, 221.0, 126.5, 230.9 |
| Photo, row 2 col 1 | 82.3, 235.4, 151.7, 339.8 |

Photo widths vary slightly by row and column because the aspect ratio is preserved within a fixed cell, so **do not hard-code a grid.** Derive positions from the document.

### 4.1 Pairing algorithm (validated)

This exact approach was run against the sample export and correctly paired 22 of 22 students across two pages, with zero errors. Reimplement it in JavaScript:

For each page:
1. Collect all text lines with their bounding boxes.
2. Collect all placed image rectangles.
3. For each image rectangle, find text lines whose top edge falls within roughly 40 pt below the image's bottom edge, and which overlap the image horizontally within a tolerance of roughly 40 pt.
4. Sort those candidates top to bottom. The first is the student name. The second is the student ID.
5. Discard the ID (see section 9).

Reference implementation that produced the validated result:

```python
cand = [
    (line.y0, line, text)
    for line, text in lines
    if line.y0 >= rect.y1 - 2
    and line.y0 < rect.y1 + 40
    and line.x1 > rect.x0 - 40
    and line.x0 < rect.x1 + 40
]
cand.sort()
name = cand[0][2]
```

Note that block-level text extraction merges text across adjacent grid cells and is unusable. **Extract at the line level.** In `pdf.js`, `getTextContent()` returns items with transform matrices from which line boxes can be reconstructed; group items into lines by shared baseline before matching.

### 4.2 Getting the photo pixels

Two options. Prefer the second unless it produces visibly poor results.

1. **Extract the embedded JPEG** via the operator list and `page.objs`. Gives the original bytes at native resolution. Fiddly, with colorspace and image-mask edge cases.
2. **Render and crop.** Render the page to a canvas at a scale factor of 3 (so a 70 pt wide photo becomes roughly 210 px), then crop the canvas to each image rectangle scaled by the same factor. Re-encode the crop as JPEG at quality 0.82, capped at 400 px on the long edge.

Option 2 is strongly preferred: it is far shorter, has no colorspace failure modes, and the output quality is more than sufficient for recognising a face on screen. The image rectangles needed for the crop are the same ones already computed for pairing, so this adds almost no code.

To obtain those rectangles in `pdf.js`, walk the operator list, maintain a transform matrix stack across `OPS.save`, `OPS.restore`, and `OPS.transform`, and record the current transform at each `OPS.paintImageXObject` or `OPS.paintJpegXObject`. The image occupies the unit square under that transform.

### 4.3 Import must be human-confirmed

Do not trust the pairing blind. After parsing, show a confirmation grid: every extracted photo with its extracted name beneath it, and a count ("Found 22 students in Chemistry_A2"). The user clicks Confirm or Cancel.

If any photo produced zero candidate text lines, or any text line was claimed by two photos, flag those cards visually in the confirmation grid and state the problem in plain language. A ragged final row, a missing-photo placeholder, or a name that wraps to two lines are the realistic failure modes, and a teacher looking at the grid will spot a mismatch instantly where an algorithm will not.

### 4.4 Class name

Derive the deck name from the PDF filename, minus the extension, with underscores converted to spaces. `Chemistry_A2.pdf` becomes `Chemistry A2`. Let the user edit it in the confirmation screen before saving.

---

## 5. Data model

The deck file is the source of truth. Browser storage is a convenience cache only, and the app must be fully functional if the cache is empty or unavailable.

**Rationale, so the agent does not substitute IndexedDB:** at a `file://` origin, Chrome treats each file as an opaque origin and blocks IndexedDB, and `localStorage` behaviour on `file://` has varied across browser versions. Progress that lives only in browser storage will silently vanish on the exact setup these users have. A file the user holds also gives them backup, transfer to a home laptop, and re-import next semester for free.

Deck file: `<Class Name>.deck.json`, plain JSON, UTF-8.

```json
{
  "schema": 1,
  "deckName": "Chemistry A2",
  "createdAt": "2026-08-05T12:00:00.000Z",
  "updatedAt": "2026-08-05T12:00:00.000Z",
  "sessionCount": 0,
  "cards": [
    {
      "id": "c1",
      "name": "Firstname Lastname",
      "preferredName": null,
      "photo": "data:image/jpeg;base64,...",
      "box": 1,
      "dueSession": 1,
      "seen": 0,
      "correct": 0,
      "missed": 0
    }
  ]
}
```

- `id` is stable and locally generated. It is not a student ID and has no meaning outside the file.
- `preferredName` is written as `null` at import and there is no UI to set it in v1. The matcher must honour it if present. This is a deliberate forward-compatibility hook: it lets the user hand-edit the JSON in the interim and lets a later version add an editor without invalidating existing decks.
- `photo` is a base64 data URL. At roughly 8 KB per photo, a 22-student deck lands around 250 KB. A combined five-class deck stays under 1.5 MB. This is fine.

**Do not store student ID numbers.** See section 9.

---

## 6. Grading

### 6.1 Normalisation

Applied to both the typed input and the target before any comparison:

1. Unicode NFD decomposition, then strip combining marks (so `Renée` matches `Renee`).
2. Lowercase.
3. Replace hyphens and apostrophes, both straight and curly, with spaces.
4. Strip all remaining punctuation.
5. Collapse runs of whitespace, then trim.

### 6.2 Matching

Compute Levenshtein distance on normalised strings. Edit tolerance scales with the length of the target:

| Target length | Allowed edits |
| --- | --- |
| 1 to 3 | 0 |
| 4 to 7 | 1 |
| 8 or more | 2 |

The zero-tolerance band for very short targets is deliberate: `Amy` and `Ami` are one edit apart and are different people. Do not soften this.

Evaluation order:

1. Input matches the full name within tolerance, or matches `preferredName` within tolerance → **Correct.** If it was not an exact string match, display the correct spelling alongside the confirmation.
2. Input matches any single token of the name within tolerance, but not the whole name → **Partial.**
3. Otherwise → **Miss.**

Multi-token input that is not the full name (for example, first plus last where the roster has a middle name) should be treated by comparing the input against the full name and also against each token; the best outcome across those comparisons wins.

### 6.3 Partial credit

A Partial reveals the full name and holds the card in its current box. It does not promote.

**Rationale:** in a room with names drawn from a dozen orthographies, a great many students go by a name that is a single token of their roster string, and it is frequently not the first token. Token matching therefore captures most of the preferred-name problem with no editing UI at all. But knowing one token is not knowing the name, so it must not advance the card.

### 6.4 The reveal button

There is always a visible "I don't know" button. Pressing it reveals the full name and scores the card as a **Miss**, unconditionally. It must be reachable by keyboard.

---

## 7. Scheduling

Five boxes. Intervals are counted **in sessions, not in days.**

**Rationale:** the real usage pattern is three hammering sessions the night before school starts, twice daily through week one, then tapering. A date-based scheduler responds to the second session of an evening by announcing that nothing is due, which is useless. Session-counted intervals let cramming work while still producing spaced behaviour when the user does space out. The loss of forgetting-curve fidelity is irrelevant when the whole job is about a hundred names over three weeks.

| Box | Interval (sessions) |
| --- | --- |
| 1 | 1 |
| 2 | 2 |
| 3 | 4 |
| 4 | 8 |
| 5 | 16 |

Session mechanics:

- Starting a session increments `sessionCount`.
- A card is due when `dueSession <= sessionCount`.
- **Correct:** `box = min(box + 1, 5)`, then `dueSession = sessionCount + interval[box]`.
- **Partial:** `box` unchanged, `dueSession = sessionCount + interval[box]`.
- **Miss:** `box = 1`, `dueSession = sessionCount + 1`.
- Increment `seen` always; increment `correct` or `missed` as appropriate. Count a Partial as neither.

Session composition:

- Default session length is 20 cards, drawn from due cards, lowest box first, shuffled within box.
- If fewer than 20 cards are due, the session is however many are due. Do not pad with cards that are not due.
- Offer a **Drill everything** mode that ignores due dates and reviews every card in the deck in random order. Box updates apply normally. This is what gets used the night before school starts.

---

## 8. Interface

Four screens. Keep it plain. This is a tool, not a product.

### 8.1 Home

- Drop zone accepting both `.pdf` (import a new class) and `.deck.json` (load an existing deck). Also a file-picker button, because drag-and-drop is not obvious to everyone.
- List of currently loaded decks with card count, number due, and a Study button on each.
- A **Study combined** control that pools two or more selected decks into one session.
- A persistent **unsaved changes** indicator whenever in-memory state differs from what was last written to disk.

Per-class decks are the storage unit; the combined pool is a session-time option. Studying a single class means the deck name itself becomes an unconscious cue that narrows the guess, which is exactly the cue that is not available in a hallway. The combined pool removes it. Both are cheap to support, so support both and let the user find out which they prefer.

### 8.2 Import confirmation

As specified in section 4.3. Editable deck name. Confirm and Cancel.

If a deck with the same name is already loaded, offer a **merge** instead of a replace: match cards by exact name string, preserve `box`, `dueSession`, and statistics for names present in both, add cards for names only in the new PDF at box 1, and list names only in the old deck for the user to keep or delete. This is what happens when a student adds or drops in week two.

### 8.3 Study

- The photo, large and centred. Nothing else competing with it.
- A single text input, autofocused.
- Enter submits. The result appears with the correct spelling. Enter again advances.
- The "I don't know" button, reachable with a key (Space is a reasonable choice when the input is empty; make sure it does not fight with typing).
- Progress indicator: position in session, and a running count of correct, partial, and missed.
- Escape ends the session early. Cards already answered keep their updates.

### 8.4 Session summary

- Counts, plus the list of names missed this session.
- Box distribution across the deck, as a simple bar or a row of counts.
- A prominent **Save deck** action, and a browser-level warning on unload if the deck is unsaved.

**Include a delete control on cards in this screen only.** Students drop, and a deck that keeps drilling a name the user will never say again is actively wrong. This is the sole exception to read-only import in v1, and it is deliberate.

### 8.5 Saving

Attempt the File System Access API (`showSaveFilePicker`) so that saving over an existing deck is one click. Fall back to a generated anchor download when it is unavailable or throws. The agent must verify what actually works at a `file://` origin in Safari and Chrome on macOS — Safari's support for that API is the weaker case and the fallback will likely be the real path there.

Prompt to save at session end. Warn on `beforeunload` when unsaved. Do not autosave silently to browser storage and let the user believe their progress is safe.

---

## 9. Privacy requirements

These are not optional and are not to be traded away for a feature.

- **Student ID numbers are extracted during parsing and immediately discarded.** They are never written to the deck file, never held in memory beyond the parse, and never displayed. A deck that holds a face and a name is a much smaller thing to leak than one that holds a key into a student information system.
- No network requests at runtime. None. The app should function identically with wifi off, and that is the test.
- No telemetry, no analytics, no error reporting.
- Deck files must be written to a location the user chooses, never silently into the folder the app lives in.
- If the project is ever put in a repository, `.gitignore` must block `*.pdf`, `*.deck.json`, and any photo output directory **from the first commit**, before any real data has ever touched the working directory.
- The README opens with the privacy note, above the instructions, in the first screen of text: deck files contain photographs of minors; do not commit them, do not email them, do not put them in shared cloud folders.

The realistic failure mode is not a sophisticated attack. It is a colleague dropping their PDFs into the project folder because that is where the app is, running `git add .` out of habit, and pushing a hundred students' faces to a public repository under their own name. The layout should make that difficult rather than convenient.

---

## 10. README requirements

Written for a teacher, not a developer. No terminal commands anywhere in the quickstart.

1. Privacy note, first.
2. What it does, in three sentences.
3. How to get your PDFs out of PowerSchool: the report is the photo roster at 12 per page. Name each file for the class it contains, because the app uses the filename as the deck name.
4. Quickstart: double-click `name-learner.html`, drag your PDF onto the window, check the names look right, click Confirm, study, click Save at the end, keep the `.deck.json` file somewhere you will find it again.
5. Where your data lives, in one paragraph. Emphasise that the deck file is the only copy and it is the user's job to keep it.
6. Known limits: built against one PowerSchool photo-roster layout, and other export formats will probably not parse.
7. An explicit statement that this is provided as-is with no support.

---

## 11. Acceptance criteria

The build is done when all of these pass.

1. `name-learner.html` opens from a `file://` URL in current macOS Safari and Chrome with no console errors.
2. Dragging a two-page, 22-student PowerSchool photo-roster export onto it yields exactly 22 cards, each with the correct name paired to the correct face, verified by eye against the PDF.
3. No student ID number appears anywhere in the saved deck file. Confirm by searching the JSON for a known ID string.
4. Every photo in the saved deck renders correctly and is recognisably the right student.
5. A typed exact name scores Correct. A one-character typo in a long name scores Correct with the spelling shown. A single token of a multi-token name scores Partial and does not change the box. Gibberish scores Miss and resets the box to 1.
6. Pressing "I don't know" scores a Miss regardless of what is typed in the input.
7. Saving produces a `.deck.json` the user can locate. Reloading that file restores all box states, due sessions, and statistics exactly.
8. Two consecutive sessions in one sitting both present cards, rather than the second announcing that nothing is due.
9. Drill-everything mode presents every card irrespective of scheduling.
10. Two decks can be loaded and studied as a single pooled session, and saving afterwards writes the correct updates back to both files.
11. Re-importing a PDF for an already-loaded deck offers a merge, preserves box state for returning students, and adds new students at box 1.
12. With wifi disabled, everything above still works.
13. Closing the tab with unsaved progress produces a browser warning.

---

## 12. Deliberately deferred

Do not build these now. They are listed so a later version does not have to rediscover them.

- Editing a card's name, adding a preferred name, or attaching a free-text note (pronunciation, seating, "plays cello"). The `preferredName` field exists in the schema and the matcher honours it; only the editor is missing. The user genuinely does not yet know whether this is wanted, and cannot know until after meeting the students, which is after the heaviest drilling is done. Adding it later is a small change and existing decks stay valid.
- Name-to-face direction (show a name, pick the face from a grid).
- Per-student audio for pronunciation.
- Printable contact-sheet export.
- Statistics across time.
