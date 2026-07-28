// ═══════════════════════════════════════════════════════════
// DOC STANDARDIZER — standalone batch version
//
// REQUIRES: the Google Docs API advanced service.
//   Apps Script editor → Services (left rail) → + → Google Docs API → Add
//
// Entry points:
//   standardizeFolder()  — recursively processes every Google Doc
//                          in FOLDER_ID and all subfolders
//   standardizeOneDoc()  — processes the single doc at DOC_URL
//
// Passes applied to each doc (all idempotent — re-runs are safe):
//   1. Smart quotes  (straight → curly)
//   2. Sub/super     (Unicode glyphs like ₂ ³ → real formatting)
//   3. Note Spaces   (paragraphs matching “✎Note Space N:” become
//                     Heading 1 and bold)
//   4. Fonts/colors  (Garamond 12pt body; Montserrat headings;
//                     Title and Note Spaces bold; white text on
//                     dark cells; NON-STANDARD text colors preserved)
//   5. Tables        (strip min row heights and paragraph spacing
//                     in cells — EXCEPT empty rows, which are
//                     treated as answer lines and left alone)
//   6. Borders       (clear paragraph bottom borders doc-wide;
//                     table CELL borders, incl. answer lines,
//                     are a different property and untouched)
// ═══════════════════════════════════════════════════════════

// ── SET THESE ──────────────────────────────────
const FOLDER_ID = 'PASTE_FOLDER_ID_HERE';
const DOC_URL   = 'PASTE_DOC_URL_HERE';    // only for standardizeOneDoc
// ───────────────────────────────────────────────

const CONFIG = {
  bodyFont:    'Garamond',
  bodySize:    12,
  headingFont: 'Montserrat',
  accentColor: '#C97B35',             // Safety Amber (reserved)
  textColor:   '#000000',
  darkThreshold: 128,                 // luminance below this → white text
  headingTypes: [
    DocumentApp.ParagraphHeading.TITLE,
    DocumentApp.ParagraphHeading.HEADING1,
    DocumentApp.ParagraphHeading.HEADING2,
    DocumentApp.ParagraphHeading.HEADING3,
    DocumentApp.ParagraphHeading.HEADING4,
  ],
};

const ALL_PASSES = { quotes: true, subsuper: true, fonts: true, tables: true };

// Paragraphs matching this become Heading 1 + bold.
// Tolerant of stray spacing around the glyph, label, and number.
const NOTE_SPACE_RE = /^\s*\u270E\s*Note\s+Space\s*\d+\s*:/;

// Colors the script considers its own and is free to overwrite.
// Anything else is treated as an intentional author color and preserved.
const MANAGED_COLORS = ['#000000', '#ffffff'];

// ───────────────────────────────────────────────
// SUB / SUPERSCRIPT MAPS
// ───────────────────────────────────────────────

const SUBS = {
  '\u2080':'0','\u2081':'1','\u2082':'2','\u2083':'3','\u2084':'4',
  '\u2085':'5','\u2086':'6','\u2087':'7','\u2088':'8','\u2089':'9',
  '\u208A':'+','\u208B':'-','\u208C':'=','\u208D':'(','\u208E':')',
  '\u2090':'a','\u2091':'e','\u2092':'o','\u2093':'x','\u2094':'\u0259',
  '\u2095':'h','\u2096':'k','\u2097':'l','\u2098':'m','\u2099':'n',
  '\u209A':'p','\u209B':'s','\u209C':'t',
  '\u1D62':'i','\u1D63':'r','\u1D64':'u','\u1D65':'v',
};

const SUPS = {
  '\u2070':'0','\u00B9':'1','\u00B2':'2','\u00B3':'3','\u2074':'4',
  '\u2075':'5','\u2076':'6','\u2077':'7','\u2078':'8','\u2079':'9',
  '\u207A':'+','\u207B':'-','\u207C':'=','\u207D':'(','\u207E':')',
  '\u2071':'i','\u207F':'n',
  '\u1D43':'a','\u1D47':'b','\u1D9C':'c','\u1D48':'d','\u1D49':'e',
  '\u1DA0':'f','\u1D4D':'g','\u02B0':'h','\u02B2':'j',
  '\u1D4F':'k','\u02E1':'l','\u1D50':'m','\u1D52':'o',
  '\u1D56':'p','\u02B3':'r','\u02E2':'s','\u1D57':'t','\u1D58':'u',
  '\u1D5B':'v','\u02B7':'w','\u02E3':'x','\u02B8':'y','\u1DBB':'z',
  '\u1D2C':'A','\u1D2E':'B','\u1D30':'D','\u1D31':'E','\u1D33':'G',
  '\u1D34':'H','\u1D35':'I','\u1D36':'J','\u1D37':'K','\u1D38':'L',
  '\u1D39':'M','\u1D3A':'N','\u1D3C':'O','\u1D3E':'P','\u1D3F':'R',
  '\u1D40':'T','\u1D41':'U','\u2C7D':'V','\u1D42':'W',
};

// ───────────────────────────────────────────────
// ENTRY POINTS
// ───────────────────────────────────────────────

function standardizeFolder() {
  const root = DriveApp.getFolderById(FOLDER_ID);
  Logger.log('Starting batch in: ' + root.getName());
  const count = processFolder(root, 0);
  Logger.log('Processed ' + count + ' docs total.');
}

function standardizeOneDoc() {
  const doc = DocumentApp.openByUrl(DOC_URL);
  const id = doc.getId();
  const name = doc.getName();
  walkDoc(doc, ALL_PASSES);
  doc.saveAndClose();
  clearParagraphBorders(id);
  Logger.log('Done: ' + name);
}

// ───────────────────────────────────────────────
// RECURSIVE FOLDER WALK
// Google Docs only; other file types skipped automatically.
// Per-file try/catch so one bad doc doesn't kill the batch.
// Border cleanup runs after saveAndClose — the two APIs must
// not touch the same doc while it's open.
// ───────────────────────────────────────────────

function processFolder(folder, depth) {
  if (depth > 10) return 0;

  let n = 0;

  const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
  while (files.hasNext()) {
    const file = files.next();
    try {
      const doc = DocumentApp.openById(file.getId());
      walkDoc(doc, ALL_PASSES);
      doc.saveAndClose();
      clearParagraphBorders(file.getId());
      n++;
      Logger.log('Done: ' + file.getName());
    } catch (e) {
      Logger.log('SKIPPED (error): ' + file.getName() + ' — ' + e.message);
    }
  }

  const subs = folder.getFolders();
  while (subs.hasNext()) {
    n += processFolder(subs.next(), depth + 1);
  }

  return n;
}

// ───────────────────────────────────────────────
// PARAGRAPH BORDER CLEANUP (Docs Advanced Service)
// Clears bottom borders on all paragraphs in the doc body,
// including paragraphs inside table cells. Table CELL borders
// (answer lines) are a different property and untouched.
// ───────────────────────────────────────────────

function clearParagraphBorders(docId) {
  const doc = Docs.Documents.get(docId);
  const content = doc.body.content;
  const endIndex = content[content.length - 1].endIndex;

  const noBorder = {
    color:  { color: { rgbColor: {} } },
    width:  { magnitude: 0, unit: 'PT' },
    padding:{ magnitude: 0, unit: 'PT' },
    dashStyle: 'SOLID',
  };

  Docs.Documents.batchUpdate({
    requests: [{
      updateParagraphStyle: {
        range: { startIndex: 1, endIndex: endIndex - 1 },
        paragraphStyle: { borderBottom: noBorder },
        fields: 'borderBottom',
      },
    }],
  }, docId);
}

// ───────────────────────────────────────────────
// DOC WALK — threads textColor down so dark cells get white text
// ───────────────────────────────────────────────

function walkDoc(doc, passes) {
  styleContainer(doc.getBody(), passes, CONFIG.textColor);
  const header = doc.getHeader();
  const footer = doc.getFooter();
  if (header) styleContainer(header, passes, CONFIG.textColor);
  if (footer) styleContainer(footer, passes, CONFIG.textColor);
  doc.getFootnotes().forEach(fn =>
    styleContainer(fn.getFootnoteContents(), passes, CONFIG.textColor));
}

function styleContainer(container, passes, textColor) {
  const n = container.getNumChildren();
  for (let i = 0; i < n; i++) {
    styleElement(container.getChild(i), passes, textColor);
  }
}

function styleElement(el, passes, textColor) {
  const type = el.getType();

  if (type === DocumentApp.ElementType.PARAGRAPH) {
    const p = el.asParagraph();
    processPara(p, p.getHeading(), passes, textColor);

  } else if (type === DocumentApp.ElementType.LIST_ITEM) {
    processPara(el.asListItem(), DocumentApp.ParagraphHeading.NORMAL, passes, textColor);

  } else if (type === DocumentApp.ElementType.TABLE) {
    const table = el.asTable();
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);

      if (passes.tables && !isEmptyRow(row)) {
        cleanRowSpacing(row);
      }

      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        styleContainer(cell, passes, pickTextColor(cell.getBackgroundColor()));
      }
    }
  }
}

function pickTextColor(bgHex) {
  if (!bgHex) return CONFIG.textColor;
  const m = /^#?([0-9a-f]{6})$/i.exec(bgHex);
  if (!m) return CONFIG.textColor;
  const n = parseInt(m[1], 16);
  const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return lum < CONFIG.darkThreshold ? '#FFFFFF' : CONFIG.textColor;
}

// ───────────────────────────────────────────────
// TABLE SPACING CLEANUP
// Empty rows (all cells whitespace-only) are answer lines: skipped.
// Other rows lose min height and in-cell paragraph spacing.
// Cell padding is untouched.
// ───────────────────────────────────────────────

function isEmptyRow(row) {
  for (let c = 0; c < row.getNumCells(); c++) {
    if (row.getCell(c).getText().trim() !== '') return false;
  }
  return true;
}

function cleanRowSpacing(row) {
  row.setMinimumHeight(0);
  for (let c = 0; c < row.getNumCells(); c++) {
    const cell = row.getCell(c);
    const n = cell.getNumChildren();
    for (let i = 0; i < n; i++) {
      const child = cell.getChild(i);
      const t = child.getType();
      if (t === DocumentApp.ElementType.PARAGRAPH ||
          t === DocumentApp.ElementType.LIST_ITEM) {
        const a = {};
        a[DocumentApp.Attribute.SPACING_BEFORE] = 0;
        a[DocumentApp.Attribute.SPACING_AFTER]  = 0;
        child.setAttributes(a);   // element-level form: paragraph attributes
      }
    }
  }
}

// ───────────────────────────────────────────────
// PER-PARAGRAPH: quotes → sub/super → note space → fonts
// Note Space detection promotes the paragraph to Heading 1 before
// the font pass runs, so it picks up the heading font automatically.
// ───────────────────────────────────────────────

function processPara(para, heading, passes, textColor) {
  if (passes.quotes)   smartQuotesInPara(para);
  if (passes.subsuper) fixSubSuper(para);

  if (passes.fonts) {
    let forceBold = false;

    if (isNoteSpace(para)) {
      if (typeof para.setHeading === 'function') {
        para.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      }
      heading = DocumentApp.ParagraphHeading.HEADING1;
      forceBold = true;
    }

    applyFonts(para, heading, textColor, forceBold);
  }
}

function isNoteSpace(para) {
  return NOTE_SPACE_RE.test(para.getText());
}

// ───────────────────────────────────────────────
// SMART QUOTES
// ───────────────────────────────────────────────

function smartQuotesInPara(para) {
  const textEl = para.editAsText();
  const s = textEl.getText();
  if (s.indexOf('"') === -1 && s.indexOf("'") === -1) return;
  const out = curlify(s);
  if (out !== s) textEl.setText(out);
}

function curlify(s) {
  let r = s;
  r = r.replace(/(^|[\s([{<\u2018\u201C])"/g, '$1\u201C');
  r = r.replace(/"/g, '\u201D');
  r = r.replace(/(^|[\s([{<\u2018\u201C])'/g, '$1\u2018');
  r = r.replace(/'/g, '\u2019');
  return r;
}

// ───────────────────────────────────────────────
// UNICODE SUB/SUPERSCRIPT → REAL FORMATTING
// Uses setTextAlignment with the TextAlignment enum (the correct API).
// ───────────────────────────────────────────────

function fixSubSuper(para) {
  const textEl = para.editAsText();
  const s = textEl.getText();

  let out = '';
  const marks = [];  // { i: finalIndex, sub: bool }
  for (let k = 0; k < s.length; k++) {
    const ch = s[k];
    if (SUBS[ch])      { marks.push({ i: out.length, sub: true  }); out += SUBS[ch]; }
    else if (SUPS[ch]) { marks.push({ i: out.length, sub: false }); out += SUPS[ch]; }
    else               { out += ch; }
  }
  if (marks.length === 0) return;

  textEl.setText(out);

  for (const mk of marks) {
    textEl.setTextAlignment(mk.i, mk.i,
        mk.sub ? DocumentApp.TextAlignment.SUBSCRIPT
               : DocumentApp.TextAlignment.SUPERSCRIPT);
  }
}

// ───────────────────────────────────────────────
// FONTS + COLOR
// Title and Note Spaces → Montserrat bold.
// Headings 1–4 → Montserrat (bold untouched unless a Note Space).
// Body → Garamond 12pt.
// Color is applied separately so author colors survive.
// ───────────────────────────────────────────────

function applyFonts(para, heading, textColor, forceBold) {
  const textEl = para.editAsText();
  const len = textEl.getText().length;
  if (len === 0) return;

  const isTitle   = heading === DocumentApp.ParagraphHeading.TITLE;
  const isHeading = CONFIG.headingTypes.indexOf(heading) !== -1;

  const attrs = {};
  attrs[DocumentApp.Attribute.FONT_FAMILY] =
      isHeading ? CONFIG.headingFont : CONFIG.bodyFont;
  if (isTitle || forceBold) {
    attrs[DocumentApp.Attribute.BOLD] = true;
  }
  if (!isHeading) {
    attrs[DocumentApp.Attribute.FONT_SIZE] = CONFIG.bodySize;
  }
  textEl.setAttributes(0, len - 1, attrs);   // no color here

  applyTextColor(textEl, len, textColor || CONFIG.textColor);
}

// Recolor only the stretches that currently sit at a managed color
// (black or white). Author-applied colors are left as-is.
function applyTextColor(textEl, len, target) {
  const uniform = textEl.getForegroundColor();   // null if mixed

  if (uniform !== null) {
    if (isManagedColor(uniform)) {
      textEl.setForegroundColor(0, len - 1, target);
    }
    return;
  }

  // Mixed paragraph: walk it, group into runs of identical color,
  // and recolor only the managed ones.
  let runStart = 0;
  let runColor = textEl.getForegroundColor(0);

  for (let i = 1; i <= len; i++) {
    const here = (i < len) ? textEl.getForegroundColor(i) : '\u0000';
    if (here !== runColor) {
      if (isManagedColor(runColor)) {
        textEl.setForegroundColor(runStart, i - 1, target);
      }
      runStart = i;
      runColor = here;
    }
  }
}

function isManagedColor(hex) {
  if (!hex) return true;   // unset / inherited counts as standard
  return MANAGED_COLORS.indexOf(hex.toLowerCase()) !== -1;
}
