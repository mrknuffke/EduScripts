#!/usr/bin/env node
/**
 * build_docs.js: generate every .docx for a PE cluster from content.json.
 *
 *   node build_docs.js content.json out/ --draft    teacher/PLC review draft
 *   node build_docs.js content.json out/ --final    Student, Answer Key, Scoring Sheet
 *   add --force to build despite checker errors (never for delivery)
 *
 * Runs check_cluster.py first and refuses to build while it reports errors.
 * Stimulus images named in stimulus.blocks[].image are embedded if found next to content.json.
 */
const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, AlignmentType, HeightRule, PageBreak, ImageRun, TableLayoutType,
} = require("docx");

// ------------------------------------------------------------------ args
const args = process.argv.slice(2);
if (args.length < 2 || !(args.includes("--draft") || args.includes("--final"))) {
  console.error("usage: node build_docs.js content.json out/ --draft|--final [--force]");
  process.exit(2);
}
const SPEC_PATH = path.resolve(args[0]);
const OUT = path.resolve(args[1]);
const MODE = args.includes("--draft") ? "draft" : "final";
const FORCE = args.includes("--force");
const spec = JSON.parse(fs.readFileSync(SPEC_PATH, "utf8"));
fs.mkdirSync(OUT, { recursive: true });

// ------------------------------------------------------------------ checker
const chk = spawnSync("python3", [path.join(__dirname, "check_cluster.py"), SPEC_PATH, "--json"], { encoding: "utf8" });
let CHECK;
try { CHECK = JSON.parse(chk.stdout); } catch (e) {
  console.error("check_cluster.py did not return JSON:\n" + chk.stdout + chk.stderr);
  process.exit(1);
}
if (CHECK.errors.length && !FORCE) {
  console.error(`Checker reports ${CHECK.errors.length} error(s). Fix them first:\n  - ` + CHECK.errors.join("\n  - "));
  process.exit(1);
}

// ------------------------------------------------------------------ design system
const meta = spec.meta;
const fmt = Object.assign({ paper: "A4", body_font: "Garamond", heading_font: "Montserrat Medium", accent: "4A90A4" }, meta.format || {});
const PAGE = fmt.paper === "Letter" ? { w: 12240, h: 15840 } : { w: 11906, h: 16838 };
const MAR = { lr: 720, t: 720, b: 1008 };
const CW = PAGE.w - 2 * MAR.lr;             // content width, DXA
const BODY = fmt.body_font, HEAD = fmt.heading_font;
const SZ = 24;                               // 12 pt
const ACCENT = fmt.accent.replace("#", "");
const ACCENT_LT = tint(ACCENT, 0.85);
const NAVY = "1A1F2E", GREY = "F2F2F2", MUTED = "666666";
const BUCKET_FILLS = ["EAF4FB", "E8F5E9", "FFF8E1", "F3E5F5", "E0F2F1", "FBE9E7", "ECEFF1", "F1F8E9"];
const LABELS = meta.level_labels || ["Not Yet Evident", "Emerging", "Developing", "Meeting", "Meeting with Distinction"];
const L = "ABCDEFGH";

function tint(hex, f) {
  const n = parseInt(hex, 16);
  const c = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map(v => Math.round(v + (255 - v) * f));
  return c.map(v => v.toString(16).padStart(2, "0")).join("").toUpperCase();
}

// ------------------------------------------------------------------ primitives
const run = (text, o = {}) => new TextRun({ text, font: o.font || BODY, size: o.size || SZ, bold: o.bold, italics: o.italics, color: o.color });
function para(content, o = {}) {
  const children = (Array.isArray(content) ? content : [content]).map(c => typeof c === "string" ? run(c, o) : c);
  return new Paragraph({
    children, alignment: o.align,
    spacing: { before: o.before ?? 0, after: o.after ?? 120, line: o.line },
    indent: o.indent, keepNext: o.keepNext, keepLines: o.keepLines,
  });
}
const h1 = t => para([run(t, { font: HEAD, size: 36 })], { after: 80 });
const h2 = t => para([run(t, { font: HEAD, size: 28, color: ACCENT })], { before: 240, after: 120, keepNext: true });
const h3 = t => para([run(t, { font: HEAD, size: 24 })], { before: 160, after: 80, keepNext: true });
const small = (t, o = {}) => para([run(t, { size: 20, color: MUTED, italics: o.italics })], { after: o.after ?? 80 });
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });
const divider = () => new Paragraph({ children: [], spacing: { before: 120, after: 120 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 1 } } });
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const THIN = { style: BorderStyle.SINGLE, size: 4, color: "BBBBBB" };

function nameLine() {
  return para([
    run("Name:\u00A0", { bold: true }), run("________________________________"),
    run("\u00A0\u00A0\u00A0Block:\u00A0", { bold: true }), run("________"),
    run("\u00A0\u00A0\u00A0Date:\u00A0", { bold: true }), run("____________"),
  ], { after: 120 });
}

function cell(content, width, o = {}) {
  const kids = (Array.isArray(content) ? content : [content]).map(c =>
    typeof c === "string" ? para([run(c, { size: o.size, bold: o.bold, color: o.color, font: o.font })], { after: o.size && o.size < 20 ? 0 : 40 }) : c);
  return new TableCell({
    children: kids.length ? kids : [para("", { after: 0 })],
    width: { size: width, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: o.pad ?? (o.size && o.size < 20 ? 30 : 60), bottom: o.pad ?? (o.size && o.size < 20 ? 30 : 60), left: 100, right: 100 },
    borders: o.borders || { top: THIN, bottom: THIN, left: THIN, right: THIN },
    columnSpan: o.span,
  });
}

/** Generic table. widths are proportions; header row optional. */
function table(rows, props, o = {}) {
  const total = props.reduce((a, b) => a + b, 0);
  const widths = props.map(p => Math.floor(CW * p / total));
  widths[widths.length - 1] += CW - widths.reduce((a, b) => a + b, 0);
  const size = o.size || SZ;
  const trs = rows.map((r, ri) => {
    const isHead = o.header && ri === 0;
    const rowFill = (Array.isArray(r) ? undefined : r.fill) || (isHead ? NAVY : undefined);
    const cells = (r.cells || r).map((c, ci) => {
      if (c && c.span) {
        const w = widths.slice(ci, ci + c.span).reduce((a, b) => a + b, 0);
        return cell(c.text, w, { size, span: c.span, bold: c.bold ?? isHead, fill: c.fill || rowFill, color: isHead || c.fill === NAVY ? "FFFFFF" : undefined });
      }
      return cell(c, widths[ci], { size, bold: isHead, fill: rowFill, color: isHead ? "FFFFFF" : undefined });
    });
    return new TableRow({ children: cells, tableHeader: isHead, cantSplit: true });
  });
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: widths, rows: trs, layout: TableLayoutType.FIXED });
}

function box(children, fill = ACCENT_LT, border = ACCENT) {
  const b = { style: BorderStyle.SINGLE, size: 6, color: border };
  return new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], layout: TableLayoutType.FIXED,
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      children, width: { size: CW, type: WidthType.DXA },
      shading: { fill, type: ShadingType.CLEAR, color: "auto" },
      margins: { top: 140, bottom: 140, left: 180, right: 180 },
      borders: { top: b, bottom: b, left: b, right: b },
    })] })],
  });
}

function answerLines(count) {
  const row = () => new TableRow({
    height: { value: 550, rule: HeightRule.ATLEAST }, cantSplit: true,
    children: [new TableCell({
      width: { size: CW, type: WidthType.DXA },
      borders: { top: NONE, left: NONE, right: NONE, bottom: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" } },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [run(" ")] })],
    })],
  });
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [CW], layout: TableLayoutType.FIXED, rows: Array.from({ length: count }, row) });
}

// ------------------------------------------------------------------ content helpers
const crits = spec.criteria.slice();
const byOrder = crits.slice().sort((a, b) => a.order - b.order);
const bucketNames = spec.buckets.map(b => b.name);
const bucketFill = n => BUCKET_FILLS[bucketNames.indexOf(n) % BUCKET_FILLS.length];
const feats = Object.fromEntries(spec.features.map(f => [f.id, f]));
const primary = c => c.buckets[0];
const esOf = b => crits.filter(c => !c.beyond && c.buckets.includes(b));
const probeOf = b => crits.find(c => c.beyond && c.buckets.includes(b));
const critOrderBucket = () => {
  const out = [];
  bucketNames.forEach(b => crits.filter(c => primary(c) === b && !c.beyond).forEach(c => out.push(c)));
  return out;
};
const keyLetter = c => L[(c.choices || []).findIndex(o => o.key)];
const keyText = c => (c.choices || []).find(o => o.key)?.text || "";
const slug = meta.slug || meta.standard.toLowerCase();
const mode = meta.distinction_mode || "probe";
const topGate = !!meta.mwd_requires_top_gate;

function level4Rule(b) {
  const g = CHECK.gates[b];
  if (mode === "probe") return `${topGate ? g.top : g.meeting}+ and ${probeOf(b)?.id || "probe"} earned`;
  if (mode === "perfect") return `${g.n} of ${g.n} (full raw score awards 4, deliberate for this standard)`;
  return "Not awarded on this instrument";
}

function imageFor(block) {
  if (!block.image) return null;
  const p = path.join(path.dirname(SPEC_PATH), block.image);
  if (!fs.existsSync(p)) { console.warn(`image ${block.image} not found; using the table/text version`); return null; }
  const buf = fs.readFileSync(p);
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);   // PNG IHDR
  const maxPx = Math.floor(CW / 1440 * 96);
  const scale = Math.min(1, maxPx / w);
  return new Paragraph({ children: [new ImageRun({ type: "png", data: buf, transformation: { width: Math.round(w * scale), height: Math.round(h * scale) } })], spacing: { after: 120 } });
}

// ------------------------------------------------------------------ shared sections
function titleBlock(subtitle) {
  return [
    h1(meta.title),
    para([run(`${meta.standard} · ${meta.unit || ""}`, { size: 20, color: MUTED })], { after: 40 }),
    subtitle ? para([run(subtitle, { font: HEAD, size: 24, color: ACCENT })], { after: 80 }) : null,
    divider(),
  ].filter(Boolean);
}

function stimulusSection() {
  const s = spec.stimulus, out = [];
  out.push(h2("The situation"));
  s.context.forEach(t => out.push(para(t)));
  s.blocks.forEach(b => {
    out.push(h3(b.title));
    const img = imageFor(b);
    if (img) out.push(img);
    else if (b.kind === "table") {
      const props = [2.2].concat(b.columns.slice(1).map(() => 1));
      out.push(table([b.columns].concat(b.rows), props, { header: true }));
    } else (b.text || []).forEach(t => out.push(para(t)));
    if (b.note) out.push(small(b.note, { italics: true }));
    if (b.source) out.push(small("Source: " + b.source, { italics: true }));
  });
  if (s.glossary && s.glossary.length) {
    out.push(para("", { after: 80 }));
    out.push(box([para([run("Terms used in this assessment", { font: HEAD, size: 22 })], { after: 80 })]
      .concat(s.glossary.map(g => para([run(g.term + ": ", { bold: true }), run(g.def)], { after: 40 }))), GREY, "BBBBBB"));
  }
  return out;
}

function itemBlock(c, withKey) {
  const out = [];
  out.push(para([run(`${c.item}. `, { bold: true }), run(c.stem)], { before: 200, after: 60, keepNext: true, keepLines: true }));
  if (c.help) out.push(para([run(c.help, { italics: true, color: MUTED, size: 22 })], { after: 80, keepNext: true }));
  if (c.type === "mc") {
    c.choices.forEach((o, i) => {
      const mark = withKey && o.key ? "  ✓ key" : "";
      out.push(para([run(`${L[i]}.  `, { bold: true }), run(o.text), run(mark, { bold: true, color: "2E7D32" })],
        { indent: { left: 567, hanging: 340 }, after: 40, keepNext: i < c.choices.length - 1 }));
      if (withKey && !o.key) out.push(para([run(`misconception: ${o.misconception}`, { size: 20, italics: true, color: MUTED })], { indent: { left: 567 }, after: 40 }));
    });
  } else if (!withKey) {
    out.push(answerLines(c.lines || Math.ceil((c.sentence_cap || 3) * 1.5) + 1));
  }
  return out;
}

function partsWithItems(withKey, extra) {
  const out = [];
  spec.parts.forEach((p, pi) => {
    if (pi === 0 && !withKey) out.push(pageBreak());
    else if (pi > 0) out.push(para("", { after: 120 }));
    out.push(box([para([run(p.title, { font: HEAD, size: 26 })], { after: 40 }), p.help ? para([run(p.help, { italics: true })], { after: 0 }) : null].filter(Boolean)));
    byOrder.filter(c => c.part === p.n).forEach(c => {
      out.push(...itemBlock(c, withKey));
      if (extra) out.push(...extra(c));
    });
  });
  return out;
}

function gateRows(withScoreCols) {
  const head = ["Bucket", "ES items", `${LABELS[1]} 1`, `${LABELS[2]} 2`, `${LABELS[3]} 3`, `${LABELS[4]} 4`];
  if (withScoreCols) head.push("Raw", "Level");
  const rows = [head];
  bucketNames.forEach(b => {
    const g = CHECK.gates[b];
    const r = [b, String(g.n), `${g.emerging}+`, `${g.developing}+`, `${g.meeting}+`, level4Rule(b)];
    if (withScoreCols) r.push("", "");
    rows.push({ cells: r, fill: bucketFill(b) });
  });
  return rows;
}

function standingBlocks(size) {
  const o = { size };
  const out = [
    para([run("Follow-through. ", { bold: true, size }), run("A student who reaches a wrong value early and reasons correctly from it loses the criterion where the error occurred, not every criterion downstream.", o)], { after: 60 }),
    para([run("Distinction modifies, it never shortcuts. ", { bold: true, size }), run(mode === "probe" ? "A student who earns a probe but misses the Meeting gate takes their gate level. Probes are scored 0 or 1 and sit outside the raw score." : mode === "perfect" ? "Level 4 is awarded only for a full raw score in that bucket." : "This instrument reports levels up to 3.", o)], { after: 60 }),
    para([run(`Raw 0 is ${LABELS[0]}. `, { bold: true, size }), run("A cross-scored criterion counts toward every bucket it is tagged to.", o)], { after: 60 }),
  ];
  if (meta.levels_scope === "per-period") {
    out.push(para([run("Accumulation. ", { bold: true, size }), run("Bucket levels accumulate across the reporting period. A level earned here does not overwrite a higher level already earned for the same bucket elsewhere.", o)], { after: 60 }));
  }
  return out;
}

function reflection(size) {
  return [para([run("Grading reflection", { font: HEAD, size: size + 4 })], { before: 120, after: 60, keepNext: true })]
    .concat(spec.buckets.map(b => para([run(b.name + ". ", { bold: true, size }), run(b.reflection, { size })], { after: 60 })));
}

// ------------------------------------------------------------------ documents
function studentDoc() {
  const info = box([
    para([run("What this assessment asks", { font: HEAD, size: 24 })], { after: 60 }),
    para("Read the situation and the data carefully. Then answer every question using what you read."),
    para([run("Resources: ", { bold: true }), run(meta.resources || "")], { after: 40 }),
    para([run("Suggested time: ", { bold: true }), run(`${meta.duration_min || 40} minutes.`)], { after: 0 }),
  ]);
  return [...titleBlock(), nameLine(), info, ...stimulusSection(), ...partsWithItems(false)];
}

function answerKeyDoc() {
  const out = [...titleBlock("Answer Key"), ...standingBlocks(SZ).slice(0, 1)];
  byOrder.forEach(c => {
    const tags = `${c.id} · ${c.buckets.join(", ")} · ${c.skill}${c.beyond ? " · Distinction probe" : ""}`;
    out.push(para([run(`${c.item}. `, { bold: true }), run(c.stem)], { before: 200, after: 40, keepNext: true }));
    if (c.type === "mc") out.push(para([run("Key: ", { bold: true }), run(`${keyLetter(c)}. ${keyText(c)}`)], { after: 40, keepNext: true }));
    out.push(para([run("What earns the point: ", { bold: true }), run(c.earns)], { after: 40, keepNext: true }));
    out.push(small(tags));
  });
  return out;
}

function scoringDoc() {
  const S = 17; // 8.5 pt working-document exception
  const rows = [["ID", "Judgment", "Bucket", "Code", "Item", "Mode", "0/1"]];
  bucketNames.forEach(b => {
    rows.push({ cells: [{ text: spec.buckets.find(x => x.name === b).label, span: 7, fill: bucketFill(b) }] });
    crits.filter(c => primary(c) === b && !c.beyond).forEach(c =>
      rows.push([c.id, c.cue, c.buckets.join(" + "), c.skill, c.item, c.type === "mc" ? "auto" : "hand", ""]));
  });
  if (mode === "probe") {
    rows.push({ cells: [{ text: "Distinction probes (outside the raw score)", span: 7, fill: GREY }] });
    bucketNames.forEach(b => { const p = probeOf(b); if (p) rows.push([p.id, p.cue, b, p.skill, p.item, "hand", ""]); });
  }
  const out = [...titleBlock("Scoring Sheet"), nameLine(),
    table(rows, [0.6, 3.4, 1.1, 0.9, 0.5, 0.6, 0.5], { header: true, size: S }),
    para([run("Rollup and gates", { font: HEAD, size: 22 })], { before: 160, after: 60 }),
    table(gateRows(true), [0.8, 0.6, 0.9, 0.9, 0.9, 2.2, 0.5, 0.6], { header: true, size: S }),
    para("", { after: 60 }),
    ...standingBlocks(S), ...reflection(S),
    small(`${crits.length} criteria · ${CHECK.n_es} evidence-statement · ${CHECK.n_probes} probes · ${bucketNames.length} buckets. Full scoring language is in the Answer Key.`, { italics: true }),
  ];
  return out;
}

function draftDoc() {
  const out = [...titleBlock("Review Draft · not for students")];
  out.push(table([
    ["Standard", `${meta.standard}. ${meta.standard_text}`],
    ["Course / unit", `${meta.course} · ${meta.unit}`],
    ["Buckets scored", spec.buckets.map(b => `${b.name} (${CHECK.counts[b.name].es})`).join(", ")],
    ["Criteria", `${crits.length} (${CHECK.n_es} evidence-statement, ${CHECK.n_probes} probes)`],
    ["Delivery / time", `${meta.delivery} · ${meta.duration_min || 40} minutes`],
    ["Distinction", mode === "probe" ? `probe per bucket; level 4 needs ${topGate ? "the top gate" : "the Meeting gate"} plus the probe` : mode],
  ], [1, 4], { size: 20 }));
  out.push(small("Read the evidence statement coverage table first: it is what the instrument is built from."));

  out.push(pageBreak(), h2("1. What students see: stimulus"), ...stimulusSection());
  out.push(pageBreak(), h2("2. Items in student order, with keys"));
  out.push(...partsWithItems(true, c => [small(`${c.id} · ${c.buckets.join(" + ")} · ${c.skill} · ${c.beyond ? "probe extending " + c.extends : "feature " + c.feature}${c.compound ? " · compound" : ""}${c.or_pathway ? " · OR-pathway" : ""}`)]));

  out.push(pageBreak(), h2("3. Evidence statement coverage"));
  const cov = [["Ref", "Observable feature", "Code", "Criteria"]];
  spec.features.forEach(f => cov.push(f.in_scope
    ? [f.ref, f.text, f.skill, (CHECK.coverage[f.id] || []).join(", ") || "NOT ASSESSED"]
    : { cells: [f.ref, f.text, "out", "Excluded: " + f.exclusion_reason], fill: GREY }));
  out.push(table(cov, [0.6, 4, 0.8, 1.6], { header: true, size: 20 }));

  if (mode === "probe") {
    out.push(h2("4. Distinction probes"));
    const pr = [["Bucket", "Probe", "Item", "Extends", "What earns the point"]];
    bucketNames.forEach(b => { const p = probeOf(b); if (p) pr.push([b, p.id, p.item, p.extends, p.earns]); });
    out.push(table(pr, [0.9, 0.8, 0.7, 0.9, 3.7], { header: true, size: 20 }));
  }

  out.push(h2("5. Criteria in bucket order"));
  const cr = [["ID", "Criterion", "Bucket", "Code", "Item", "Mode", "What earns the point"]];
  bucketNames.forEach(b => {
    cr.push({ cells: [{ text: spec.buckets.find(x => x.name === b).label, span: 7, fill: bucketFill(b) }] });
    crits.filter(c => primary(c) === b).forEach(c => cr.push([c.id + (c.beyond ? " (probe)" : ""), c.criterion, c.buckets.join(" + "), c.skill, c.item, c.type === "mc" ? "auto" : "hand", c.earns]));
  });
  out.push(table(cr, [0.9, 2.5, 0.9, 1.0, 0.65, 0.75, 2.5], { header: true, size: 18 }));

  out.push(h2("6. Gates"));
  out.push(table(gateRows(false), [0.9, 0.7, 1, 1, 1, 2.2], { header: true, size: 20 }));
  out.push(...standingBlocks(20));
  out.push(...reflection(20));

  if (spec.screening && spec.screening.length) {
    out.push(pageBreak(), h2("7. 3D screening report"));
    const sc = [["#", "Item", "Rating", "Evidence", "Change offered"]];
    spec.screening.forEach(s => sc.push([String(s.n), s.item, s.rating, s.evidence, s.change || ""]));
    out.push(table(sc, [0.3, 1.2, 0.7, 2.4, 2.2], { header: true, size: 18 }));
  }

  out.push(h2("8. Open decisions"));
  out.push(small("Each is a question with a default. Mark up any you want changed."));
  out.push(small(`Checker: ${CHECK.errors.length} errors, ${CHECK.warnings.length} warnings. Key positions ${Object.entries(CHECK.key_positions).map(([k, v]) => `${k}=${v}`).join(", ")}.`, { italics: true }));
  let n = 1;
  (spec.open_decisions || []).forEach(d => out.push(para([run(`${n++}. `, { bold: true }), run(d.q + " "), run(`Default: ${d.default}`, { italics: true })], { after: 60 })));
  if (spec.held_back_terms?.length) {
    out.push(para([run(`${n++}. `, { bold: true }), run("Held-back terms, not used in any item or key: " + spec.held_back_terms.map(h => `${h.term} (${h.substitute})`).join("; ") + ". "), run("Default: correct.", { italics: true })], { after: 60 }));
  }
  CHECK.warnings.forEach(w => out.push(para([run(`${n++}. `, { bold: true }), run(w + " "), run("Default: keep as drafted.", { italics: true })], { after: 60 })));
  return out;
}

// ------------------------------------------------------------------ write
function makeDoc(children) {
  return new Document({
    styles: { default: { document: { run: { font: BODY, size: SZ } } } },
    sections: [{ properties: { page: { size: { width: PAGE.w, height: PAGE.h }, margin: { top: MAR.t, bottom: MAR.b, left: MAR.lr, right: MAR.lr } } }, children }],
  });
}
async function write(name, children) {
  const f = path.join(OUT, `${slug}-${name}.docx`);
  fs.writeFileSync(f, await Packer.toBuffer(makeDoc(children)));
  console.log("wrote " + f);
}

(async () => {
  if (CHECK.errors.length) console.warn(`WARNING: building with ${CHECK.errors.length} checker error(s) because --force was given.`);
  if (MODE === "draft") await write("DRAFT-review", draftDoc());
  else {
    await write("Student", studentDoc());
    await write("Answer-Key", answerKeyDoc());
    await write("Scoring-Sheet", scoringDoc());
  }
})();
