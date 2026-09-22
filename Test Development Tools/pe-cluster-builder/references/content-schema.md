# content.json: the single source of truth

Every output (review draft, student paper, answer key, scoring sheet, Apps Script) is generated
from this one file. Change content here and regenerate; never hand-edit an output. A worked example
is in `assets/example-content.json`.

All text fields are plain Unicode: curly quotes, Unicode sub/superscripts (CO₂, 10⁻³), no
em-dashes, no LaTeX, no Markdown.

## Top level

```json
{
  "meta": {},
  "scope": [],
  "held_back_terms": [],
  "features": [],
  "buckets": [],
  "stimulus": {},
  "parts": [],
  "criteria": [],
  "derived_checks": [],
  "screening": [],
  "open_decisions": []
}
```

## meta

| Field | Type | Notes |
|---|---|---|
| `standard` | string | e.g. "HS-LS2-3". Exactly one. |
| `standard_text` | string | The PE text, with clarification statement and assessment boundary. |
| `framework` | string | "NGSS", "AP Biology CED", "IB DP Physics", a state framework, etc. |
| `course`, `unit` | string | |
| `title` | string | Short context title, e.g. "Two Compost Bins". |
| `slug` | string | Lowercase, hyphens. Used in file names. |
| `delivery` | "paper" or "form" | |
| `duration_min` | number | Default 40. |
| `resources` | string | The allowed-resources line printed on the paper and Form. |
| `distinction_mode` | "probe", "perfect", "none" | From the profile. |
| `mwd_requires_top_gate` | boolean | Default false. |
| `min_items_per_bucket` | number | From the profile. |
| `levels_scope` | "per-assessment" or "per-period" | Controls the Accumulation block. |
| `gate_proportions` | object | `{"developing":0.45,"meeting":0.70,"top":0.85}` by default. |
| `sentence_cap_words` | number | Default 25. |
| `level_labels` | array of 5 strings | Default ["Not Yet Evident","Emerging","Developing","Meeting","Meeting with Distinction"]. |
| `format` | object | `paper` ("A4" or "Letter"), `body_font`, `heading_font`, `accent` (hex, no #), `spelling` ("UK" or "US"). |
| `form` | object | Form route only: `teachers` (array), `access_code`, `access_code_help`, `collect_emails`. |

## scope

`[{ "code": "LS2-3.1", "text": "Explain that ..." }]` The course's taught-skill list filtered to
this standard. Codes come from the course's registry, never from memory.

## held_back_terms

`[{ "term": "chemosynthesis", "substitute": "microbes that use chemical energy from vent water" }]`
The checker fails any student-facing text or key that uses a held-back term.

## features

```json
{ "id": "F1", "ref": "1a-i", "text": "Energy from photosynthesis and respiration drives ...",
  "in_scope": true, "skill": "LS2-3.1", "exclusion_reason": "" }
```

Every observable feature row in the evidence statement, in or out. Out-of-scope rows set
`in_scope: false` and give `exclusion_reason`. They stay in the file so the draft shows what was
excluded and why.

## buckets

```json
{ "name": "DCI", "label": "DCI · LS2.B Cycles of Matter and Energy Transfer",
  "reflection": "C3 vs C4: ..." }
```

`name` is the tag used on criteria: short, no spaces preferred (DCI, SEP, CCC, or the school's
Power Standard codes such as DCI.2). Every bucket the standard maps to under the profile.

## stimulus

```json
{
  "context": ["Paragraph one.", "Paragraph two."],
  "blocks": [
    { "key": "table1", "kind": "table", "title": "Table 1. ...",
      "columns": ["Measurement", "Open bin", "Sealed bin"],
      "rows": [["Starting mass (kg)", "10.0", "10.0"]],
      "note": "Illustrative data constructed for this assessment.",
      "source": "", "constructed": true, "image": "" },
    { "key": "sourceB", "kind": "text", "title": "Source B. ...",
      "text": ["Paragraph."], "source": "Summary of ...", "constructed": false, "image": "" }
  ],
  "glossary": [{ "term": "aerobic", "def": "with oxygen present" }]
}
```

`image` (optional) is a PNG file name. On paper it is embedded if present in the build folder; on
the Form it is looked up in Drive, with the text or table as fallback.

## parts

`[{ "n": 1, "title": "Part 1 · What happened in the bins", "help": "Use Table 1." }]` Page-break
groups. Most clusters have two.

## criteria

One object per scored judgment; each is also one student-facing item. Parts (a) and (b) are
separate entries.

| Field | Required | Notes |
|---|---|---|
| `id` | yes | C1..Cn, grouped by primary bucket. |
| `item` | yes | Student-facing number: "1", "4a", "4b". |
| `order` | yes | Integer position in student order. Unique. |
| `part` | yes | Matches a `parts[].n`. |
| `type` | yes | "mc" (auto-scored) or "cr" (hand-scored). |
| `buckets` | yes | Array of bucket names. Several = cross-scored. Probes: exactly one. |
| `skill` | yes | A code from `scope`. |
| `feature` | ES criteria | A feature ID with `in_scope: true`. |
| `beyond` | probes | true for a Distinction probe. |
| `extends` | probes | The feature ID the probe pushes past. |
| `stem` | yes | Task verb first. |
| `help` | no | Where to look, a gloss, the sentence cap. |
| `sentence_cap` | cr | Integer. Also set `lines` for paper. |
| `lines` | cr | Answer lines on paper (calibration in `build.md`). |
| `choices` | mc | `[{ "text": "", "key": true, "misconception": "" }]`, four options, one key, every distractor names its misconception. |
| `criterion` | yes | Full criterion text, self-contained. |
| `earns` | yes | What earns the point, in a marker's words. |
| `cue` | yes | Three to six words for the scoring sheet. |
| `compound` | no | true if it joins two constructs; `criterion` must end "Both required." |
| `or_pathway` | no | true if either of two pathways earns it. |

## derived_checks

```json
{ "label": "Open bin mass lost", "expr": "10.0 - 6.2", "expected": "3.8", "tolerance": 0.001 }
```

Every derived number that appears in the stimulus or a key. `expr` uses numbers, + - * / ( ) only.
The checker evaluates it, compares with `expected`, and confirms `expected` appears somewhere in the
stimulus or keys.

## screening, open_decisions

`screening`: `[{ "n": 1, "item": "Phenomenon or problem", "rating": "Yes", "evidence": "", "change": "" }]`
written at Step 5.

`open_decisions`: `[{ "q": "Keep C11 cross-scored to SEP and CCC?", "default": "Yes" }]`. The checker
appends its own warnings to the draft's list automatically; this array holds the human-written ones.
