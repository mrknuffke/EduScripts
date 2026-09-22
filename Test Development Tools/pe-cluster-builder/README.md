# PE Cluster Builder

A discipline-neutral assessment-authoring tool that turns a single NGSS Performance Expectation (or any standards-based learning target) into a complete, defensible item cluster in roughly one LLM session — then generates every deliverable from a single JSON source file.

---

## 🎯 What Is This Tool?

The **PE Cluster Builder** produces a stand-alone assessment instrument on exactly **one standard**, designed to be given as soon as that standard is taught. Everything that varies by course — reporting buckets, scope list, item floors, fonts, paper size, local context — lives in a **course profile**, so the same tool serves chemistry, biology, physics, earth science, or a non-NGSS course without any edits.

Each build delivers:

| Output | Description |
| :--- | :--- |
| **Teacher / PLC Review Draft** (`.docx`) | Stimulus, all items, answer key, criteria text, and open decisions — for review before anything goes to students |
| **Student Paper** (`.docx`) | Clean stimulus + items, no answers |
| **Answer Key** (`.docx`) | Full model answers and rubric |
| **Scoring Sheet** (`.docx`) | 0/1-per-criterion scoring grid rolling up to a 1–4 SBG level per reporting bucket |
| **Google Form + Auto-scoring Workbook** (`.gs`) | Ready-to-paste Apps Script; creates the Form and a linked scoring spreadsheet in one click |

---

## 🛠️ Prerequisites

**Python 3.9+** (for the checker and Form generator):
```bash
# No additional packages required — uses only the standard library
python3 --version
```

**Node.js 18+** (for the .docx builder):
```bash
cd "Test Development Tools/pe-cluster-builder"
npm install        # installs the docx package (~30 seconds, one time)
```

---

## 🚀 Build Workflow

All commands assume `content.json` is in the current directory and outputs go to `out/`.

```bash
# 1. Validate your spec after every edit
npm run check
# or: python3 scripts/check_cluster.py content.json

# 2. Generate a teacher / PLC review draft (Step 6 in the skill workflow)
npm run draft
# or: node scripts/build_docs.js content.json out/ --draft

# 3a. Generate the final student paper set (Step 7 — paper route)
npm run final
# or: node scripts/build_docs.js content.json out/ --final

# 3b. Generate the Google Form Apps Script (Step 7 — Form route)
npm run form
# or: python3 scripts/make_form_script.py content.json out/
```

> [!IMPORTANT]
> `build_docs.js` calls `check_cluster.py` automatically and **refuses to build `--final` while the checker reports errors**. Fix all errors before delivering to students.

---

## 📁 File Reference

```
pe-cluster-builder/
├── SKILL.md                          ← Main LLM instruction file
├── references/
│   ├── build.md                      ← .docx geometry, components, validation
│   ├── content-schema.md             ← content.json field reference
│   ├── design-rules.md               ← Item architecture, coverage, language rules
│   ├── domain-adaptation.md          ← Discipline-specific traps; non-NGSS frameworks
│   ├── scoring.md                    ← Gate arithmetic, Distinction rules
│   └── screening-tool.md             ← 11-item 3D quality screen
├── assets/
│   ├── course-profile-template.md    ← Fill this in once per course; run the interview
│   ├── example-content.json          ← Complete worked cluster (HS-LS2-3) for reference
│   └── form-builder-template.gs      ← Apps Script template (never edit by hand)
└── scripts/
    ├── check_cluster.py              ← Structural, item, and text-hygiene checker
    ├── build_docs.js                 ← Generates all .docx outputs
    └── make_form_script.py           ← Generates the .gs from content.json
```

---

## 🤖 Using with an LLM

The PE Cluster Builder is designed to run **inside a long-context LLM session**. The LLM writes and refines `content.json` interactively with the teacher; the scripts then generate the deliverables locally.

### Step 1 — Download the files

Clone the full repo, or download just this tool:

```bash
# Option A: clone the whole repo
git clone https://github.com/mrknuffke/EduScripts.git
cd EduScripts/Test\ Development\ Tools/pe-cluster-builder

# Option B: download a ZIP from GitHub
# → Code → Download ZIP → unzip → navigate to Test Development Tools/pe-cluster-builder/
```

### Step 2 — Choose which files to upload

Upload these **8 files** as project knowledge. They are the complete specification the LLM needs:

| File | Why |
| :--- | :--- |
| `SKILL.md` | Core workflow, non-negotiables, all 8 steps |
| `references/design-rules.md` | Item architecture and language rules |
| `references/scoring.md` | Gate arithmetic and Distinction rules |
| `references/screening-tool.md` | 11-item quality screen |
| `references/content-schema.md` | Every field in `content.json` |
| `references/build.md` | Output format and build commands |
| `references/domain-adaptation.md` | Discipline-specific traps; non-NGSS mappings |
| `assets/course-profile-template.md` | The setup interview the LLM runs on first use |

**Optional enrichment** (upload if you want the LLM to reference a completed example):
- `assets/example-content.json`

### Step 3 — Set up your LLM project

#### Claude Projects (recommended)
1. Go to [claude.ai](https://claude.ai) → **Projects** → **New Project**.
2. Name it, e.g. *PE Cluster Builder — AP Bio*.
3. Click **Add content** → upload the 8 files above.
4. Optionally paste your filled-in course profile into the project instructions box so it is always available.
5. Start a new conversation inside the project with the prompt below.

#### ChatGPT Projects
1. Go to [chatgpt.com](https://chatgpt.com) → **Projects** → **New Project**.
2. Open the project → **Files** tab → upload the 8 files above.
3. Use the same starting prompt below.

#### Antigravity / Gemini IDE (auto-discovery)
Drop the entire `pe-cluster-builder/` folder into `.agents/skills/` at the root of any repo:
```
YourRepo/
└── .agents/
    └── skills/
        └── pe-cluster-builder/    ← this whole folder
            └── SKILL.md           ← Antigravity reads this automatically
```
Antigravity discovers and loads the skill on startup — no manual uploads needed.

### Step 4 — Starting prompt

Paste this to kick off a session (edit the bracketed parts):

```
Use the PE Cluster Builder skill.

Course: [e.g. AP Biology, Grade 11]
Standard: [e.g. HS-LS2-3]
Delivery mode: [paper | Google Form]

[Paste your course profile here, OR say "run the course profile interview" if this is your first time.]
```

The LLM will walk through the full 8-step workflow — one question at a time — and produce a `content.json` that you save locally. Then run the build scripts to generate your deliverables.

> [!NOTE]
> `content.json` is your live work file. It is listed in `.gitignore` and stays on your machine only. Finished clusters can be archived as named files (e.g. `clusters/hs-ls2-3-compost-bins.json`) if you want to version them.

---

## 📐 Course Profile

The course profile is a one-time setup document that captures everything about your course that varies across builds:

- Standards framework (NGSS, AP, IB, state)
- Reporting taxonomy and bucket map
- Scope list with codes
- Item floor per bucket
- Distinction mode (`probe`, `perfect`, or `none`)
- SBG level labels
- Paper size, fonts, accent colour
- Locale for context anchoring

Once filled in and saved as `references/course-profile.md`, the LLM skips the setup interview and loads it automatically at the start of every session.

The template is in [`assets/course-profile-template.md`](assets/course-profile-template.md).

---

## 🔒 Data & Privacy

All files generated by this tool (`.docx`, `.gs`) are output to `out/` and are excluded from version control. No student data is ever written into `content.json` or any committed file — the spec contains only assessment structure, items, and criteria text.

---

## 📖 Further Reading

- [`SKILL.md`](SKILL.md) — complete LLM workflow (8 steps)
- [`references/design-rules.md`](references/design-rules.md) — item design rationale
- [`assets/example-content.json`](assets/example-content.json) — fully worked HS-LS2-3 cluster
