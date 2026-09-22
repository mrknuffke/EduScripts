# Test Development Tools

A collection of AI-assisted assessment and item-authoring tools for science teachers. Each tool follows a common pattern: an LLM session builds a structured specification file, local scripts then generate the deliverables.

---

## 🛠️ Tool Directory

| Tool | Framework | Purpose |
| :--- | :--- | :--- |
| [**`pe-cluster-builder`**](pe-cluster-builder/README.md) | NGSS / any framework | Builds a single-standard item cluster → teacher draft, student paper, answer key, scoring sheet, and optional Google Form with auto-scoring Apps Script |
| `ngss-exam-builder` | NGSS | *(coming)* Full summative exam spanning multiple Performance Expectations |
| `ap-bio-frq-generator` | AP Biology | *(coming)* Free-response question generator aligned to AP Biology CED |
| `ap-bio-exam-maker` | AP Biology | *(coming)* Full AP Biology exam assembler (MCQ + FRQ sets) |

---

## 🤖 Common LLM Setup Pattern

All tools in this directory follow the same setup:

1. **Download** — clone the repo or download a ZIP from GitHub.
2. **Upload to an LLM project** — each tool's `README.md` lists the exact files to upload.
3. **Start a session** — paste the tool's starting prompt; the LLM interviews you and builds the spec.
4. **Run the scripts locally** — generate `.docx`, `.gs`, or other deliverables from the spec.

See each tool's `README.md` for platform-specific instructions (Claude Projects, ChatGPT Projects, Antigravity).

---

## ➕ Adding a New Tool

Drop a new subdirectory here following this structure:

```
new-tool-name/
├── SKILL.md                 ← Main LLM instruction file (required)
├── README.md                ← Tool docs with LLM setup section
├── LICENSE
├── references/              ← Reference docs uploaded as LLM project knowledge
├── assets/                  ← Templates and examples
└── scripts/                 ← Local build/validation scripts
```

Then add a row to the table above and a row to the root [`README.md`](../README.md) index.
