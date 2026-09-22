# Course Profile

Fill this in once per course, save it as `references/course-profile.md` inside the skill folder,
re-zip and re-upload the skill. Claude loads it at the start of every cluster build and only asks
for what is missing.

If you would rather not fill it in by hand, start a chat with "Set up a course profile for the
PE cluster builder" and Claude will ask these questions one at a time, then give you the finished
file.

Interview order: Claude asks section by section, one question at a time, in the order below.

---

## 1. Course

- **Course name:**
- **Grade band:** (MS / HS / other)
- **Framework:** (NGSS / NGSS-derived state standards / AP ___ / IB ___ / other: ___)
- **Where the teacher gets evidence statements or equivalent:** (default for NGSS:
  nextgenscience.org evidence statements; for other frameworks see
  `references/domain-adaptation.md` section 5)

## 2. Reporting buckets

- **Taxonomy:** generic NGSS dimensions (DCI / SEP / CCC) **or** a custom Power Standard system
- **If custom, the map.** One row per bucket; list every PE that feeds it.

| Bucket code | Bucket title | PEs that feed it |
|---|---|---|
| | | |

- **Notes on the map** (e.g. a PE that feeds two content buckets, a PE split across units):

## 3. Scope

- **Scope list source:** (skill registry / learning targets / unit review sheet / other)
- **Code format:** (e.g. LS2-3.1, U4.2)
- **Where the list lives:** (paste it below, or say it will be uploaded per build)

```
(paste the scope list here, one code and statement per line, grouped by PE, or leave blank)
```

## 4. Scoring

- **Minimum criteria per bucket (item floor):** (default 3)
- **Levels apply:** per assessment / per reporting period (buckets accumulate)
- **Level labels:** (default NYE, Emerging, Developing, Meeting, Meeting with Distinction)
- **Gate proportions:** Developing ___ % / Meeting ___ % / Top ___ % (defaults 45 / 70 / 85)
- **Distinction mode:** probe (default) / perfect / none
- **Distinction requires the top gate rather than Meeting:** yes / no (default no)

## 5. Format

- **Paper:** A4 / Letter
- **Body font / heading font:** (defaults Garamond / Montserrat Medium)
- **Accent colour (hex):** (default 4A90A4)
- **Spelling:** UK / US
- **Sentence cap in stems and options:** (default 25 words; 18 for MS or multilingual cohorts)
- **Default duration (minutes):** (default 40)
- **Default resources line:** (e.g. "Calculator allowed. No notes. No AI assistants.")

## 6. Context

- **Locale for context anchoring:** (city or region, and any local institutions, organisms or
  systems that make good contexts)
- **Contexts to avoid:** (e.g. contexts already used in unit storylines, sensitive topics)

## 7. Delivery

- **Default delivery:** paper / Google Form plus scoring workbook
- **Teacher roster for the Form dropdown:**
- **Do students sign in to Google (collect emails)?** yes / no

## 8. Standing held-back terms

Terms this course deliberately does not introduce, with the substitute reasoning path. Claude
still asks per build for unit-specific ones.

| Term | Substitute reasoning the course uses |
|---|---|
| | |
