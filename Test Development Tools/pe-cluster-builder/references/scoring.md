# Scoring

How criteria roll up to reported levels, what the scoring sheet must carry, and what the Form
route's workbook adds for analysis.

---

## 1. Binary criteria, levels at the bucket

Every criterion scores **0 or 1.** The 1-4 level is decided per bucket from the count of criteria
met, never by a holistic judgment on a single criterion. Binary scoring is faster, more reliable
between markers, and less prone to halo effects; the nuance lives in the count.

Default scale (the profile may relabel it):

| Level | Label |
|---|---|
| NYE | Not Yet Evident |
| 1 | Emerging |
| 2 | Developing |
| 3 | Meeting |
| 4 | Meeting with Distinction |

---

## 2. Gates, derived per bucket

The **raw score counts evidence-statement criteria only.** Probes sit outside the denominator: a
student who meets every feature and skips the probe lands at Meeting.

Gates come from each bucket's evidence-statement criterion count *n*, using the profile's
proportions (defaults: Developing 45 percent, Meeting 70 percent, top gate 85 percent), each
rounded **up**. Emerging is any raw score of 1 or more; NYE is 0 only.

| n | Emerging 1 | Developing 2 | Meeting 3 |
|---|---|---|---|
| 3 | 1+ | 2+ | 3+ |
| 4 | 1+ | 2+ | 3+ |
| 5 | 1+ | 3+ | 4+ |
| 6 | 1+ | 3+ | 5+ |
| 7 | 1+ | 4+ | 5+ |
| 8 | 1+ | 4+ | 6+ |
| 9 | 1+ | 5+ | 7+ |
| 10 | 1+ | 5+ | 7+ |

`check_cluster.py` computes the table from the actual counts. Never copy a gate table from another
build.

---

## 3. Distinction modes

Set once in the profile; may be overridden per build with a reason.

| Mode | Level 4 is awarded when | Sheet shows |
|---|---|---|
| `probe` (default) | Meeting gate met **and** the bucket's probe earned | A probe column per bucket |
| `probe` + `mwd_requires_top_gate` | Top gate (85 percent) met **and** probe earned | Same, with the stricter threshold printed |
| `perfect` | Full raw score in that bucket, no probe | "MwD rule: full raw score awards 4, deliberate for this standard" in the rollup and gate rows |
| `none` | Never on this instrument | Levels capped at 3, stated on the sheet |

**Distinction modifies, it never shortcuts.** A student who earns a probe but misses the Meeting
gate takes their gate level. Print this on the sheet; it is the question a parent asks.

`perfect` is legitimate only when fully satisfying every criterion in the bucket already
demonstrates integration across contexts. A single hard item does not qualify.

---

## 4. The scoring sheet

One working document, marked while looking at a student paper. **Target one page, two at most.**

Contents, in order:

1. **Header:** standard code, cluster title, name/block line.
2. **Criteria table in bucket order:** ID, short cue phrase (three to six words capturing the
   judgment), bucket tag(s), scope code, item, auto/hand, score box. Full what-earns-the-point text
   lives in the Answer Key, not here.
3. **Probe rows** (mode `probe`): one per bucket, cue phrase, item, score box.
4. **Rollup and gates, merged:** one row per bucket: ES count, Emerging, Developing, Meeting,
   level-4 rule, raw score box, reported level box.
5. **Three standing blocks:**
   - **Follow-through.** A student who reaches a wrong value early and reasons correctly from it
     loses the criterion where the error occurred, not everything downstream.
   - **Accumulation** (only if the profile says levels are per reporting period). Bucket levels
     accumulate; a level earned here does not overwrite a higher level earned for the same bucket
     elsewhere.
   - **Grading reflection.** One line per bucket naming the pair of criteria that should separate
     students, and what a class-wide result on that pair says about instruction. Written last,
     from the finished criteria, never from the bucket name.

Compact styling is allowed on this sheet only (about 8 to 9 pt table text, tight cell margins). If
a large criteria set still needs a third page, accept it; do not drop a required block.

**Tag with bucket names only,** never framework practice numbers. "SEP.3" as a school's
Communicating standard and "SEP 3" as the NGSS practice Planning and Carrying Out Investigations
have been confused on a real sheet.

---

## 5. The answer key

Every item in student order: the correct letter and its full text, or the full what-earns-the-point
for a constructed response or probe, with criterion ID and bucket under each. Carries the
follow-through statement. This is the document handed to a substitute marker or used in a parent
conversation. It is the only place full scoring language appears on the paper route.

---

## 6. Form route workbook

Built by the generated Apps Script. Sheets:

| Sheet | Read when | Holds |
|---|---|---|
| Form Responses 1 | never directly | raw Form output |
| Scoring | during grading | per-criterion 0/1 (auto for selected response, hand-entered and yellow for constructed), bucket raw, probe, level |
| Key & Gates | during grading | criterion key, what earns each point, derived gates, probes, standing blocks |
| Features Analysis | after grading | per-feature class percentages, probe results, level distribution, by-teacher split, chart |
| Student x Feature | after grading | one row per student, 1/0 per feature: the reteaching list |

Features Analysis runs on live formulas, so it fills in as hand-scoring proceeds. Worth knowing:

- **Probes are reported in their own block,** never mixed with feature coverage. A probe at 20
  percent is a hard question working as designed.
- **The split flag.** Where one feature is assessed by two criteria whose class percentages differ
  by more than 30 points, the row flags. That usually means one of the two items is broken, not
  that students half-learned the feature.
- **By-teacher split.** A gap on one feature is a teaching signal; a gap on every feature is
  usually timing or delivery.
- **Colour thresholds** on combined feature percentage: red below 50, amber 50 to 70.

Student x Feature marks a feature met only when every criterion assessing it was earned, and stays
blank until all are scored, so a half-graded set does not read as failure.

Quiz grade release cannot be set from Apps Script. Forms quizzes default to manual release; say so
rather than implying the script controls it.
