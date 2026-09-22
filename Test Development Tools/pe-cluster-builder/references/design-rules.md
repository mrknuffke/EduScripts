# Design Rules

The rules a cluster is drafted against. Read before writing any item. Every rule here is either
enforced by `scripts/check_cluster.py` or listed in the draft's open decisions when it is bent.

---

## 1. Coverage is driven by the evidence statement

The evidence statement sets the criterion count. The bucket set does not. Work features to
criteria to buckets, in that order.

1. **Enumerate every observable feature** in the evidence statement, every numbered and lettered
   row. Give each an ID (F1, F2 ...) and keep its statement reference (e.g. "2a-iii").
2. **Filter by scope.** A feature with no matching code on the course's scope list is out. This
   filter runs before design, not after. Record the reason for each exclusion.
3. **Every surviving feature is assessed at least once.** No feature is dropped for length, for
   grading load, or because it is hard to write an item for.
4. **A feature may be assessed twice** where it carries enough weight. The two exposures must be
   distinguishable: typically one applies the feature in the situation given, the other carries it
   to a case that was not given. Two items asking the same thing in different words are one
   criterion.
5. **Tag each criterion with the bucket or buckets it evidences.** Bucket counts are the output.
6. **Write one Distinction probe per bucket** (section 4), if the profile's Distinction mode is
   `probe`.

### What the counts look like

Uneven, and that is correct. A content-heavy statement produces a fat DCI bucket and a thin CCC
one. Do not pad a thin bucket with an invented criterion and do not trim a fat one for symmetry.

Two things bound it:

- **Every bucket the PE maps to under the profile is scored.** A mapped bucket with no criteria is
  a design failure. Fix it by cross-scoring (section 2) or by re-reading the statement for the
  feature that carries that dimension.
- **The item floor per bucket** from the profile. Below it the bucket's gate stops distinguishing
  levels. Raise any shortfall with the teacher; resolve by cross-scoring, by extending a prompt, or
  by an explicit provisional-level note on the sheet. Never by padding.

**Split PEs.** Where a course teaches one PE across two units, assess it once, late, as a single
cluster covering all its features, with the earlier material carried as prior knowledge. Do not
build two half-clusters.

---

## 2. Criterion architecture

**One item, one criterion** is the default. It keeps the Form route auto-scorable and the paper
route fast. Parts (a) and (b) of a question are separate items.

### Cross-scoring: one criterion, several buckets

A well-built item can genuinely evidence two dimensions at once. Such a criterion carries several
bucket tags, is scored **once**, and counts in **every** bucket it is tagged to. Licensed only
where all hold:

- The item genuinely demands both dimensions: a student could get one right and the other wrong.
- The criterion text names what is required in each dimension, so a second marker scoring it for
  DCI and for CCC reaches the same 0 or 1.
- The tag list appears on the criterion in the draft and in every output.

Cross-scoring correlates bucket levels, so bound it:

- **At least half of every bucket's criteria are exclusive to it, and never fewer than two.**
- **Any criterion tagged to three or more buckets** is listed in open decisions and approved.

### Compound criteria (AND), bounded

A criterion may join **at most two** constructs, and only when one feature genuinely joins them:

- Both constructs are stated explicitly in the criterion text.
- The criterion text ends with "Both required."
- The what-earns-the-point entry names both, separately.
- If a student could plausibly produce one and not the other, they are two criteria.

Three constructs is never licensed. Every compound criterion goes in open decisions.

### OR-pathway criteria, bounded

The mirror image: either of two constructs is independently sufficient. Use it where the unit
taught two independent, non-overlapping mechanisms for the same phenomenon, so a stem written
toward one would penalise a student reasoning correctly from the other.

- Both pathways are named in what-earns-the-point, each stated as fully sufficient.
- The stem cues neither pathway (no vocabulary or diagram matching only one).
- The pathways are genuinely independent, not two phrasings of one chain.
- The grading reflection notes which pathway students used; a lopsided split is a teaching signal.

Every OR-pathway criterion goes in open decisions.

### Criteria within a bucket must be able to diverge

Build a deliberate spread in every bucket: one most of the class should get, some mid, one that
separates. Criteria that move together produce bucket scores at the extremes and nothing between.
Name the intended split in the grading reflection.

### Numbering

Criterion IDs run C1 to Cn grouped by the bucket each is primarily read for. Item IDs follow
student order (1, 2a, 2b, 3 ...). The two orders diverge, and that is correct. Every criterion
records its item.

---

## 3. Criterion writing rules

- **One construct per criterion** (compound carve-out above).
- **State the full-credit condition, not the topic.** "Identifies the sealed bin as the only one
  where oxygen ran out, and links that to methane appearing only there," not "anaerobic
  conditions."
- **Name the specific science wherever correctness is checkable,** so a second marker reaches the
  same judgment without asking anyone.
- **Name the item** the criterion is read from.
- **Extend the prompt, never invent the criterion.** A criterion a student could not earn from the
  question as written is a scoring trap. If a bucket needs another criterion, add the task.
- **Every criterion names its scope code.**
- **Classification criteria accept every classification the unit teaches.**
- **No criterion or key relies on a term the unit has not named,** even where the underlying
  science is in scope (the held-back terms from Step 1).

**Follow-through error.** Where criteria chain (a derived value feeds a later judgment), score each
on its own logic. A student who slips once and reasons correctly from their own wrong value loses
one criterion, not every one downstream. State this in the answer key.

---

## 4. Meeting with Distinction probes (Distinction mode `probe`)

Distinction is **directly asked and directly scored.** Each bucket carries exactly one
beyond-statement probe, written as **part (b) of an item that already exists in that bucket.**
Part (a) asks the evidence-statement thing; part (b) pushes past it. Constructed response,
two-sentence cap.

- **One probe per bucket, never cross-scored.** Distinction in DCI, SEP and CCC are different
  accomplishments. The checker rejects a probe with more than one tag.
- **Scored 0 or 1 like any criterion,** against a stated what-earns-the-point as concrete as any
  other.
- **No second pathway.** A student who volunteers a beyond-statement move elsewhere on the paper
  does not earn the probe. Scanning a whole paper for unprompted moves is unreliable between
  markers.
- **Considerations phrased around "unprompted" behaviour are rewritten.** The construct is "can go
  beyond when asked," not "went beyond on their own." That shift is deliberate; the unprompted
  signal belongs in holistic projects.

What makes a good probe, by bucket:

| Bucket | The push past the statement |
|---|---|
| DCI | Carry the mechanism to a condition the stimulus did not show, or connect it to a second content area the unit taught |
| SEP | A more sophisticated practice move: name what evidence would strengthen or weaken the argument, identify a limitation of the data or model, propose a refinement that acts on a flaw the student named |
| CCC | Apply the crosscutting lens to a second system or a different scale |

Other Distinction modes (set in the profile): `perfect` awards 4 for a full raw score with no probe;
`none` caps the instrument at 3. See `references/scoring.md`.

---

## 5. Context rules

The context is **adjacent to the unit storyline, not a repeat of it.** Same phenomenon family,
different system. The student should recognise the kind of problem, not the problem.

- Reusing the unit's driving question, its data sets, worked examples or lab artefacts is a defect.
  So is a context so remote that the first ten minutes go to decoding it.
- **Hand the student someone else's investigation:** another group's data, a technician's log, a
  field station's record, an incident report. The student did not run it, which is what makes
  limitation-and-refinement items answerable.
- **The stimulus contains everything needed.** No criterion depends on a value, diagram or
  convention that lives only in the unit's slides.
- **Local anchoring** per the profile's locale: named places, organisations, organisms or systems
  students could plausibly meet. A generic textbook scenario fails screening item 10.
- **Constructed data** is usually right, so intended splits are reachable. It must be internally
  consistent (recompute every mean, difference and percentage by script), labelled as illustrative
  in the stimulus, and must not contradict real-world magnitudes. **Real data** is quoted with its
  source and date. Never present constructed data as real.
- **No two quantities in the stimulus share a value** in a way that lets a student copy across or
  makes a criterion unmarkable. The checker flags repeated numbers for review.

Screening item 11 (novel context) is satisfied by adjacency. The context must be new; the skills
must not be.

---

## 6. Item architecture and time

Forty minutes (or the profile's duration) covers reading the stimulus and answering every item.
Budget about five minutes for the stimulus and roughly three minutes per item: a ceiling of about
**thirteen scored judgments**, i.e. nine or ten evidence-statement criteria plus one probe per
bucket. A probe as part (b) costs a minute or two, not a full slot.

The criterion count comes from the evidence statement, so it comes first and the clock is checked
against it. Where coverage exceeds the time, say so at the draft stage and let the teacher choose:
a longer sitting, two sittings, or some features assessed once rather than twice. Never silently
compress items to fit.

**Selected response** for anything with a single defensible answer: rankings, identifications,
predictions with one correct direction, reading a bound from data. A good four-option item on a
prediction-plus-reason carries as much reasoning as a written one and costs nothing to score.

**Constructed response** where the reasoning is the evidence: choosing between conflicting
patterns, revising a claim, naming a mechanism, proposing a refinement. Cap at three or four
sentences and say so in the item.

Aim for roughly two thirds selected response as a consequence, not a quota. Cross-scored criteria
are usually constructed response.

### Selected-response rules (all checked by script)

- **Distractors carry misconceptions.** Every wrong option is a nameable error students in this
  course actually make: the plausible wrong mechanism, the reversed direction, the right answer for
  the wrong reason. Record the misconception on each distractor.
- **Assign key positions before writing options.** No position is the key on more than half the
  items; across four or more items, keys occupy at least three positions; never all in one.
  Generated item sets drift to option A unless planned. Forms built by Apps Script cannot shuffle
  options, so the spec's positions are what students see.
- **Distractor length matches the key.** All options within a few words of each other. A key that
  is visibly the longest and most qualified is answerable without the science. Also check the key
  is not systematically the longest across the set.
- Four options by default. "All of the above" and "none of the above" are not used.

---

## 7. Language accessibility

Scored on science, not reading. Every item ships at the lowest reading load that keeps the
construct.

- **Define every course term at first use,** inline, in the stimulus or the item's help text. If
  the term is the thing assessed, the gloss states its meaning without giving the answer.
- **One clause per instruction.** Split any sentence with two demands.
- **Sentence cap** of roughly 25 words in stems and options (the profile may lower it; 18 is
  sensible for middle school or multilingual cohorts).
- **No nested conditionals** in a stem.
- **Task verb first** in constructed-response stems: "Name one ..." not "Considering everything
  above, what ...".
- **Numbers in the stimulus, not the stem.**
- **A glossary block** on the stimulus page collecting every defined term.
- **Renamed terms:** introduce the course's current term once with a bridge to the older one
  ("principal energy level, sometimes called a shell"), then use the current term alone.

---

## 8. What goes in open decisions

Every one of these is a numbered question with a default in the draft:

- every compound and OR-pathway criterion
- every cross-scored criterion, and any tagged to three or more buckets
- every bucket below the item floor
- every constructed data set, and every real data set needing a citation
- every place the time budget is tight
- anything the evidence statement licenses but the scope list does not
- any held-back term that constrained a key
