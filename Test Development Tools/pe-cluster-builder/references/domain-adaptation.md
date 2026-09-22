# Domain Adaptation

The cluster architecture (one standard, features to criteria to buckets, anchored items, binary
criteria, derived gates, probes) is discipline-neutral. What changes by domain is where the
features come from, what the buckets are, and which traps cost the most rework. Read the section
for the course at Step 1.

---

## 1. What changes and what does not

| Stays fixed in every domain | Changes by domain or course (lives in the course profile) |
|---|---|
| One standard per cluster | The framework and its codes |
| Every criterion anchored to a feature or a declared probe | Where features come from (evidence statement or teacher-built list) |
| 0/1 criteria, levels by derived gate | Bucket taxonomy |
| Draft before build, one source of truth | Scope list and its source |
| Key distribution, length matching, named misconceptions | Item floor, Distinction mode, levels scope |
| Stimulus contains everything; data recomputed by script | Locale for context, paper size, fonts, spelling |
| Language accessibility rules | Sentence cap (lower for younger or multilingual cohorts) |

---

## 2. Setting up a new course: the recipe

1. Copy `assets/course-profile-template.md`, fill it in (or let Claude run the interview), save it
   as `references/course-profile.md` inside the skill folder, re-zip, re-upload. One profile per
   course. A teacher with two courses keeps two profile files and names the one to load.
2. **Bucket map.** Generic NGSS (DCI / SEP / CCC, one bucket each, labelled with the PE's own
   dimension) is the default. A school Power Standard system replaces it: paste the table of
   Power Standard to PE mappings into the profile.
3. **Scope list.** Whatever the department uses to say what was actually taught: a skill registry
   with codes, learning targets, "I can" statements, a unit review sheet. It needs stable codes.
   If none exists, number the unit's learning targets (U3.1, U3.2 ...) and use those.
4. Run one cluster end to end on a PE the teacher knows well, and compare the draft to what they
   would have written. Adjust the profile, not the skill.

---

## 3. By discipline

### Chemistry (HS-PS1, parts of PS2 and PS3)

- **Numbers that collide.** A 1:1 mole ratio makes "mass consumed" and "mass remaining" equal and
  a criterion unmarkable. After deriving, look for any two different quantities sharing a value.
- **Formulas in Unicode everywhere**, student prose included: H₂O, SO₄²⁻, 6.02 × 10²³, ΔH.
- **Classification criteria accept every classification the unit teaches** (acid-base and double
  replacement, if both are taught for the same reaction).
- **Held-back mechanism terms** are common: shielding, hybridisation, orbital notation. Ask.
- Particle diagrams as stimulus: supply as PNG with a text description fallback.

### Physics (HS-PS2, PS3, PS4)

- **Quantitative distractors come from specific algebra errors,** each named: sign error, inverted
  ratio, missing square, unit slip (g vs kg, km/h vs m/s), using mass where weight is needed.
- **Every derived value goes in `derived_checks`.** Physics clusters carry more of them than any
  other domain.
- **Units and significant figures:** state in the stem whether they are assessed. If not, the key
  accepts any reasonable rounding and the criterion says so.
- **Assessment boundaries are tight and specific** (for example, HS-PS2-1 limits to one-dimensional
  motion of macroscopic objects at non-relativistic speeds). Read the boundary line before every
  stem.
- **Graphs:** supply the graph as PNG with the data table as the fallback, so the Form route still
  works if the image is missing. A graph-reading item must be answerable from the table too, or it
  becomes a "read the picture" item on the fallback.
- Calculator policy goes in `meta.resources`.

### Life science (HS-LS1 to LS4)

- **Teleological stems and keys are defects.** Avoid "in order to", "needs to", "so that it can"
  for organisms or populations. They are also the richest source of distractors: the teleological
  explanation is often the best-performing wrong option.
- **Common distractor sources:** plants get their mass from soil; energy is recycled in food webs;
  individuals evolve; mutations happen because the organism needs them; decomposers destroy matter.
- **Human genetics and health contexts:** never use student or family data; avoid contexts that
  stigmatise a condition or map traits onto race. Non-human systems usually carry the same
  construct with less risk.
- **Statistics PEs** (e.g. HS-LS3-3, HS-LS4-3) need real distributions or constructed data whose
  summary statistics are recomputed by script, including any percentages in keys.
- **Modelling PEs that need a drawn model** (e.g. HS-LS1-2, HS-LS1-5) fit the paper route. On the
  Form route, Apps Script cannot add a file-upload question, so switch to evaluating or completing a
  given model, and flag that change as an open decision.
- **Local organisms and ecosystems** for context anchoring (the profile's locale).

### Earth and space science (HS-ESS1 to ESS3)

- **Real data is strongly preferred** and is what screening item 8 is looking for. Public sources
  such as NOAA, NASA, USGS and national meteorological or geological services publish downloadable
  sets. Cite source, data set name and access date in the block's `source` field.
- **Scale is usually the CCC.** Stems should make students reason across time or space scales
  explicitly, which also satisfies screening item 5.
- **ESS3 human-impact PEs:** keep stems on evidence and mechanism. Items that ask students to
  endorse a policy measure opinion, not the standard.
- Map and cross-section stimuli need a text fallback that preserves the spatial relationships the
  items depend on (e.g. a table of depths and ages by site).

### Engineering (HS-ETS1, and PEs whose SEP is designing solutions)

- **A problem, not a phenomenon.** Screening items 1 and 9 read "problem" and "designing".
- Features usually concern criteria and constraints, trade-offs, and evaluating a solution against
  them. Hand the student someone else's design with its test results.
- **Probes that work:** evaluate the design against a constraint the stimulus did not state, or
  name an unintended consequence and a test for it.
- Many science PEs carry an engineering practice (for example HS-PS2-3, HS-LS2-7, HS-ESS3-2).
  Treat those as engineering clusters for context design even though the DCI bucket is science.

### Environmental science and integrated courses

Same rules. If the course bundles PEs from several domains into one Power Standard, the cluster
still assesses one PE; the profile's map tells you which buckets that PE feeds.

---

## 4. By grade band

- **Middle school:** NGSS publishes MS evidence statements. Lower the sentence cap to about 18
  words, keep constructed responses to two or three sentences, and expect fewer features, so a
  30 to 35 minute cluster is often right. Glossary every term beyond the unit's core vocabulary.
- **Elementary:** NGSS evidence statements exist for K-5, but a 40-minute independent written
  cluster is rarely appropriate. Use the feature-to-criterion logic to build a short observed task
  or interview protocol instead, and skip the Form route.

---

## 5. Frameworks other than NGSS

The one thing to replace is the **evidence statement**: something that decomposes the standard
into observable features a marker can point at.

| Framework | What stands in for the evidence statement | Typical buckets |
|---|---|---|
| NGSS-derived state standards | Usually the NGSS evidence statements themselves; check the state has not rewritten the PE. Some states publish item specifications with misconceptions and boundaries worth using. | As NGSS, or the state's reporting categories |
| AP sciences | The Course and Exam Description: the Learning Objective and its Essential Knowledge statements for content, the Science Practice skill for the practice | The course's Science Practices, plus a content bucket, or the school's own lumping |
| IB DP sciences | The syllabus understandings for the topic plus the assessment objectives | The assessment objectives, or the school's lumping |
| GCSE / A level | Specification content statements plus the assessment objectives (AO1 to AO3) | The assessment objectives |
| A school's own standards | The standard's success criteria or learning targets | Whatever the gradebook reports |

When no published decomposition exists, **build a teacher-authored feature list**:

1. Split the standard into clauses.
2. Rewrite each clause as an observable student performance with a verb a marker can point at
   (identifies, calculates, predicts, justifies, revises), not "understands" or "knows".
3. Add the practice dimension explicitly: what the student does with the content.
4. Mark every feature `ref: "teacher"` in `content.json` and put "Confirm the teacher-authored
   feature list" as open decision 1 in the draft.
5. Everything downstream is unchanged.

Screening items 2, 4, 5 and 6 are NGSS-specific. For AP or IB, read item 2 as "matches the
science practice", item 4 as "elicits the content", item 5 as "names the practice or skill in
student language" (or mark it not applicable), and item 6 as "uses the framework's own language".
Record the translation in the report.

---

## 6. Outside science

The engine transfers anywhere a standard can be broken into observable features and a stimulus
can carry a task. It is untested outside science, so treat the first build as a pilot.

- **Mathematics:** content standard clauses as features; the mathematical practices (or the
  school's process strands) as the practice bucket. Distractors come from named procedural and
  conceptual errors. `derived_checks` matters more than anywhere else.
- **History and social studies:** a source set as the stimulus; features from the standard and the
  disciplinary skill (sourcing, corroboration, contextualisation, claim and evidence). Screening:
  replace item 5 with "names the disciplinary concept", and read items 1 and 9 as "an inquiry
  question the sources can answer". Attribution rules for real sources are stricter, not looser.
- **Language and literature:** weakest fit. Most standards are holistic and a holistic rubric
  serves them better. Use a cluster only for tightly defined skills (identifying an author's claim
  and evidence, for example).

What does not transfer outside science: the DCI / SEP / CCC bucket names, the CCC cross-scoring
logic (replace with whatever second dimension the framework has, or drop cross-scoring), and the
eleven-item screen as written.
