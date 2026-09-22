# 3D Screening Tool

Eleven items describing what a three-dimensional assessment looks like. Used twice: as **design
inputs** before drafting (SKILL.md Step 2) and as **one report** on the finished spec (Step 5). It
is a report, not a gate; no rating stops a build.

For non-NGSS frameworks, items 2, 4, 5 and 6 need translating to that framework's dimensions; see
`domain-adaptation.md`.

## Vocabulary

- **PE:** the whole Performance Expectation.
- **DCI:** the content. **SEP:** the practice. **CCC:** the crosscutting thinking tool.
- **Phenomenon:** an observable event in the natural world. **Problem:** a human need or want
  (engineering).
- **Stimuli:** the information (data, text, images) the prompts require. **Prompts:** the questions.

## The eleven items

Rated **No / Partially / Yes.**

1. Contains a phenomenon (science) or a problem (engineering).
2. The prompts match the SEP and engage students in sense-making.
3. The stimuli carry multiple and sufficient information to use the SEP (e.g. more than one data
   set).
4. The prompts elicit observable understanding of the DCI.
5. The prompts explicitly name the CCC.
6. The prompts use language from grade-appropriate progressions (SEP, DCI, CCC).
7. The response space fits the observable features the student must produce.
8. Everything is scientifically accurate and properly attributed (no invented data presented as
   real; sources given).
9. The prompts point toward explaining a phenomenon or designing a solution.
10. The phenomenon or problem is authentic, interesting, and requires figuring something out.
11. The phenomenon or problem is novel to the unit, so transfer is shown.

## As design inputs

| Items | Decision to settle before drafting |
|---|---|
| 1, 9, 10 | What is the phenomenon or problem, is it locally anchored and named, and does the cluster point toward explaining or designing? |
| 11 | Which context is novel to the unit? Confirm it was not used in instruction. |
| 2, 4, 5 | Which items carry the SEP, DCI and CCC, and does at least one stem name the CCC in student language? |
| 3 | What stimuli does the SEP need, and does the time budget afford them? |
| 6 | Which evidence-statement bullets are being lifted into stems? |
| 7 | How many features must each constructed response produce, and how much answer space does that take? |
| 8 | Where does the data come from, and how is it attributed or labelled illustrative? |

An item raised here is a sentence. The same item raised in the report is a rewrite.

## As a report

| # | Item (short) | Rating | Evidence | Change offered |
|---|---|---|---|---|

- **Evidence is specific:** "Item 5 names 'energy drives the cycling of matter'," not "the CCC is
  present."
- **A change is offered for every No and Partially:** replacement wording, an added data set, a
  revised context, a different line count. Blank on a Yes.
- Do not rate Yes to avoid writing a change. Where the constraints make an item unreachable (item 3
  in a 40-minute cluster often affords one data set plus one text source), rate it honestly and
  say what would be needed.

Ask which changes to make. Apply accepted ones in `content.json`, re-run the checker, report only
the items the change touched. The report lives in the conversation and the draft; never in a
student-facing or final teacher-facing document.

## Item notes

- **5:** "Explicitly" means the stem names the idea in student language (cause and effect, scale,
  stability and change, energy and matter, patterns, systems, structure and function). Exercising
  it without naming it is Partially.
- **6:** The evidence statement's observable-feature bullets are the progression language. Lift
  phrasing from them rather than paraphrasing.
- **7:** Count the features a constructed response must produce against its sentence cap and answer
  lines. A four-feature explanation given two lines is a No.
- **8:** Yes means the numbers were recomputed by script, not read.
- **10:** Anchored to the profile's locale and named. A generic textbook scenario is a No.
- **11:** The phenomenon is new; the skills are not. Both constraints hold at once.
