---
name: survey-report-writing
description: "Write or refine 412YZ and IL annual survey report narrative. Use when: writing section-by-section report insights, comparing this year's data tables against last year's report, calling out year-over-year differences, updating 05_report_412YZ.js or equivalent IL script with finalized narrative text."
argument-hint: "Section name or 'all' to step through every section"
---

# Survey Report Writing

Skill for producing data-driven narrative insights in the annual Auberle / 412 Youth Zone participant survey report. Covers both the **412YZ** survey (s012–s114 + online) and the **IL** survey (s001–s011). Each cycle produces a `.docx` report from a Node.js script (`05_report_412YZ.js` / `05_report_IL.js`) using the `docx` npm package.

---

## When to Use

- Refining narrative in `report_412YZ_v2.docx` or the IL equivalent
- Working through the report one section at a time to add year-over-year insight
- After `04_analyze_412YZ.py` has been run (analysis sheets must exist)
- Updating or finalizing narrative text hard-coded in the JS report-generation script

---

## Inputs Required Before Starting

| Input | Path | Purpose |
|---|---|---|
| Last year's report | attached `.docx` or `report/example/` | Source of narrative voice, insight patterns, benchmark data |
| This year's analysis | `output/412YZ/analysis_412YZ.xlsx` | Data tables for current cycle |
| This year's draft report | attached `.docx` or `report/412YZ/report_412YZ.docx` | Current draft to refine |
| Report generation script | `scripts/05_report_412YZ.js` | Target file for locked-in changes |

---

## Section Order

Work through sections in this strict order — each section corresponds to a function in `05_report_412YZ.js`:

| # | Section | Script function | Analysis sheet(s) |
|---|---|---|---|
| 0 | Title + intro | `sec_title` | `01_age` |
| 1 | Age distribution | `sec_age` | `01_age` |
| 2 | Gender × Sexual Orientation | `sec_gender_orient` | `02_gender_orient` |
| 3 | Race/Ethnicity (counted once) | `sec_race` | `03_race_once`, `04_race_multi` |
| 4 | Coach satisfaction | `sec_coach_satisfaction` | `05_q1` |
| 5 | Communication | `sec_communication` | `06_communication` |
| 6 | Stable Housing | `sec_housing` | `07_housing`, `08_housing_reasons` |
| 7 | Employment & Education | `sec_education_employment` | `09_education`, `10_employment`, raw CSV |
| 8 | Job tenure | `sec_job_tenure` | `11_job_tenure`, raw CSV |
| 9 | Employment by age | `sec_employment_by_age` | `10_employment` |
| 10 | Job barriers | `sec_job_barriers` | `12_job_barriers`, raw CSV |
| 11 | Reasons left/quit a job | `sec_left_job` | `13_left_job` |
| 12 | Transportation | `sec_transportation` | `14_transport` |
| 13 | Voter registration | `sec_voter_reg` | `15_voter_reg` |
| 14 | Zone visit frequency | `sec_zone_visit` | `16_visit` |
| 15 | Program impact | `sec_program_impact` | `17_impact`, raw CSV |
| 16 | Staff/Peer respect + Environment | `sec_respect_environment` | `18_respect`, `19_environment` |
| 17 | Banking | `sec_banking` | `20_banking`, raw CSV |
| 18 | NPS | `sec_nps` | `21_nps` |
| 19 | Additional comments | `sec_comments` | `22_comments` |

---

## Per-Section Workflow

For **each** section, execute these steps in order:

### Step A — Read last year's section
Load and read the corresponding section in last year's report (attached `.docx` or `report/example/`).
- Extract the key narrative sentences (1–4 sentences per section is typical).
- Identify where Arthur inserted data-driven insights (distribution statements, max/min callouts, trend comparisons).
- Note the exact phrasing pattern (e.g., "About X% of respondents reported…").

### Step B — Read this year's draft narrative + data table
- Read the corresponding section function in `05_report_412YZ.js`.
- Read this year's data from the analysis sheet (use `output/412YZ/analysis_412YZ.xlsx`).
- Identify all numeric findings available: percentages, counts, breakdowns by age/gender.

### Step C — Compare: similar or different?

**Similar** = the dominant finding is directionally the same as last year (e.g., Black youth are still the largest group; most youth still report stable housing).

**Different** = the dominant finding changed materially (e.g., employment rate dropped >5 pp; a new demographic group is now largest; a barrier jumped to #1).

### Step D — Write narrative snippet

**If similar:** Replicate last year's sentence structure, substituting this year's numbers. Keep the same tense, hedging words ("about", "approximately"), and ordering.

**If different:** Write the current-year finding first, then add a sentence explicitly flagging the change: _"This represents a [increase/decrease] from [X%] in [prior year]."_

#### Insight writing rules
1. **Distribution**: Note which category dominates and whether the distribution is concentrated or spread (e.g., "Three-quarters of respondents fall into the 18–20 age group").
2. **Max/min**: Call out the highest and lowest groups explicitly when the gap is meaningful (≥10 pp between top and bottom).
3. **Actionable framing**: End with an implication when data suggests a program opportunity (e.g., "This suggests continued focus on [area] for older youth may be warranted").
4. **Concise**: Target 2–4 sentences per section. Never exceed 6.
5. **No speculation**: Only state what the data shows. Do not infer causes.

### Step E — Output snippet for review
Present the proposed narrative text as a fenced block:

```
SECTION: [Section name]
---
[Proposed narrative text — 2–4 sentences]
---
YEAR-OVER-YEAR NOTE: [Similar / Changed — brief description if changed]
```

### Step E2 — LLM refinement pass
Before showing the snippet to the user, run a self-critique against these criteria and revise once:

1. **Voice match**: Does it sound like the prior-year report? If not, adjust register (professional but plain, not academic).
2. **Precision**: Are the numbers pulled from the data table correct? Verify each figure.
3. **Flow**: Does the narrative lead with the most important finding? Reorder if a supporting detail is leading.
4. **YoY flag**: If the finding changed materially, is the change called out explicitly with the prior-year value and year label?
5. **Length**: Is it within 2–4 sentences? Trim anything that restates the table without adding interpretation.

Revise the snippet based on this check, then present the **revised** version in Step E's format.

### Step F — Incorporate edits
Wait for user feedback. Apply any corrections to the snippet before proceeding.

### Step G — Lock into script
Update the hard-coded string literals inside `makePara()` / `makeBullet()` / `makeHeading()` calls in the corresponding section function in `scripts/05_report_412YZ.js`. Replace only the narrative strings; do not alter table-building logic, data-loading code, or `makeTable()` / `makeCaption()` calls.

---

## Insight Patterns by Section

| Section | Typical last-year insight pattern |
|---|---|
| Age | "Most respondents were age 18–20"; note if share shifts YoY |
| Gender | Pct female vs male; note Trans/NB count |
| Race | "About X% identified as Black"; Multi-racial share as secondary note |
| Coach satisfaction | Top-2 box trend across 5 items; note any dip or rise >3 pp |
| Communication | % reporting "not enough"; most-common frequency for that group |
| Housing | % stable; largest single sleeping category for unstable group |
| Housing reasons | Top 2 reasons; age skew if present |
| Employment & Education | Dual enrollment rate; "not in school and unemployed" share |
| Job tenure | % at job ≥6 months; signal of job stability |
| Employment by age | Age group with highest unemployment; age group fully employed |
| Job barriers | #1 barrier and its %; total respondents affected |
| Reasons left job | Quit vs fired split; top quit sub-reason |
| Transportation | % licensed by age group; % with vehicle access; #1 primary mode |
| Voter registration | % registered overall; 18–20 vs 21–23 gap; top reason not registered |
| Zone visit frequency | % frequent visitors; top reason to come; top barrier for rare visitors |
| Program impact | % helped in at least one area; top 2–3 areas; age group most helped |
| Respect/Environment | Staff vs peer respect comparison; lowest environment item |
| Banking | % with account; top money method; account usage split |
| NPS | Score + tier (Excellent ≥50, Good 0–49, Needs Work <0); pct Promoters |
| Comments | n comments; pull 1–2 representative verbatims |

---

## Year-over-Year Benchmarks Available

The script hardcodes Q1 coach satisfaction benchmarks in `Q1_BENCHMARKS` (Sep-19, Mar-22, Feb-23, Feb-24, Feb-25). For all other sections, prior-year numbers must come from the attached last-year `.docx`. Always state the year when citing prior data.

---

## NPS Interpretation Reference

| Score | Label |
|---|---|
| 50+ | Excellent |
| 0–49 | Good |
| -1 to -49 | Needs Improvement |
| -50 or below | Critical |

---

## Quality Checks Before Locking

- [ ] Every number in the narrative can be traced to `analysis_412YZ.xlsx` or the raw CSV
- [ ] `placeholder_inline()` calls remain for any stat not computable from current data
- [ ] No sentence speculates on causes not in the survey
- [ ] Year-over-year flag is present when a key metric changed direction or by ≥5 pp
- [ ] Running `node scripts/05_report_412YZ.js` produces no errors after changes

---

## Applying This Skill to the IL Report

The same workflow applies to the IL report. Key differences:
- Script: `scripts/05_report_IL.js` (not yet created; will be modeled after `05_report_412YZ.js` using the same `docx` npm package and section function pattern)
- Analysis: `output/IL/analysis_IL.xlsx`
- Survey IDs: s001–s011 (paper) + online
- Some sections will not exist in the IL report (e.g., Zone visit frequency questions differ)
- IL sample is smaller (~30 total); percentages will have wider uncertainty — note this in interpretation
