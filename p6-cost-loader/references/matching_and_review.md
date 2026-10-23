# Matching, review files and tuning

## Contents
1. How a BOQ line is matched
2. Statuses
3. BOQ checks
4. File formats: decisions, suggestions, rules, weights, lexicon
5. Tuning
6. Known limits

## 1. How a BOQ line is matched

1. **Schedule side.** Each activity's floor and zone are read from its name, falling back to its WBS path. They're then stripped to give the *work type* ("L05 Zone A - Slab Rebar" → "slab rebar"; "Lift Installation - Lift 3" → "lift installation lift"). Activities with the same work type form a group. Milestones and WBS summaries are left out.
2. **BOQ side.** The text is the description, weighted most heavily (earlier words count more because the item comes first and qualifiers trail), plus the headings above it, the bill/section, and any location column. Location comes from the most specific text that has one: description first, then headings ("LEVEL 05"), then the section.
3. **Vocabulary.** Both sides go through `assets/trade_lexicon.json`, which rewrites synonyms to canonical trade terms (shuttering → formwork, sanitary ware → plumbing + secondfix). Stop-words, numbers and units are dropped.
4. **Score.** TF-IDF cosine between the line and every work type. A work type on none of the line's floors gets a ×0.6 penalty.
5. **Decide.** A line is loaded automatically (`AUTO`) only if:
   - the top score ≥ `--auto-threshold` (0.45),
   - it leads the runner-up by ≥ `--margin` (0.08),
   - the description itself shares a term with the work type (a match on the bill heading alone is not enough),
   - and the line's floors/zones exist for that work type.
   Otherwise it goes to `REVIEW` with a reason.
6. **Split.** The value goes to the matching activities on the line's floors/zones (all of them if the line names none), by `equal` / `duration` / `weights`. Splitting works in integer cents with largest-remainder rounding, so the parts always sum to the line exactly. Zone filtering applies only if the schedule has zones for that work type, and floor filtering only if it has floors.

Precedence per line: planner decision → rule → `--prelims-activity` → below-bill-total flag → preliminaries flag → negative-amount flag → automatic match.

## 2. Statuses

| Status | Meaning | In the P6 import? |
|---|---|---|
| AUTO | Confident automatic match | yes |
| RULE | Matched by a rules-file line or `--prelims-activity` | yes |
| DECIDED | Planner decision applied | yes |
| REVIEW | Waiting for the planner (reason column says why) | no |
| EXCLUDED | Planner/rule excluded it from the schedule | no |
| NIL | Zero value (unpriced or 0.00) | no |

Reconciliation identity: `BOQ total = AUTO + RULE + DECIDED + REVIEW + EXCLUDED` (NIL is 0). The Summary sheet shows the difference, which must be 0.00.

Typical REVIEW reasons:
- preliminaries / provisional sum / BWIC / contingency
- ambiguous (top two within margin)
- low confidence
- no credible match (a scope item missing from the schedule)
- matched on heading only
- BOQ floor not in the schedule (e.g. "Level 18" in a 17-floor schedule)
- negative amount (omission/adjustment)

## 3. BOQ checks (BOQ Issues sheet)

| Code | Severity | What |
|---|---|---|
| ARITHMETIC | ERROR | Qty × Rate differs from Amount by more than `--tolerance` (0.01) |
| SUBTOTAL_MISMATCH | ERROR | A "Total / Carried to summary / Collection" row ≠ the priced lines above it (checked against the segment since the previous total and against the sheet to date; brought-forward rows ignored) |
| BELOW_BILL_TOTAL | WARNING | Priced row below a "Total carried to Summary / Bill total" row, so it isn't in the carried total. The line is held in REVIEW. |
| AMOUNT_MISSING | WARNING | Amount blank; qty × rate loaded instead |
| UNPRICED | WARNING | Quantity with no rate or amount |
| NEGATIVE / NEGATIVE_QTY | WARNING | Omissions, credits, sign errors |
| ZERO_QTY_WITH_AMOUNT | WARNING | Qty 0 but an amount |
| POSSIBLE_DUPLICATE | WARNING | Identical description, context, qty, rate and amount |
| RATE_OUTLIER | WARNING | Same description + unit (3 or more instances) priced more than 25% off the median rate |
| AMOUNT_WITHOUT_QTY | INFO | Amount with no quantity and a unit that isn't a lump-sum unit |
| INCLUDED_ELSEWHERE | INFO | Priced as "Incl", "N/A", "by others", etc. |
| DUPLICATE_ACTIVITY_ID, BAD_DECISION, MISSING_WEIGHT | WARNING | Schedule / decision / weights problems |

Amounts are rounded half-up to cents per line. Any sub-cent difference against the raw BOQ is shown on the Summary sheet.

## 4. File formats

**Decisions / suggestions** use a CSV, or the edited review workbook (its *Review Queue* sheet is read):

```csv
BOQ Key,Decision,Activity IDs,Note
Bill 1 - Preliminaries!R6,MANUAL,A1030,Site staff to PM LOE
Bill 7 - MEP Services!R12,ACCEPT,,FCUs belong to HVAC not ceilings
Bill 9 - Provisional Sums!R6,EXCLUDE,,Contingency held outside baseline
Bill 1 - Preliminaries!R10,MANUAL,*,Crane hire spread over all activities by duration
Bill 6 - Facade!R24,MANUAL,A4560:60;A4570:40,Split by planner
```

- Decisions: `ACCEPT` (= candidate 1), `ACCEPT2`, `ACCEPT3`, `MANUAL`, `EXCLUDE`.
- Activity IDs accept exact IDs, wildcards (`*`, `?`, `[...]`), optional `:weight`, and `worktype:<regex>` (every activity of a trade, all floors). Separate items with `;`. Wildcards and work types skip milestones. A long list in one cell is fine.
- `--suggestions` fills the "Suggested Decision / Suggested IDs / Suggestion Note" columns without applying anything. `--decisions` applies.
- A decision may also override an AUTO line, which is useful when a spot check finds a wrong automatic match.

**Rules** (`--rules rules.csv`) are applied in order, first match wins. `pattern` is a case-insensitive regex over description + headings + section.

```csv
pattern,target,note
fan coil,worktype:^hvac,FCUs are HVAC even when ceiling concealed
builder'?s work|bwic,activity:*,Spread BWIC over everything
\bcontingenc,exclude,Held outside the baseline
lift,activity:A77*,Lifts 1-4
```

Targets:
- `worktype:<regex>` matches against the work type or label, then floor/zone filtering applies as usual.
- `activity:<ids/wildcards>` names activities directly.
- `exclude` keeps the line out.

**Weights** (`--split weights --weights floors.csv`) take one row per floor, optionally per zone. Floor names are parsed like the schedule's ("Level 5", "L05", "GF", "Roof"):

```csv
floor,zone,weight
GF,,1500
Level 1,A,600
Level 1,B,400
```

An activity's weight is its floor/zone weight, or the floor weight shared across its zones. Activities with no weight get 1 and are reported as MISSING_WEIGHT.

**Lexicon** (`--lexicon my.json`, default `assets/trade_lexicon.json`) is `{ "canonical": ["synonym", "multi word phrase", ...] }`. A phrase listed under two terms carries both. Add local trade names, abbreviations and languages here. It's the most effective lever when a whole family of lines is mismatched or ambiguous.

## 5. Tuning

- **Spot-check AUTO lines first.** Sort the BOQ Lines sheet by Score ascending and look at the lowest-scoring AUTO lines. If they're wrong, raise `--auto-threshold` or `--margin`, or fix the vocabulary.
- **Many correct-looking REVIEW lines** marked "ambiguous": a missing lexicon term or a repeating pattern is usually the cause. A rule or lexicon entry fixes it for every floor at once.
- **Schedule naming matters.** Activities that differ only by floor/zone group cleanly. If the same trade is named differently on different floors ("Blockwork L5" / "Block Work L6"), the lexicon usually unifies it. Check the work-type list in `inspect`.
- **First fix vs second fix, rebar vs concrete vs formwork:** keep these as separate activities in P6 and as separate BOQ lines, and the matcher will separate them. If the schedule bundles them (e.g. a single "Slab L05"), all those BOQ lines correctly land on that one activity.

## 6. Known limits

- Multi-row descriptions (text above, numbers on the last row) are handled through the heading context, not merged into one line.
- Merged header cells, two-row headers and European decimal commas may need `--boq-cols` / `--boq-header-row`, or a quick cleanup of the BOQ.
- Formula-only BOQs saved by a tool that didn't store calculated values read as blank. Open the file in Excel and save it once.
- Collection pages inside a bill sheet are recognised by "Page N / Collection / Total" wording. Unusual wording can double-count, and the subtotal check will flag it.
