---
name: p6-cost-loader
description: Cost-load a Primavera P6 schedule from a priced Bill of Quantities (BOQ). Matches every BOQ line to P6 activities, splits costs across floors and zones to the exact cent, flags uncertain matches and BOQ errors (qty x rate, subtotals, duplicates, negatives, unpriced lines) for the planner, reconciles totals to zero difference, and writes a P6 resource-assignment import file plus an Excel review workbook. Use this whenever someone wants to cost-load or resource-load a schedule with money, map BOQ / bill items / cost breakdown to P6 or Primavera activities, put budgeted cost on activities, prepare an earned-value or cash-flow baseline from a BOQ, or has an XER (or P6 activity export) alongside a priced BOQ - even if they never say "cost loading".
---

# P6 Cost Loader

Cost-loading a schedule is mostly repetitive matching: hundreds of BOQ lines, each belonging to one kind of work, which the schedule breaks out by floor and zone. This skill automates that and leaves the planner in charge of everything uncertain:

- **Matches** each priced BOQ line to a P6 *work type*, meaning the activities that share a name once floor/zone are stripped ("L05 Zone A - Slab Rebar" → *Slab Rebar*).
- **Splits** the line's value across that work type's activities on the floors and zones the BOQ names (or across all of them when it names none), using largest-remainder rounding so the parts add up to the line exactly.
- **Flags** anything uncertain (ambiguous or low-confidence matches, preliminaries, negative lines, floors the schedule doesn't have) and every BOQ arithmetic problem. Flagged lines are *not* loaded until the planner decides.
- **Reconciles**: BOQ total = loaded + pending review + excluded, difference 0.00. Allocations = loaded. The import file read back from disk = loaded.

The planner makes the final call on flagged lines. Your job is to run the engine, make sure the inputs were read correctly, give well-reasoned suggestions on the review queue, and report clearly. Wrong cost loading quietly corrupts earned value and cash-flow forecasts for the rest of the project, so don't force a match to raise the automation percentage.

## Inputs

| Input | Formats | Notes |
|---|---|---|
| BOQ | `.xlsx`/`.xlsm`/`.csv`, one or many bills/sheets, several files allowed | Header row and columns auto-detected (Item, Description, Unit, Qty, Rate, Amount; optional Location, Bill/Section). Summary/cover sheets skipped. |
| Schedule | `.xer` (best), P6 Excel export (`TASK` sheet), or any table with Activity ID + Activity Name | WBS, activity type and original duration used when present. Milestones never receive cost. |
| Optional | weights CSV, rules CSV, decisions/suggestions | See `references/matching_and_review.md`. |

Legacy `.xls` must be re-saved as `.xlsx`. If the user hasn't said these, use the defaults and mention them; don't block on them:
- **Cost resource** (default `BOQ-COST`): a P6 nonlabor resource at price/unit 1.00, so budgeted units = budgeted cost. This is the standard cost-loading pattern.
- **Split basis** (default `equal`): `duration` weights by original duration; `weights` uses a floor/zone CSV (e.g. floor areas). This is the right choice when floors differ in size.
- **Preliminaries / provisional sums** go to the review queue by default. Typical choices are an LOE activity, spreading over all activities, or excluding them.

## Workflow

The script is `scripts/p6_cost_loader.py`. It needs Python 3.8+ and `openpyxl`. Run it with the skill directory's path; it takes about a second for ~1,000 lines.

### 1. Inspect the parsing

```bash
python scripts/p6_cost_loader.py inspect --boq BOQ.xlsx --p6 schedule.xer
```

Check before going further, because every later number depends on this:
- Every bill sheet was found with the right columns, and no real bill was skipped. Fix with `--boq-sheets`, `--boq-cols "description=C,qty=E,rate=F,amount=G"`, or `--boq-header-row N`.
- Each bill's **parsed total vs its printed bill total**. `inspect` prints both and marks any DIFF. Then compare the overall total with the BOQ's summary/grand-total page, if it has one. A difference is usually a BOQ error (arithmetic, a wrong total row, or rows sitting below the bill total) rather than a parsing problem. Either way, explain every difference before moving on, because it is the first thing the QS will ask about.
- The floors list looks like the building (B2…L17, RF), and zones were detected if the schedule uses them.
- The work types are meaningful trades, not fragments. If activity names carry no trade (e.g. just "L05 Zone A"), the WBS leaf is used instead. Say so if that happens.
- For an XER holding several projects, pass `--project <short name>`.

### 2. Run

```bash
python scripts/p6_cost_loader.py run --boq BOQ.xlsx --p6 schedule.xer --out <outdir> --name <project>
```

Useful options: `--split duration|weights --weights floors.csv`, `--rules rules.csv`, `--decisions <edited review workbook or CSV>`, `--suggestions suggestions.csv`, `--prelims-activity <LOE id>`, `--resource-id`, `--cost-account`, `--p6-template <TASKRSRC export>`, `--amount-basis computed` (load qty×rate instead of the stated amount when they disagree), and `--auto-threshold` / `--margin` (matching strictness).

Outputs in `<outdir>`:
- `<name>_review.xlsx` is for the planner. Sheets: Summary (reconciliation, coverage, issue counts), **Review Queue** (editable Decision / Activity IDs / Planner Note columns), BOQ Lines, Allocations, Activity Totals, BOQ Issues, Unloaded Activities, Parsing.
- `<name>_P6_import_TASKRSRC.xlsx` is the P6 resource-assignment import: one assignment of the cost resource per activity.
- `<name>_assignments.csv` is the same data in plain form.
- `<name>_summary.json` is for you: checks, reconciliation, issues, and the full review queue with top-3 candidates per line.

If the script exits with `RECONCILIATION FAILED`, or any `checks` value in the summary JSON is false, stop. Tell the user the outputs must not be imported, and investigate. This should never happen; if it does, it's a bug worth reporting with the inputs.

### 3. Spot-check the automatic matches

`auto_check` in the summary JSON lists each distinct BOQ description pattern that was auto-loaded, with its work type, line count, value and lowest score, weakest first. Read the first 20–30 entries with a construction eye. Is "LED light fittings" on second fix, and are FCUs on HVAC rather than ceilings? A wrong family is fixed for all floors at once with a rule or lexicon entry (planner's OK first), or per line with a decision.

### 4. Work the review queue

This is where your judgement adds the most. Read `review_queue` from the summary JSON (sorted by value, largest first). For each line, use construction knowledge to propose a decision:
- `ACCEPT` / `ACCEPT2` / `ACCEPT3` take candidate 1/2/3, split across its activities on the line's floors/zones.
- `MANUAL` with `Activity IDs` takes explicit IDs, wildcards (`A45*`), weights (`A100:60; A200:40`), or whole trades (`worktype:first fix; worktype:basement mep`). `*` spreads over every non-milestone activity, which is a common treatment for preliminaries when used with `--split duration`.
- `EXCLUDE` keeps the line out of the schedule (e.g. contingency, provisional sums held outside the baseline). It stays visible in the reconciliation.

Write the proposals to a CSV with columns `BOQ Key, Decision, Activity IDs, Note`, giving a one-line reason in Note, and rerun with `--suggestions that.csv`. They then appear beside the planner's Decision column in the Review Queue sheet. **Pass them as `--decisions` only when the user explicitly approves them** ("accept your suggestions", "go with those"). The budget is the planner's call, and a review step that rubber-stamps itself is no review.

When the same ambiguity repeats across floors (e.g. "Fan coil units, ceiling concealed" on 18 floors), propose one **rule** instead of 18 decisions, e.g. `fan coil,worktype:^hvac`. Show it to the user and apply it with `--rules` once they agree. Rules and the lexicon are described in `references/matching_and_review.md`.

For a large queue (hundreds of lines), work through it in value order and say how far you got. Value is concentrated in the top lines.

### 5. Report

Keep it short and concrete. Until the planner has decided the queue, the import file holds only the loaded lines, so call it interim and give its percentage. Use this shape:

```
## Cost loading – <project>
- <N> priced BOQ lines → <M> activities across <W> work types (<s> s)
- Loaded automatically: <x>% of value (<n> lines)  ·  Pending review: <n> lines / <value>  ·  Excluded: <value>
- Reconciliation: BOQ <total> = loaded + pending + excluded → difference 0.00 ✓ · import file = loaded ✓
- BOQ errors to raise with the QS: <arithmetic / subtotal / duplicate items with sheet!row and amounts>
- Biggest review items: <top 3–5 with your suggestion>
- Files: <review workbook>, <P6 import>, <csv>
- Next: planner fills Decision in the Review Queue → rerun with --decisions → import (see steps)
```

State the BOQ errors plainly, as findings for the quantity surveyor. The stated Amount is what gets loaded unless `--amount-basis computed` is used, so a wrong amount flows into P6 until someone fixes the BOQ. If the bill/summary totals agree with qty × rate rather than the stated amounts, say so and offer `--amount-basis computed`. Rows found below a bill's final total are held in review (`BELOW_BILL_TOTAL`) because they are not in the contract sum.

### 6. Iterate and import

- The planner edits the **Review Queue** sheet (Decision column has a dropdown). Rerun with `--decisions <that workbook>`. Decisions are keyed by `BOQ Key` (`Sheet!R<row>`). If rows are inserted in the BOQ, keys shift, so rerun from a fresh review workbook.
- Decisions that can't be applied (unknown IDs, milestones, a missing candidate) are listed as `BAD_DECISION` in BOQ Issues and in the console. Tell the user.
- For the P6 side (creating the cost resource, importing, verifying totals, re-import pitfalls), read `references/p6_import.md` and walk the user through it.

## What this does not do

It does not time-phase cost (P6 spreads each assignment over the activity's dates and resource curve), check quantities against drawings, or judge whether rates are fair. It flags rate outliers only between identical items. Say so if the user expects those.
