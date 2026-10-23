# P6 Cost Loader: a Claude skill for cost-loading Primavera P6 from a BOQ

Cost-loading a P6 schedule from a priced Bill of Quantities used to take days of copy-paste. This Claude skill does the repetitive part in seconds and leaves every uncertain call to the planner.

- ☑️ **Matches BOQ lines to P6 activities.** Activities are grouped by trade once floor and zone are stripped from the name ("L05 Zone A - Slab Rebar" → *Slab Rebar*). A construction vocabulary you can extend handles synonyms (shuttering = formwork, sanitaryware = plumbing second fix).
- ☑️ **Splits costs across floors and zones.** It reads "Levels 1 to 17", "Zone A", "Basement 2" or a "LEVEL 05" heading. Splits are equal, by duration, or by floor-area weights, and the parts always add up to the line exactly.
- ☑️ **Flags uncertain matches and BOQ errors.**
  - Flagged and held back until the planner decides: ambiguous matches, preliminaries, provisional sums, negative lines, floors the schedule doesn't have, and rows below a bill total.
  - Reported as BOQ errors: qty × rate ≠ amount, wrong bill totals, duplicates, rate outliers and unpriced lines.
- ☑️ **Checks totals to the cent.** BOQ total = loaded + pending review + excluded → **difference 0.00**. The P6 import file is read back from disk and must equal the loaded total; if any check fails, the run stops.

Automation doesn't replace the planner. Uncertain matches stay flagged, Claude suggests a decision for each, and the planner makes the final call.

## What you get

| File | Purpose |
|---|---|
| `<name>_P6_import_TASKRSRC.xlsx` | P6 resource-assignment import: one row per activity. Cost is carried by a nonlabor resource at 1.00/unit, so budgeted units = budgeted cost. |
| `<name>_review.xlsx` | Review workbook with these sheets: Summary (reconciliation), **Review Queue** (Decision dropdown for the planner), BOQ Lines, Allocations, Activity Totals, BOQ Issues, Unloaded Activities. |
| `<name>_assignments.csv` | Same assignments as plain CSV |
| `<name>_summary.json` | Machine-readable checks, issues and review queue (used by Claude) |

## Install

**Claude (claude.ai / Claude desktop):** download [`p6-cost-loader.skill`](p6-cost-loader.skill) and add it under Settings › Capabilities › Skills. You can also open it in a chat and click **Save skill**.

**Claude Code:** copy the `p6-cost-loader/` folder into `~/.claude/skills/`, or into `.claude/skills/` inside a project.

Requires Python 3.8+ and `openpyxl` (`pip install openpyxl`).

Then just ask, for example: *"Cost-load my P6 programme from this BOQ: schedule.xer and BOQ.xlsx, equal split."*

## Run the engine directly

```bash
# 1. Check how the BOQ and schedule were read (columns, bill totals, floors, work types)
python p6-cost-loader/scripts/p6_cost_loader.py inspect --boq BOQ.xlsx --p6 schedule.xer

# 2. Match, split, reconcile, write outputs
python p6-cost-loader/scripts/p6_cost_loader.py run --boq BOQ.xlsx --p6 schedule.xer --out out --name MyProject

# 3. After the planner fills the Decision column in the Review Queue
python p6-cost-loader/scripts/p6_cost_loader.py run --boq BOQ.xlsx --p6 schedule.xer --out out --name MyProject \
       --decisions out/MyProject_review.xlsx
```

Inputs:
- **BOQ:** `.xlsx`/`.csv`, with multiple bills and headers auto-detected.
- **Schedule:** `.xer`, a P6 Excel export, or any activity list with Activity ID and Name.

Useful options:
- `--split duration|weights --weights floors.csv`
- `--rules rules.csv`
- `--suggestions`
- `--prelims-activity`
- `--resource-id`
- `--cost-account`
- `--p6-template <your P6 TASKRSRC export>`
- `--amount-basis computed`

See [`references/matching_and_review.md`](p6-cost-loader/references/matching_and_review.md) for file formats and tuning, and [`references/p6_import.md`](p6-cost-loader/references/p6_import.md) for importing into P6 and verifying the result.

## Try it on the demo project

`examples/` holds a synthetic 17-storey tower:
- `TWR17_schedule.xer` (679 activities) and `TWR17_P6_export.xlsx` (the same schedule as a P6 Excel export)
- `TWR17_BOQ.xlsx` (9 bills, 758 priced lines, with deliberate errors)
- `truth.csv` (the intended activity for every line)

```bash
python p6-cost-loader/scripts/p6_cost_loader.py run --boq examples/TWR17_BOQ.xlsx --p6 examples/TWR17_schedule.xer --out out
python examples/score.py examples/truth.csv out/cost_loading_review.xlsx
```

Result on the demo:

| | |
|---|---|
| Run time | ~1–2 s |
| Loaded automatically | 740 lines, 82% of value, **0 wrong matches** |
| Held for review | 17 lines: preliminaries, provisional sums, a "Level 18" item in a 17-floor building, an omission row, BWIC, signage, louvres |
| Planted BOQ errors | all caught: 2 arithmetic, 2 bill-total mismatches, rows below a bill total, a blank amount, an unpriced line, an "Incl" line |
| Reconciliation difference | **0.00** |

`examples/make_test_data.py` regenerates the demo.

## Importing into P6, in short

1. In P6, create the resource `BOQ-COST`: Nonlabor, Price/Unit 1.00, *Drive activity dates* off.
2. Back up the project, then use File › Import › Spreadsheet (XLS) › Resource Assignments.
3. Check that the project's *Budgeted Total Cost* equals the import total on the Summary sheet, to the cent.

If P6 rejects the column headers, export one resource assignment from your own P6 and rerun with `--p6-template`.

## Limits

- It doesn't time-phase cost; P6 spreads each assignment over the activity dates.
- It doesn't check quantities against drawings or judge whether rates are fair.
- Legacy `.xls` files must be re-saved as `.xlsx`.
