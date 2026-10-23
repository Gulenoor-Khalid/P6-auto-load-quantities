# Getting the costs into P6

The import file is `<name>_P6_import_TASKRSRC.xlsx`. It has one sheet, `TASKRSRC` (resource assignments), with one row per activity. Row 1 holds P6 field codes and row 2 holds labels, as in P6's own Excel export:

| task_code | rsrc_id | target_qty | target_cost | (acct_id) |
|---|---|---|---|---|
| Activity ID | Resource ID | Budgeted Units | Budgeted Cost | Cost Account |

Units equal cost because the resource is priced at 1.00 per unit.

## 1. Create the cost resource (once per database)

In P6 Professional, go to Enterprise › Resources › Add:
- Resource ID `BOQ-COST`, or whatever was passed as `--resource-id`. Name it something like "BOQ Cost".
- Type **Nonlabor**. Some organisations use Material, which also works.
- Units & Prices: **Price/Unit = 1.00**.
- Uncheck **Drive activity dates by default**, so the cost resource never moves the schedule.
- Calculate costs from units: on.

Excel import does not create resources, so the resource must exist before importing.

## 2. Use your P6 version's exact layout (recommended)

Field codes on TASKRSRC differ slightly between P6 versions and configurations; some expect `task_id` instead of `task_code`. The safest approach:
1. In P6, assign `BOQ-COST` to any one activity.
2. Use File › Export › Spreadsheet (XLS) › Resource Assignments to export that project.
3. Rerun the loader with `--p6-template <that export>.xlsx`. The script keeps the export's own headers and other sheets (e.g. USERDATA), clears its data rows, and writes the cost rows into the matching columns (`task_code`/`task_id`, `rsrc_id`, `target_qty`, `target_cost`, `acct_id`).

## 3. Import

1. **Back up the project** or import into a copy first.
2. If an earlier cost-loading run was imported, delete the existing `BOQ-COST` assignments first. Re-importing can add assignments rather than replace them, which would double the cost.
3. Use File › Import › Spreadsheet (XLS) › Resource Assignments, select the file, and select the target project.
4. Don't rename headers, don't add formulas, and don't edit row 2.

## 4. Verify in P6 (always)

- In the Activities view, add the **Budgeted Total Cost** column. The project total must equal the "P6 import file total" on the Summary sheet, to the cent.
- In the Resource Assignments view, filter on `BOQ-COST` and check that the number of assignments equals the import row count.
- Check a few activities against the **Activity Totals** sheet.
- If the totals differ, look for activities that weren't found (ID typos, wrong project), duplicate assignments, or activities whose duration type recalculated units. For the last case, use *Fixed Duration & Units* for cost-loaded activities.

## Notes

- **Activity types:** LOE activities can take cost (useful for preliminaries). Milestones and WBS summaries cannot, so the script never assigns to them.
- **Cost accounts:** `--cost-account CA-01` stamps one account on every row. For per-trade accounts, split the run by bill/section (one run per `--boq-sheets` group) or edit the CSV before import.
- **Expenses instead of resources:** some organisations load BOQ cost as project expenses. `<name>_assignments.csv` gives activity ID and amount for that route. Expenses are usually entered through the Expenses window or an XER/API tool rather than Excel import.
- **P6 EPPM (web):** use the same resource setup. Import resource assignments through the EPPM spreadsheet import or the P6 Integration API using the CSV.
