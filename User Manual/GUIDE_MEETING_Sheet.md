# MEETING Sheet Guide

## When

Use `MEETING` to record whether an employee attended a meeting on each day of
the current calculation period.

Use it only for employees whose `Need Meeting ?` flag in `CONFIG` is `YES`.
Employees with `NO`, `-`, or a blank flag are not included in the active
MEETING layout.

Be aware that the Early Start checkbox column at `IN` will not be working for 
individual with `Need Meeting ?` flag as `Yes` in the `CONFIG`

## Where

- Open the `MEETING` sheet.
- Row 1 contains employee names.
- Rows 2–32 represent days 1–31 of the calculation period.
- Enter attendance in the cell where the employee column and day row meet.

## How

1. In `CONFIG`, set the employee's `Need Meeting ?` flag to `YES`.
2. Select **Salary Tool → Sync Employee Tabs**.
3. On `MEETING`, choose one of the dropdown values for each applicable day:
   - `✓` — attended
   - `✗` — did not attend
   - `MC` — medical leave
   - `PH` — public holiday
4. Leave the cell blank when no meeting status should be recorded.
5. Run Sync after changing employee names, employee order, or meeting flags.

## Important Sync behavior

Sync rebuilds the MEETING attendance layout from `CONFIG`:

- Active columns contain only `YES` employees, in CONFIG order.
- Reordering employees moves their attendance data with their names.
- A supported employee rename preserves the employee's MEETING attendance.
- A new employee receives a blank MEETING column.
- An employee changed from `YES` to another flag falls out of the active layout;
  that column's displayed header and attendance data are cleared.
- Dropdown validation is reapplied to all available MEETING column positions.

## Side notes: changes to the whole workbook

Running **Sync Employee Tabs** for a MEETING change also performs the normal
whole-workbook synchronization:

- Employee tabs are created, updated, or renamed to match `CONFIG`.
- `IN` and `OUT` employee pairs are reordered to match `CONFIG`.
- The `SUM` dashboard is rebuilt in the same employee order, with formulas and
  links refreshed.
- Employee summary blocks and formula recalculation are refreshed.
- Existing employee attendance data is preserved by employee name where the
  layout supports it.

Therefore, review the Sync confirmation message after running it. Do not use
Sync merely to edit a MEETING cell; edit the MEETING cell directly unless
CONFIG or employee layout has changed.

## Cautions

- Do not type employee names into MEETING headers; Sync owns the headers.
- Do not insert or delete employee columns manually.
- Maintain employee names and meeting flags in `CONFIG`, which is the source of
  truth.
- Before changing a `YES` flag to another value, verify that the employee's
  MEETING attendance is no longer needed in the active layout; the next Sync
  clears that displayed column and its data.
- A duplicate or reserved employee name in `CONFIG` aborts Sync and must be
  corrected before the workbook can be synchronized.
