# Reset Manual Changes Guide

Use `Salary Tool → Reset Manual Changes` to reset regular employee attendance
inputs and manual colours across the workbook. This is separate from
`Reset Manual Changes (Office)`.

The reset clears all employee time entries on `IN` and `OUT`, including text and
`x`/`X` markers. Check the scope below before using it; it is not limited to the
selected cells or the currently open tab.

## When

- Leave the `OUT` rest-control cell blank for Auto:
  - Weekday: 1 hour rest.
  - Saturday: 0 hours rest.
  - Sunday: 0 hours rest.
- Select `0h` when no rest was taken.
- Select `1h` when a one-hour rest was taken, including on Saturdays.
- Incomplete attendance calculates 0 rest until both check-in and checkout exist.

## Where

- Enter overrides on `OUT`.
- Use the rest-control cell immediately to the right of each employee checkout column.
- Do not edit the employee sheet's calculated `Rest Duration` column.

## How

1. Enter or correct check-in and checkout times normally.
2. Leave the control blank for Auto, or select `0h`/`1h`.
3. Orange checkout text indicates effective 0-hour rest; normal text indicates effective 1-hour rest or incomplete attendance.
4. To reset time entries, manual colours, checkbox selections, rest overrides,
   and meeting attendance together, select
   `Salary Tool → Reset Manual Changes` (formerly `Reset Rest Overrides`).
   See the scope below before using it.
5. Changing Month or Year does not reset overrides automatically.

Invalid pasted values produce a visible calculation error and should be corrected.

## Reset Manual Changes

This action applies to the workbook regardless of which tab is open:

- Resets manual font colours to default and cell fills to no fill on `IN`, `OUT`,
  and regular employee sheets listed in `CONFIG!A5:A44`. It covers all cells below
  row 1, including date-column formatting, totals, and blank areas.
- Clears every entry in the employee time columns on `IN` and `OUT`, rows 2–32.
  This includes numbers, times, text, `x`/`X` markers, and any formulas placed in
  those time cells. Only time columns with a nonblank employee header are cleared;
  unnamed spare time columns, row 1 headers, and time entries below row 32 are kept.
  Column A date values and formulas are not cleared.
- Unticks manually controlled checkboxes on `IN` and `OUT` in rows 2–32.
  The checkbox controls remain visible with their unchecked value, and
  formula-controlled checkboxes keep their formulas. In the current layout,
  `IN` has checkboxes; `OUT` has rest dropdowns.
- Clears active employees' `OUT` rest overrides in rows 2–32 back to blank/Auto,
  preserving the `0h`/`1h` dropdown controls. Formula cells in these control columns
  are skipped. Active employees come from `CONFIG!A5:A44` in CONFIG order.
- Resets `MEETING!B2:AO32` to `✓` wherever both the row's date and the column's
  employee header are present. Cells within that range without a date or employee
  header are cleared. Attendance dropdown controls, headers, and date formulas remain.

Row 1 headers, office employee sheets, and other tabs such as `CONFIG`, `MEETING`,
`SUM`, and templates are excluded from the colour reset. Employee-sheet checkbox
values are not reset.

Office employee sheets and the office reset action are not changed by this reset.
Regular employee-sheet values and formulas are preserved; their calculated results
may change when attendance inputs are cleared.

Date-column values and formulas, conditional-formatting rules, number formats,
borders, other styling, validation, and protections are preserved. Contents of the
target time cells and MEETING attendance cells are cleared or replaced as described
above. Automatic colours and calculated results continue to follow the existing
rules after inputs reset. For example, an automatic colour may change when its
manual checkbox input is unticked.

1. Open `Salary Tool → Reset Manual Changes`. Reload the workbook if the new menu
   name is not visible.
2. Wait for `RESET OK`, then press `OK`.

The action clears manual inputs and colours across all target sheets; it is not
limited to the selected cells. You must have edit access to all target cells,
including protected employee-sheet areas. If access is insufficient, the reset
stops before making changes and reports the affected sheet.
