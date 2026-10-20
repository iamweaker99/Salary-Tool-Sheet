# Data, Formatting, Checkbox and Dropdown Cleanup Guide

The `Salary Tool` menu provides three separate cleanup actions:

- `Remove data`
- `Remove formatting, checkbox & dropdown`
- `Remove data & formatting (Office)`

These replace the former combined `Reset Manual Changes` action and the former
`Reset Manual Changes (Office)` menu name. Each action applies to its target
sheets regardless of which tab or cells are selected.

## Confirmation before changes

1. Open `Salary Tool` and select the required action. Reload the workbook if the
   updated menu names are not visible.
2. Read the bullet-point scope in the confirmation popup, ending with `Continue?`.
3. Choose `OK` to proceed, or `Cancel` to exit without changes. Closing the popup
   also exits without changes.
4. After a successful cleanup, read the completion message and press `OK`.

## Remove data

- **IN/OUT:** Clear every entry in named employee time columns, rows 2–32.
  This includes numbers, times, text, `x`/`X` markers, and any formulas placed in
  those time cells. A nonblank employee header identifies each time column.
- **MEETING:** Set `MEETING!B2:AO32` to `✓` wherever both the row's date and the
  column's employee header are present. Clear other attendance cells within that
  range. Dropdown controls remain.

This action preserves direct font/background colours, IN checkbox selections,
OUT rest dropdown selections, date-column values and formulas, row 1 headers,
unnamed spare time columns, and time entries below row 32. Regular employee sheets
and office employee sheets are not directly edited; calculated results and
automatic colours may change when attendance inputs change.

## Remove formatting, checkbox & dropdown

- **Regular employee sheets:** Reset direct font colours to default and cell fills
  to no fill on sheets listed in `CONFIG!A5:A44`. This covers all cells below row 1,
  including totals and blank areas.
- **IN/OUT:** Reset direct font/background colours on all cells below row 1,
  including date-column formatting and blank areas.
- **IN:** Uncheck manually controlled checkboxes in rows 2–32, including spare
  employee positions. Checkbox controls remain visible with their unchecked value.
  Formula-controlled checkboxes keep their formulas.
- **OUT:** Clear active employees' rest dropdown selections in rows 2–32 back to
  blank/Auto. Preserve `0h`/`1h` dropdown controls and skip formula cells. Active
  employees come from `CONFIG!A5:A44` in CONFIG order.

Employee time entries on IN/OUT and MEETING attendance are not changed. Date values
and formulas are preserved. Row 1 headers, office employee sheets, `CONFIG`,
`MEETING`, `SUM`, and templates are excluded from the colour reset.
Regular employee-sheet values, formulas, and checkbox values are preserved.

Despite its name, this action resets only direct font/background colours and the
specified controls. Conditional-formatting rules, number formats, borders, other
styling, validation, and protections remain. Automatic colours and calculated
results may change when checkboxes or rest dropdown selections reset.

## Remove data & formatting (Office)

This is the former office reset with a new menu name; its cleanup scope is unchanged.
It uses the office employee names in `CONFIG!K6:K26`.

- **Office employee sheets:** Clear manual entries in `B2:G32` and `M2:M32`,
  preserving formula cells.
- **Office employee sheets:** Reset direct font and background colours across the
  entire sheet, including row 1.

Conditional-formatting rules, number formats, borders, other styling, validation,
and protections remain. Regular employee sheets, IN, OUT, and MEETING are not edited.
Missing office sheets are listed in the completion message.

## Edit access

Each action checks edit access to its target ranges before making changes.
You need access to the affected protected areas as well. If access is insufficient,
the action stops before making changes and reports the affected sheet or range.

## OUT rest duration and overrides

- Enter overrides in the `OUT` rest-control cell immediately to the right of each
  employee checkout column. Do not edit the employee sheet's calculated
  `Rest Duration` column.
- Leave the control blank for Auto: weekdays use 1 hour of rest; Saturdays and
  Sundays use 0 hours.
- Select `0h` when no rest was taken, or `1h` when a one-hour rest was taken,
  including on Saturdays.
- Incomplete attendance calculates 0 rest until both check-in and checkout exist.
- Orange checkout text indicates effective 0-hour rest; normal text indicates
  effective 1-hour rest or incomplete attendance.
- Use `Remove formatting, checkbox & dropdown` to clear rest overrides back to
  Auto. `Remove data` preserves their selections.
- Changing Month or Year does not reset overrides automatically.

Invalid pasted values produce a visible calculation error and should be corrected.
