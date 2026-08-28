# SUM Sheet Guide

## When

Use `SUM` when you need to review the current month's employee totals in one place:

- Normal Working Hours
- Overtime
- Overday

Run **Salary Tool → Sync Employee Tabs** after changing the employee names or order in `CONFIG`, adding an employee, removing an employee, or correcting a rename. Sync updates the dashboard to match the current `CONFIG` list.

## Where

- Open the `SUM` sheet to see the dashboard.
- The period shown in `SUM!B1` and `SUM!D1` comes directly from the Month and Year cells in `CONFIG`.
- Employee rows begin at row 4 and follow the active employee order in `CONFIG`.
- On an employee sheet, the monthly summary is in `I34:M34`:
  - `J34`: Normal Working Hours
  - `K34`: Overtime
  - `L34`: Overday
  - `M34`: link back to that employee's row on `SUM`

## How

1. Maintain the employee list and order in `CONFIG`.
2. Select **Salary Tool → Sync Employee Tabs**.
3. Open `SUM` and review the employee totals.
4. Click an employee name on `SUM` to jump to that employee's summary block.
5. Click `SUM dashboard` in the employee sheet's `M34` cell to return to that employee's dashboard row.

The totals on `SUM` are direct references to the employee sheets. They are not recalculated independently. Duration values use elapsed-time formatting, so totals above 24 hours display correctly.

## Important notes

- Do not manually maintain the generated dashboard area `A1:D43`; Sync rebuilds it.
- Notes, charts, or other content outside `A1:D43` are not owned by Sync.
- If a daily calculation contains an error, the error remains visible in the employee summary and on `SUM`.
- Removed employees' sheets are preserved as historical sheets, but they no longer appear on `SUM`.
- If `SUM` is deleted, running Sync recreates it.
- A duplicate employee name in `CONFIG` stops Sync and must be corrected before the dashboard can update.
