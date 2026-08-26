# Rest Duration Feature Guide

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
4. At the start of a new calculation period, select `Salary Tool → Reset Rest Overrides`.
   This clears only rest overrides; dates, checkouts, and other data remain unchanged.
5. Changing Month or Year does not reset overrides automatically.

Invalid pasted values produce a visible calculation error and should be corrected.
