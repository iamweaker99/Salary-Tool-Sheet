/**
 * Salary Tool — Apps Script (Phase 3: Code Migration + Phase 4: ADD MEETING)
 *
 * Menu + Sync Employee Tabs + Protect/Unprotect All Tabs + formula recalc.
 * Phase 4 (ADD MEETING): sync also maintains the MEETING attendance sheet —
 * header columns for CONFIG flag=YES employees in CONFIG order, plus strict
 * ✓/✗/MC/PH dropdowns on every column position (40-slot capacity).
 * Design: MEM_SALARY_TOOL_PHASE_CODE.md (function specs), MEM_SALARY_TOOL_PHASE_LAYOUT.md,
 * MEM_SALARY_TOOL_PHASE_FORMULA.md, MEM_SALARY_TOOL_ADD_MEETING.md.
 * Code idiom: Mappooling_Sheet_Code.js.
 *
 * Core idea: the employee name lives in hidden cell A1 of each employee tab;
 * every INDEX/MATCH formula references $A$1, so sync NEVER rewrites formulas —
 * it only copies _Template, renames the tab, and sets A1.
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const PROPERTIES = PropertiesService.getScriptProperties();

/** Sheets that are never treated as employee tabs (also excluded from protection ops). */
const RESERVED_SHEETS = ["CONFIG", "IN", "OUT", "MEETING", "_Template", "Test_Harness", "Sheet1"];

/** CONFIG layout: A5:A44 = employee names (source of truth), B5:B44 = standard check-in times. */
const CONFIG_NAMES_RANGE = "A5:A44";
const CONFIG_TIMES_RANGE = "B5:B44";
/** CONFIG C5:C44 = "Need Meeting ?" flags — "YES" marks a meeting-required employee. */
const CONFIG_FLAGS_RANGE = "C5:C44";

/**
 * MEETING layout: row 1 = employee headers, rows 2-32 = attendance cells
 * (✓ / ✗ / MC / PH / blank). Capacity = 40 employee slots (B..AO); the Phase 4
 * migration step expands MEETING to 41 columns (A + 40) BEFORE any sync runs,
 * but the rewrite clamps to the live grid size anyway (safe against smaller).
 */
const MEETING_CAPACITY_COLS = 40;

/** IN/OUT data zone: row 1 = employee headers, rows 2-32 = days 1-31. */
const DATA_LAST_ROW = 32;

/** ScriptProperties flag for the CONFIG buttons (see drawButtonsIfNeeded). */
const BUTTONS_DRAWN_FLAG = "CONFIG_BUTTONS_DRAWN";

/**
 * Smoke-test function for `clasp run`.
 */
function hello() {
  return "Hello from clasp! Push successful ✅";
}

/**
 * Menu setup — runs automatically when the spreadsheet opens.
 * Also attempts to draw the CONFIG buttons on first open.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Salary Tool')
    .addItem('Sync Employee Tabs', 'syncEmployeeTabs')
    .addItem('Reset Rest Overrides', 'resetRestOverrides')
    .addSeparator()
    .addItem('Protect All Tabs', 'protectAllTabs')
    .addItem('Unprotect All Tabs', 'unprotectAllTabs')
    .addToUi();
  drawButtonsIfNeeded();
}

/**
 * Draws blue "Sync Employee Tabs" + grey "Protect All Tabs" buttons on CONFIG,
 * first open only. Buttons are visual only — the menu is the reliable path.
 *
 * The Apps Script drawing API can read/reposition/remove drawings
 * (getDrawings, setOnAction, remove) but has no documented creation method, so
 * the creation attempt below is guarded: any failure is caught and a
 * ScriptProperties flag prevents retrying on every open. If getDrawings() itself
 * is unavailable, the same flag decides whether to skip.
 */
function drawButtonsIfNeeded() {
  const config = SS.getSheetByName("CONFIG");
  if (!config) return;
  try {
    if (config.getDrawings().length > 0) return; // already drawn
  } catch (e) {
    // getDrawings unavailable in this runtime — fall back to the flag.
    if (PROPERTIES.getProperty(BUTTONS_DRAWN_FLAG) === "true") return;
  }
  try {
    const syncButton = SpreadsheetApp.newDrawing()
      .setShapeType(SpreadsheetApp.ShapeType.ROUNDED_RECTANGLE)
      .setPosition(2, 3, 0, 0)
      .setSize(180, 32)
      .setFillColor("#1a73e8")
      .setText("Sync Employee Tabs")
      .build();
    config.insertDrawing(syncButton);
    const protectButton = SpreadsheetApp.newDrawing()
      .setShapeType(SpreadsheetApp.ShapeType.ROUNDED_RECTANGLE)
      .setPosition(2, 5, 0, 0)
      .setSize(180, 32)
      .setFillColor("#9aa0a6")
      .setText("Protect All Tabs")
      .build();
    config.insertDrawing(protectButton);
    PROPERTIES.setProperty(BUTTONS_DRAWN_FLAG, "true");
  } catch (e) {
    // Drawing creation not supported — menu covers both actions. Don't retry.
    PROPERTIES.setProperty(BUTTONS_DRAWN_FLAG, "true");
  }
}

/**
 * Core sync. Reads CONFIG A5:C44 (names, check-in times, meeting flags; blank
 * rows skipped), ensures every employee has a tab (copy _Template → rename →
 * set A1) and an IN/OUT column pair, rewrites IN/OUT to CONFIG order, rewrites
 * the MEETING attendance sheet to flag=YES order, and force-recalculates all
 * formulas. Returns a single summary string — no dialogs anywhere in the core
 * path, so it works headless (menu, button, clasp run).
 */
function syncEmployeeTabs() {
  const summary = [];
  const config = SS.getSheetByName("CONFIG");
  if (!config) {
    return "SYNC ABORTED: CONFIG sheet not found.";
  }

  // 1. Read the employee list from CONFIG (skip blank rows — A is the source of truth).
  const namesValues = config.getRange(CONFIG_NAMES_RANGE).getValues().map(r => String(r[0]).trim());
  const timeValues = config.getRange(CONFIG_TIMES_RANGE).getValues().map(r => r[0]);
  const flagValues = config.getRange(CONFIG_FLAGS_RANGE).getValues().map(r => String(r[0]).trim());
  const employees = [];
  for (let i = 0; i < namesValues.length; i++) {
    if (namesValues[i] === "") continue; // blank slot
    employees.push({ name: namesValues[i], time: timeValues[i], flag: flagValues[i] });
  }
  if (employees.length === 0) {
    return "SYNC: CONFIG has no employees — nothing to do.";
  }
  const employeeNames = employees.map(e => e.name);
  summary.push("CONFIG employees: " + employeeNames.length);
  // MEETING attendance: only employees flagged "YES" (anything else — "NO", "-",
  // "" — is not a meeting participant), in CONFIG order.
  const meetingNames = employees.filter(e => e.flag === "YES").map(e => e.name);
  summary.push("MEETING employees (flag=YES): " + meetingNames.length);
  const noTime = employees.filter(e => e.time === "" || e.time === null).map(e => e.name);
  if (noTime.length > 0) {
    summary.push("WARN: standard check-in time missing for: " + noTime.join(", "));
  }

  // 2. Inventory sheets and find orphaned tabs (exist but not in CONFIG, not reserved).
  const allSheets = SS.getSheets();
  const sheetByName = {};
  allSheets.forEach(s => { sheetByName[s.getName()] = s; });
  const orphanedTabs = allSheets.filter(s =>
    !RESERVED_SHEETS.includes(s.getName()) && !employeeNames.includes(s.getName()));
  if (orphanedTabs.length > 0) {
    summary.push("Orphaned tabs (left alone): " + orphanedTabs.map(s => s.getName()).join(", "));
  }

  const template = SS.getSheetByName("_Template");
  const inSheet = SS.getSheetByName("IN");
  const outSheet = SS.getSheetByName("OUT");
  if (!template || !inSheet || !outSheet) {
    return "SYNC ABORTED: _Template / IN / OUT sheet missing.";
  }

  // 3. Rename detection (design scenario 7): exactly one pair to clear, exactly
  //    one new name, and an orphaned tab named after the cleared pair → treat as
  //    a rename: keep the IN/OUT data (under the new header) and rename the tab.
  const inPairs = readPairData(inSheet);
  const renameMap = {}; // newName -> oldName
  const pendingClear = Object.keys(inPairs.pairs).filter(n => !employeeNames.includes(n));
  const newNames = employeeNames.filter(n => !inPairs.pairs[n]);
  if (pendingClear.length === 1 && newNames.length === 1 &&
      sheetByName[pendingClear[0]] && !sheetByName[newNames[0]]) {
    renameMap[newNames[0]] = pendingClear[0];
    summary.push("Detected rename: " + pendingClear[0] + " → " + newNames[0] +
      " (tab renamed, IN/OUT data preserved)");
  }

  // 4. Ensure every employee has a tab.
  let created = 0;
  let updated = 0;
  employeeNames.forEach(name => {
    let sheet = sheetByName[name];
    if (!sheet && renameMap[name]) {
      // Rename: reuse the old (orphaned) tab instead of creating a new one.
      sheet = sheetByName[renameMap[name]];
      sheet.setName(name);
      sheetByName[name] = sheet;
      delete sheetByName[renameMap[name]];
    }
    if (sheet) {
      updated++;
    } else {
      sheet = template.copyTo(SS);
      sheet.setName(name);
      sheetByName[name] = sheet;
      created++;
    }
    setA1(sheet, name, summary); // protection-aware (copies inherit _Template protection)
  });
  summary.push("Employee tabs — created: " + created + ", updated: " + updated);

  // 5. IN/OUT column management: read all pairs into memory, rewrite headers +
  //    data in CONFIG order (append / reorder / delete-clear in one pass), and
  //    apply per-column role properties. Columns are never deleted.
  rewriteInOutLayout(inSheet, employeeNames, renameMap);
  rewriteInOutLayout(outSheet, employeeNames, renameMap);
  summary.push("IN/OUT columns rewritten in CONFIG order (" + employeeNames.length + " pairs)");

  // 6. MEETING attendance layout: single-column rewrite in flag=YES order (same
  //    non-destructive semantics — vacated columns cleared, columns kept; rename
  //    inherits attendance via renameMap), dropdowns applied to every position.
  const meetingSheet = SS.getSheetByName("MEETING");
  if (meetingSheet) {
    rewriteMeetingLayout(meetingSheet, meetingNames, renameMap);
    summary.push("MEETING columns rewritten (" + meetingNames.length + " attendance columns)");
  } else {
    summary.push("WARN: MEETING sheet not found — attendance sync skipped.");
  }

  // 7. Force formula recalc — REQUIRED: formulas that referenced a sheet/column
  //    that didn't exist at write time cache #REF! permanently otherwise.
  const recalcSkipped = forceFormulaRecalc();
  if (recalcSkipped.length > 0) {
    summary.push("WARN: recalc skipped on: " + recalcSkipped.join(", "));
  }

  summary.push("Sync complete.");
  return "SYNC OK\n" + summary.join("\n");
}

/**
 * Clears only current employee rest-control cells in OUT, returning them to
 * blank/Auto. The alternating checkout/control layout is preserved.
 */
function resetRestOverrides() {
  const config = SS.getSheetByName("CONFIG");
  const out = SS.getSheetByName("OUT");
  if (!config || !out) {
    return "RESET ABORTED: CONFIG / OUT sheet missing.";
  }

  const employeeNames = config.getRange(CONFIG_NAMES_RANGE).getValues()
    .map(row => String(row[0]).trim())
    .filter(name => name !== "");
  const pairCapacity = Math.floor((out.getMaxColumns() - 1) / 2);
  if (employeeNames.length > pairCapacity) {
    return "RESET ABORTED: OUT does not have enough employee control columns.";
  }

  employeeNames.forEach((name, index) => {
    const controlColumn = 3 + index * 2; // C, E, ..., alternating after checkout
    out.getRange(2, controlColumn, DATA_LAST_ROW - 1, 1).clearContent();
  });
  return "RESET OK: " + employeeNames.length + " rest-control column(s) cleared.";
}

/**
 * Reads an IN/OUT sheet's column pairs into memory. Each pair = time column
 * (header = employee name) + adjacent checkbox column (blank header).
 * Returns { pairs: {name: {time: [31], cb: [31]}}, order: [...] }.
 */
function readPairData(sheet) {
  const maxCols = sheet.getMaxColumns();
  const values = sheet.getRange(1, 2, DATA_LAST_ROW, maxCols - 1).getValues();
  const pairs = {};
  const order = [];
  for (let i = 0; i < Math.floor((maxCols - 1) / 2); i++) {
    const header = values[0][i * 2];
    if (header !== "" && header !== null && String(header).trim() !== "") {
      const name = String(header).trim();
      pairs[name] = {
        time: values.slice(1).map(r => r[i * 2]),
        cb: values.slice(1).map(r => r[i * 2 + 1])
      };
      order.push(name);
    }
  }
  return { pairs, order };
}

/**
 * Rewrites a sheet's pair layout to CONFIG order (non-destructive):
 * 1. Read all pairs into memory (data preserved, keyed by header name).
 * 2. Overwrite the full B.. grid: headers + data at positions 0..N-1.
 *    Deleted employees fall out of the rewrite → their columns end up blank
 *    (cleared header + data, columns kept — never deleted).
 *    Renamed employees inherit their old pair's data via renameMap.
 * 3. Apply the "Column Role Properties" rule to every assigned pair:
 *    time col = TIME HH:MM format + cleared validation;
 *    rest-control col = strict 0h/1h validation + no TIME format.
 *    (Validation is a cell property that survives value clearing — stale rules
 *    from former checkbox columns would reject time entry, and grid-expansion
 *    the rest-control model remains blank/0h/1h. See design doc.)
 */
function rewriteInOutLayout(sheet, employeeNames, renameMap) {
  const maxCols = sheet.getMaxColumns();
  const pairCount = Math.floor((maxCols - 1) / 2); // full pairs from B (82 cols -> 40)
  const existing = readPairData(sheet);
  const matrix = Array.from({ length: DATA_LAST_ROW }, () => Array(maxCols - 1).fill(""));
  // Existing rest-control values are preserved; new positions remain blank.
  // Legacy invalid values remain visible until the explicit reset action.
  // applied below makes the engine throw "violates the data validation rules"
  // on any non-boolean when pending writes are flushed (observed 2026-08-01 —
  // the old "@" TEXT format coerced boolean writes to "false" strings, and the
  // violation surfaced at the next run's read). FALSE = unchecked = same
  // semantics as empty.
  employeeNames.forEach((name, n) => {
    const t = n * 2;                 // 0-based pair position within the B-relative zone
    const c = n * 2 + 1;
    const oldName = renameMap[name]; // data to inherit on rename
    const pair = existing.pairs[name] || (oldName && existing.pairs[oldName]);
    matrix[0][t] = name;             // time column header
    // matrix[0][c] stays ""          // rest-control column header stays blank
    for (let r = 1; r < DATA_LAST_ROW; r++) {
      matrix[r][t] = pair ? pair.time[r - 1] : "";
      matrix[r][c] = pair ? pair.cb[r - 1] : "";
    }
  });
  sheet.getRange(1, 2, DATA_LAST_ROW, maxCols - 1).setValues(matrix);
  // Column Role Properties on EVERY pair position, not just employee ones: time
  // cols = HH:MM + cleared validation, checkbox cols = strict BOOLEAN only —
  // NO number format (Fix A, 2026-08-01: the old "@" TEXT format coerced boolean
  // writes to "false" strings — Bug 4 root cause).
  // This also kills stale spill validation in vacated/capacity positions
  // (validation is a cell property that survives value clearing).
  for (let n = 0; n < pairCount; n++) {
    applyTimeColumnProperties(sheet, 2 + n * 2);     // 1-based time column
    applyRestOverrideColumnProperties(sheet, 3 + n * 2); // 1-based control column
  }
}

/** Time column (rows 2-32): TIME HH:MM format, NO data validation. */
function applyTimeColumnProperties(sheet, col) {
  const data = sheet.getRange(2, col, DATA_LAST_ROW - 1, 1);
  data.setNumberFormat("HH:MM");
  data.setDataValidation(null); // clearDataValidation doesn't exist — null removes validation
  sheet.getRange(1, col).setDataValidation(null); // header too (spill belt-and-braces)
}

/**
 * Rest-control column (rows 2-32): strict 0h/1h validation; blank means Auto.
 * Deliberately no number format is applied.
 * Sheets coerce the sync's boolean writes to the STRING "false", which then
 * violated the strict BOOLEAN validation when pending writes were flushed at the
 * next run's read ("cell C2 violates..." — Bug 4 root cause, 2026-08-01). With
 * no number format is applied.
 */
function applyRestOverrideColumnProperties(sheet, col) {
  const data = sheet.getRange(2, col, DATA_LAST_ROW - 1, 1);
  const strictRestOverride = SpreadsheetApp.newDataValidation()
    .requireValueInList(["0h", "1h"], true)
    .setAllowInvalid(false)
    .build();
  data.setDataValidation(strictRestOverride);
  sheet.getRange(1, col).setDataValidation(null);
}

/**
 * Reads the MEETING sheet's attendance columns into memory. Each column =
 * header (row 1 = employee name) + 31 attendance values (rows 2-32).
 * Returns { pairs: {name: [31 values]}, order: [...] }.
 */
function readMeetingData(sheet) {
  const maxCols = sheet.getMaxColumns();
  const values = sheet.getRange(1, 2, DATA_LAST_ROW, maxCols - 1).getValues();
  const pairs = {};
  const order = [];
  for (let i = 0; i < maxCols - 1; i++) {
    const header = values[0][i];
    if (header !== "" && header !== null && String(header).trim() !== "") {
      const name = String(header).trim();
      pairs[name] = values.slice(1).map(r => r[i]);
      order.push(name);
    }
  }
  return { pairs, order };
}

/**
 * Rewrites the MEETING attendance layout to flag=YES CONFIG order (non-destructive,
 * single-column variant of rewriteInOutLayout):
 * 1. Read all attendance into memory (data preserved, keyed by header name).
 * 2. Overwrite headers + data at positions 0..N-1 of the B-relative zone.
 *    Flag-flipped employees (YES→NO) fall out of the rewrite → their columns
 *    end up blank (header + data cleared, columns kept — never deleted).
 *    Renamed employees inherit their old column's attendance via renameMap.
 * 3. Apply the attendance dropdown to EVERY column position — employee and
 *    future-capacity alike (validation is a cell property that survives value
 *    clearing — Phase 3 Bug 4 lesson), so future employees / grid expansions
 *    inherit it.
 *
 * CLAMP: only columns B..min(41, maxCols) are ever written. The migration step
 * expands MEETING to 41 columns before sync runs, but the clamp keeps this safe
 * against a smaller grid (e.g. today's 26-column MEETING). Columns beyond
 * capacity are never touched. Capacity is 40 = CONFIG's 40 slots, so meetingNames
 * always fits; a larger list would fail loudly in setValues, never silently.
 */
function rewriteMeetingLayout(sheet, meetingNames, renameMap) {
  const colCount = Math.min(MEETING_CAPACITY_COLS, sheet.getMaxColumns() - 1); // B-relative positions
  if (colCount < 1) return; // degenerate grid (no column B) — nothing to maintain
  const existing = readMeetingData(sheet);
  const matrix = Array.from({ length: DATA_LAST_ROW }, () => Array(colCount).fill(""));
  meetingNames.forEach((name, n) => {
    const oldName = renameMap[name]; // attendance to inherit on rename
    const data = existing.pairs[name] || (oldName && existing.pairs[oldName]);
    matrix[0][n] = name;             // attendance column header
    if (data) {
      for (let r = 1; r < DATA_LAST_ROW; r++) matrix[r][n] = data[r - 1];
    }
  });
  sheet.getRange(1, 2, DATA_LAST_ROW, colCount).setValues(matrix);
  for (let n = 0; n < colCount; n++) {
    applyMeetingAttendanceValidation(sheet, 2 + n); // 1-based column
  }
}

/**
 * Attendance column (rows 2-32): strict ONE_OF_LIST dropdown [✓, ✗, MC, PH].
 * Blank cells are permitted (no requireValue) — the neutral "not attended"
 * state. NOTE: deliberately NO number format, same as the checkbox columns
 * (Phase 3 Bug 4 / Fix A: formats coerce writes to strings that violate strict
 * validation).
 */
function applyMeetingAttendanceValidation(sheet, col) {
  const data = sheet.getRange(2, col, DATA_LAST_ROW - 1, 1);
  const attendance = SpreadsheetApp.newDataValidation()
    .requireValueInList(['✓', '✗', 'MC', 'PH'], true)
    .setAllowInvalid(false)
    .build();
  data.setDataValidation(attendance);
  sheet.getRange(1, col).setDataValidation(null); // header too (spill belt-and-braces)
}

/**
 * Writes the employee name into hidden A1. If the cell is protected
 * (non-warning-only), unprotect → write → re-protect, preserving each
 * protection's configuration (range, description, sheet exceptions, editors).
 */
function setA1(sheet, name, summary) {
  try {
    sheet.getRange("A1").setValue(name);
  } catch (e) {
    try {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
        .concat(sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET))
        .filter(p => !p.isWarningOnly());
      const configs = protections.map(captureProtection);
      protections.forEach(p => p.remove());
      sheet.getRange("A1").setValue(name);
      configs.forEach(cfg => restoreProtection(sheet, cfg));
    } catch (e2) {
      summary.push("WARN: could not write A1 on '" + sheet.getName() + "': " + e2.message);
    }
  }
}

/** Snapshot of a protection's config so it can be recreated after a write. */
function captureProtection(p) {
  return {
    type: p.getType(),
    description: p.getDescription(),
    a1Notation: p.getType() === SpreadsheetApp.ProtectionType.RANGE
      ? p.getRange().getA1Notation()
      : null,
    unprotectedRanges: p.getType() === SpreadsheetApp.ProtectionType.SHEET
      ? p.getUnprotectedRanges().map(r => r.getA1Notation())
      : [],
    editors: p.getEditors().map(u => u.getEmail()).filter(Boolean)
  };
}

/** Recreates a protection captured by captureProtection on the same sheet. */
function restoreProtection(sheet, cfg) {
  let p;
  if (cfg.type === SpreadsheetApp.ProtectionType.RANGE) {
    p = sheet.getRange(cfg.a1Notation).protect();
  } else {
    p = sheet.protect();
    if (cfg.unprotectedRanges.length > 0) {
      p.setUnprotectedRanges(cfg.unprotectedRanges.map(a1 => sheet.getRange(a1)));
    }
  }
  if (cfg.description) p.setDescription(cfg.description);
  cfg.editors.forEach(email => p.addEditor(email));
}

/**
 * Locks A1 (employee name) + C2:D32 (check-in/check-out formulas) on every
 * non-reserved tab; all other cells stay editable. Sheet protection with
 * description "Employee Tab". Reserved sheets (CONFIG, IN, OUT, MEETING,
 * _Template, Test_Harness, Sheet1) are never touched.
 */
function protectAllTabs() {
  const summary = [];
  let protectedCount = 0;
  SS.getSheets()
    .filter(s => !RESERVED_SHEETS.includes(s.getName()))
    .forEach(sheet => {
      // Start from a clean slate (also clears any protection inherited from _Template copies).
      sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
      sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
      // Unlock everything except A1 and C2:D32.
      const maxRows = sheet.getMaxRows();
      const maxCols = sheet.getMaxColumns();
      const unlocked = [];
      const dataRows = Math.min(DATA_LAST_ROW, maxRows);
      if (dataRows >= 2) {
        if (maxCols >= 2) unlocked.push(sheet.getRange(2, 2, dataRows - 1, 1));            // B2:B32 dates
        if (maxCols >= 5) unlocked.push(sheet.getRange(2, 5, dataRows - 1, maxCols - 4));  // E2:last formulas
      }
      if (maxRows > DATA_LAST_ROW) {
        unlocked.push(sheet.getRange(DATA_LAST_ROW + 1, 1, maxRows - DATA_LAST_ROW, maxCols));
      }
      const p = sheet.protect();
      p.setDescription("Employee Tab");
      p.setUnprotectedRanges(unlocked);
      protectedCount++;
    });
  summary.push("Protected: " + protectedCount + " tab(s).");
  return "PROTECT OK\n" + summary.join("\n");
}

/**
 * Removes all protections (range + sheet) from every non-reserved tab.
 * Reserved sheets are never touched.
 */
function unprotectAllTabs() {
  let unprotectedCount = 0;
  SS.getSheets()
    .filter(s => !RESERVED_SHEETS.includes(s.getName()))
    .forEach(sheet => {
      sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
      sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
      unprotectedCount++;
    });
  return "UNPROTECT OK: " + unprotectedCount + " tab(s) unprotected.";
}

/**
 * Re-applies every formula in place on all non-reserved sheets — forces Google
 * Sheets to re-evaluate references that were cached as #REF! when the target
 * sheet or IN/OUT column didn't exist at formula-write time. REQUIRED after
 * every sync. Returns names of sheets where recalc was skipped (if any).
 */
function forceFormulaRecalc() {
  const skipped = [];
  SS.getSheets()
    .filter(sheet => !RESERVED_SHEETS.includes(sheet.getName()))
    .forEach(sheet => {
      try {
        const maxRows = sheet.getMaxRows();
        const maxCols = sheet.getMaxColumns();
        const formulas = sheet.getRange(1, 1, maxRows, maxCols).getFormulas();
        // Bounding box of the cells that actually contain formulas.
        // IMPORTANT: full-grid setFormulas(getFormulas()) would CLEAR non-formula
        // cells (getFormulas() returns "" for them) — observed 2026-08-01: it wiped
        // A1 employee names; it would also wipe user-entered B-column dates.
        let minR = maxRows, maxR = 0, minC = maxCols, maxC = 0;
        for (let r = 0; r < maxRows; r++) {
          for (let c = 0; c < maxCols; c++) {
            if (formulas[r][c] !== "") {
              if (r < minR) minR = r;
              if (r > maxR) maxR = r;
              if (c < minC) minC = c;
              if (c > maxC) maxC = c;
            }
          }
        }
        if (maxR >= minR) {
          const box = sheet.getRange(minR + 1, minC + 1, maxR - minR + 1, maxC - minC + 1);
          box.setFormulas(box.getFormulas());
        }
      } catch (e) {
        skipped.push(sheet.getName());
      }
    });
  return skipped;
}
