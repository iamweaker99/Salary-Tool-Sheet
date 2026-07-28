/**
 * Test function — writes a timestamp to the Sheet to verify clasp push works.
 */
function addTimestamp() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const now = new Date();
  sheet.getRange("A1").setValue("Last clasp push test: " + now.toISOString());
}

/**
 * Simple hello function to test clasp run.
 */
function hello() {
  return "Hello from clasp! Push successful ✅";
}
