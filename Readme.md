/**
 * BATTLEAX 3.0 — Round Suggestion Form backend
 *
 * SETUP INSTRUCTIONS:
 * 1. Create a new Google Sheet. Name it something like "BATTLEAX 3.0 Suggestions".
 * 2. In the sheet, go to Extensions > Apps Script.
 * 3. Delete any starter code in Code.gs and paste this entire file in.
 * 4. Click the disk icon to save the project (give it any name).
 * 5. Click "Deploy" > "New deployment".
 *    - Click the gear icon next to "Select type" and choose "Web app".
 *    - Description: anything, e.g. "BATTLEAX form handler"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 *    - Click "Deploy".
 * 6. Authorize the script when prompted (click through the "unsafe" warning —
 *    this is your own script, it's safe).
 * 7. Copy the "Web app URL" you're given. It looks like:
 *    https://script.google.com/macros/s/AKfycb.../exec
 * 8. Paste that URL into index.html, replacing PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE
 * 9. Run setupSheet() once manually from the Apps Script editor (select it
 *    from the function dropdown at the top, then click Run) to create headers.
 *
 * If you ever update this script after it's deployed, you must create a
 * "New deployment" again (or use "Manage deployments" > edit > new version)
 * for changes to take effect on the live URL.
 */

const SHEET_NAME = "Suggestions";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    sheet.appendRow([
      new Date(),
      data.fullName || "",
      data.classSem || "",
      data.eventRound || "",
      data.roundName || "",
      data.suggestionType || "",
      data.ideaDetail || "",
      data.conductMethod || "",
      data.skills || "",
      data.teamType || "",
      data.teamSize || "",
      data.fitReason || "",
      data.comments || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "BATTLEAX 3.0 form backend is live" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    writeHeaders(sheet);
  }
  return sheet;
}

function writeHeaders(sheet) {
  const headers = [
    "Timestamp",
    "Full Name",
    "Class / Semester",
    "Event or Round",
    "Suggested Round Name",
    "Type of Suggestion",
    "Idea Detail",
    "How Round Should Be Conducted",
    "Skills Tested",
    "Individual / Team",
    "Team Size",
    "Why It Fits BATTLEAX 3.0",
    "Additional Comments"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * Run this once manually (select from function dropdown, click Run)
 * to set up the sheet with proper headers before going live.
 */
function setupSheet() {
  const sheet = getOrCreateSheet();
  if (sheet.getRange(1, 1).getValue() !== "Timestamp") {
    writeHeaders(sheet);
  }
  Logger.log("Sheet ready: " + sheet.getName());
}
