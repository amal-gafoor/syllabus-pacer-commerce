/**
 * Syllabus Pacer — Google Sheets bridge
 * ------------------------------------------------------------------
 * Appends each generated plan (identified by EPSID) as a new row.
 *
 * SETUP (one time, ~3 min):
 *   1. Open the submissions Google Sheet (or any Apps Script project).
 *   2. Extensions → Apps Script. Delete any boilerplate, paste this file.
 *      Submissions always go to the sheet in SPREADSHEET_ID below,
 *      into a "Submissions" tab (created automatically on first save).
 *   3. Deploy → New deployment → gear icon → Web app.
 *        - Execute as:      Me   (must have edit access to that sheet)
 *        - Who has access:  Anyone
 *      Click Deploy, authorise, and COPY the "/exec" web-app URL.
 *   4. Put that URL in the Next.js app's .env.local as SHEETS_WEBAPP_URL.
 * ------------------------------------------------------------------
 */

// The spreadsheet submissions are written to (kept separate from the app).
const SPREADSHEET_ID = "16-bnOkYr9PSRoxX0Ds66yYXQUvtwmDyt1Fdu5zvgvOA";
const SUBMISSIONS_SHEET = "Submissions"; // tab within that spreadsheet

// NOTE: the two topic columns are appended at the END on purpose. Sheets that
// already hold rows written by the older, pages-only version keep their existing
// columns lined up; only the new fields land to the right of "Chapter list".
//
// "Total pages" and "Pages/day" are retained as placeholders and written blank.
// The commerce syllabus has no page counts, so the planner schedules by topic
// and "Total topics" / "Topics/day" carry the real figures. The two dead columns
// stay put so historical science rows keep their headings.
const SUBMISSION_HEADERS = [
  "Timestamp", "EPSID", "Name", "Phone", "Class", "Medium",
  "Start", "Finish by", "Weekly offs", "Extra leave",
  "Chapters", "Total pages", "Pages/day", "Projected finish", "Chapter list",
  "Total topics", "Topics/day",
];

/** POST → append one plan submission as a new row. */
function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
    if (!sheet) sheet = ss.insertSheet(SUBMISSIONS_SHEET);
    if (sheet.getLastRow() === 0) sheet.appendRow(SUBMISSION_HEADERS);
    else widenHeaderIfNeeded(sheet);

    sheet.appendRow([
      new Date(),
      data.epsid || "",
      data.name || "",
      data.phone || "",
      data.class || "",
      data.medium || "",
      data.fromDate || "",
      data.toDate || "",
      (data.weeklyOff || []).join(", "),
      data.extraLeave || 0,
      data.totalChapters || 0,
      "", // Total pages — retired, see SUBMISSION_HEADERS
      "", // Pages/day   — retired, see SUBMISSION_HEADERS
      data.finish || "",
      (data.chapters || []).join("; "),
      data.totalTopics || 0,
      data.topicsPerDay || "",
    ]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/**
 * A sheet created by the earlier pages-only version has a 15-column header and
 * no "Total topics" / "Topics/day". Fill those in once so the new values that
 * appendRow writes to columns 16-17 aren't sitting under blank headings.
 */
function widenHeaderIfNeeded(sheet) {
  const width = sheet.getLastColumn();
  if (width >= SUBMISSION_HEADERS.length) return;
  const missing = SUBMISSION_HEADERS.slice(width);
  sheet.getRange(1, width + 1, 1, missing.length).setValues([missing]);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
