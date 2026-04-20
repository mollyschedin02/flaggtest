/**
 * Paste this into your Apps Script project (bound to nothing, or any project).
 * REQUIRED: Set SPREADSHEET_ID to your Google Sheet’s ID from the URL:
 * https://docs.google.com/spreadsheets/d/THIS_PART_IS_THE_ID/edit
 *
 * Deploy → New deployment → Web app
 * - Execute as: Me
 * - Who has access: Anyone
 * Then paste the Web App URL into index-grid.html as LOG_ENDPOINT_URL.
 */

const SPREADSHEET_ID = "PASTE_YOUR_SPREADSHEET_ID_HERE";

function doPost(e) {
  try {
    if (!SPREADSHEET_ID || SPREADSHEET_ID === "PASTE_YOUR_SPREADSHEET_ID_HERE") {
      return jsonResponse({ ok: false, error: "Set SPREADSHEET_ID in the script" }, 500);
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Votes");
    if (!sheet) {
      sheet = ss.insertSheet("Votes");
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "timestamp",
        "questionNumber",
        "questionText",
        "selectedImage",
      ]);
    }

    const raw = e && e.postData && e.postData.contents;
    if (!raw) {
      return jsonResponse({ ok: false, error: "Missing POST body" }, 400);
    }

    const payload = JSON.parse(raw);
    const timestamp = payload.timestamp || new Date().toISOString();
    const answers = Array.isArray(payload.answers) ? payload.answers : [];

    if (answers.length === 0) {
      return jsonResponse({ ok: false, error: "No answers provided" }, 400);
    }

    const rows = answers.map(function (a) {
      return [
        timestamp,
        a.questionNumber || "",
        a.questionText || "",
        a.selectedImage || "",
      ];
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    return jsonResponse({ ok: true, rowsWritten: rows.length }, 200);
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) }, 500);
  }
}

function doGet() {
  return jsonResponse({ ok: true, message: "Use POST to submit votes." }, 200);
}

function jsonResponse(obj, statusCode) {
  obj.status = statusCode;
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
