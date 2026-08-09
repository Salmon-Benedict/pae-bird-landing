// Deploy this bound to a Google Sheet with a tab named "Codes":
//   A1 = "Code" (header)
//   A2:A101 = up to 100 promo codes, one per row
//
// Deploy > New deployment > Web app > Execute as "Me" > Who has access "Anyone" > Deploy.
// Paste the resulting /exec URL into PROMO_CODE_ENDPOINT in index.html.

function doGet(e) {
  if (e.parameter.peek === '1') {
    return handlePeek();
  }
  return handleClaim();
}

function handlePeek() {
  var sheet = SpreadsheetApp.openById('1yQWJcAmKPw5hAz_o-vWMWLHlfF4Yk6Xx8XmZ9amfRtI').getSheetByName('Codes');
  var lastRow = sheet.getLastRow();
  var remaining = Math.max(0, lastRow - 1);
  return jsonOutput({ remaining: remaining });
}

function handleClaim() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.openById('1yQWJcAmKPw5hAz_o-vWMWLHlfF4Yk6Xx8XmZ9amfRtI').getSheetByName('Codes');
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return jsonOutput({ code: null, remaining: 0 });
    }
    var code = String(sheet.getRange(2, 1).getValue()).trim();
    sheet.deleteRow(2);
    var remaining = sheet.getLastRow() - 1;
    return jsonOutput({ code: code, remaining: remaining });
  } finally {
    lock.releaseLock();
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
