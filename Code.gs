const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('The Quad Connect')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return {
    events:      getSheetData(ss, 'Events'),
    showAndTell: getSheetData(ss, 'ShowAndTell'),
    wins:        getSheetData(ss, 'Wins'),
    directory:   getSheetData(ss, 'Directory')
  };
}

function getSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1)
    .filter(r => r.some(c => c !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i];
      });
      return obj;
    });
}
