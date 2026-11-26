// Google Apps Script — paste in Extensions > Apps Script of your Google Sheet
const SHEET_NAME = "Leads";

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const productNames = Object.keys(data.products);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Shop Name", "District", "City", "Contact Number"]
      .concat(productNames)
      .concat(["Bio Bag Note"]));
    sheet.setFrozenRows(1);
  }

  const row = [new Date(), data.shopName, data.district, data.city, "'" + data.contact]
    .concat(productNames.map(n => data.products[n]))
    .concat([data.bioBagNote]);
  sheet.appendRow(row);

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
