// Google Apps Script — deploy as Web app (Execute as: Me, Access: Anyone)

// Google Sheet-oda peyar (exact-ah irukkanum). ID type panna thevai illa.
const SHEET_FILE_NAME = "Get your offer price";
const SHEET_NAME = "Leads";

function getSpreadsheet() {
  const files = DriveApp.getFilesByName(SHEET_FILE_NAME);
  if (!files.hasNext()) throw new Error("Sheet not found: " + SHEET_FILE_NAME);
  return SpreadsheetApp.open(files.next());
}

// Browser-la web app URL open pannina "OK" nu theriyanum (deployment test)
function doGet() {
  return ContentService.createTextOutput("OK - web app is working");
}

// Editor-la idhai ONE TIME "Run" pannunga: permission kekkum (Allow pannanum),
// apparam sheet-la test row varum. Error irundha Execution log-la theriyum.
function testWrite() {
  const fake = {
    postData: { contents: JSON.stringify({
      shopName: "Test Shop", district: "Test", city: "Test", contact: "9999999999",
      products: { "Test Product": "Type: Plain | Qty: 1" }, bioBagNote: ""
    }) }
  };
  doPost(fake);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const ss = getSpreadsheet();
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
