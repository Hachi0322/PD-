const SHEET_NAME = "Order Requests";
const ORDER_HEADERS = ["Timestamp", "Product ID", "Product Name", "Full Name", "Phone Number"];
const PREVIOUS_ORDER_HEADERS = ["Timestamp", "Full Name", "Phone Number"];

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Missing order request data.");
    }

    const parameters = e.parameter || {};
    const data = Object.keys(parameters).length
      ? parameters
      : JSON.parse(e.postData.contents);
    const requiredFields = ["productId", "productName", "fullName", "mobile"];

    requiredFields.forEach(function (field) {
      if (typeof data[field] !== "string" || !data[field].trim()) {
        throw new Error("Missing required field: " + field);
      }
    });

    if (!/^(09\d{9}|\+639\d{9})$/.test(data.mobile.trim())) {
      throw new Error("Invalid Philippine mobile number.");
    }

    const sheet = getOrderSheet();

    sheet.appendRow([
      new Date(),
      safeCell(data.productId),
      safeCell(data.productName),
      safeCell(data.fullName),
      safeCell(data.mobile)
    ]);

    return jsonResponse({ success: true });
  } catch (error) {
    console.error("Order request was not saved:", error);
    return jsonResponse({ success: false, error: error.message });
  }
}

function setupOrderSheet() {
  getOrderSheet();
}

function getOrderSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("Bind this script to the Google Sheet that should receive orders.");
  }

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getLastRow()
    ? sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    : [];
  const hasCurrentHeaders = currentHeaders.length === ORDER_HEADERS.length &&
    ORDER_HEADERS.every(function (header, index) {
      return currentHeaders[index] === header;
    });
  let migratedRows = [];

  if (!hasCurrentHeaders) {
    const hasPreviousHeaders = currentHeaders.length === PREVIOUS_ORDER_HEADERS.length &&
      PREVIOUS_ORDER_HEADERS.every(function (header, index) {
        return currentHeaders[index] === header;
      });
    const previousRows = hasPreviousHeaders && sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, PREVIOUS_ORDER_HEADERS.length).getValues()
      : [];
    migratedRows = previousRows.map(function (row) {
      return [row[0], "", "", row[1], row[2]];
    });

    sheet.clearContents();
  }

  const maxColumns = sheet.getMaxColumns();
  if (maxColumns < ORDER_HEADERS.length) {
    sheet.insertColumnsAfter(maxColumns, ORDER_HEADERS.length - maxColumns);
  } else if (maxColumns > ORDER_HEADERS.length) {
    sheet.deleteColumns(ORDER_HEADERS.length + 1, maxColumns - ORDER_HEADERS.length);
  }

  sheet.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]);
  if (migratedRows.length) {
    sheet.getRange(2, 1, migratedRows.length, ORDER_HEADERS.length).setValues(migratedRows);
  }
  sheet.setFrozenRows(1);
  return sheet;
}

function doGet() {
  return ContentService.createTextOutput("Defense Parasite Plus order form endpoint is ready.");
}

function safeCell(value) {
  const text = String(value).trim();
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
