const SPREADSHEET_ID = "1eqGyFo2rWBrYMVOsJmG37HKZo1_wgohgdaumVsvsH5U";
const SHEET_NAME = "Register"; 
const FUNADHOO_SHEET_NAME = "Funadhoo"; 

function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'getData') {
    return ContentService.createTextOutput(JSON.stringify(getRecordsData()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'getFunadhooData') {
    return ContentService.createTextOutput(JSON.stringify(getFunadhooRecordsData()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "Online", message: "TVC API Active" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    if (action === "create") {
      const res = createRecord(data);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    } else if (action === "update") {
      const res = updateRecord(data);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    } else if (action === "delete") {
      const res = deleteRecord(data.rowIndex);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    }
    
    throw new Error("Invalid action specified.");
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getRecordsData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];

  const headers = data[0].map(h => h.toString().trim());
  const rows = data.slice(1);
  
  return rows.map((row, index) => {
    let obj = { rowIndex: index + 2 };
    headers.forEach((header, colIndex) => {
      let val = row[colIndex];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      obj[header] = val;
    });
    return obj;
  });
}

function getFunadhooRecordsData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(FUNADHOO_SHEET_NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => h.toString().trim());
  const rows = data.slice(1);
  
  return rows.map((row, index) => {
    let obj = { rowIndex: index + 2 };
    headers.forEach((header, colIndex) => {
      let val = row[colIndex];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      obj[header] = val;
    });
    return obj;
  });
}

function createRecord(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  
  let newRow = headers.map(header => data[header] !== undefined ? data[header] : "");
  sheet.appendRow(newRow);
  
  const newRowIndex = sheet.getLastRow();
  data.rowIndex = newRowIndex;
  
  const pdfResult = generateQuotationPDF(data);
  return { success: true, message: "Record created successfully", pdf: pdfResult };
}

function updateRecord(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  
  const rowIndex = parseInt(data.rowIndex);
  if (!rowIndex || rowIndex < 2) throw new Error("Invalid row index for update.");
  
  headers.forEach((header, colIndex) => {
    if (data[header] !== undefined) {
      sheet.getRange(rowIndex, colIndex + 1).setValue(data[header]);
    }
  });
  
  const pdfResult = generateQuotationPDF(data);
  return { success: true, message: "Record updated successfully", pdf: pdfResult };
}

function deleteRecord(rowIndex) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  sheet.deleteRow(parseInt(rowIndex));
  return { success: true, message: "Record deleted" };
}

function generateQuotationPDF(formData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("TVC_Quotation") || ss.getSheetByName(SHEET_NAME);

  if (!sheet) throw new Error("Target sheet for PDF export not found.");

  try {
    sheet.getRange("A1:H1").setFontFamily("Faruma");
    sheet.getRange("A4:H4").setFontFamily("Faruma");

    ["A15:H15", "A18:H18", "A21:H21"].forEach(rng => {
      sheet.getRange(rng)
        .setFontFamily("Faruma")
        .setWrap(true)
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center");
    });

    sheet.setRowHeight(15, 40);
    sheet.setRowHeight(18, 40);

    sheet.getRange("A13:H13")
      .setFontFamily("Arial")
      .setFontSize(14)
      .setWrap(true)
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center");

    sheet.getRange("A19:H19")
      .setFontFamily("Arial")
      .setFontSize(12)
      .setWrap(true)
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center");

    sheet.setRowHeight(13, 70);
    sheet.setRowHeight(19, 40);
  } catch (e) {}

  const range = sheet.getRange("A18:H19");
  const values = range.getValues();
  const isNotBlank = values.some(row => row.some(cell => cell !== ""));
  let borderApplied = false;

  if (isNotBlank) {
    range.setBorder(true, true, true, true, true, true);
    borderApplied = true;
  }

  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
    'format=pdf' +
    '&size=A5' +
    '&portrait=false' +
    '&fitw=true' +
    '&top_margin=0' +
    '&bottom_margin=0' +
    '&left_margin=0.5' +
    '&right_margin=0.5' +
    '&gridlines=false' +
    '&printtitle=false' +
    '&sheetnames=false' +
    '&pagenumbers=false' +
    '&horizontal_alignment=CENTER' +
    '&vertical_alignment=MIDDLE' +
    '&range=A1:H22' +
    `&gid=${sheet.getSheetId()}` +
    '&scale=2';

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Failed to export PDF code: ${response.getResponseCode()}`);
  }

  const nameVal = formData.beneficiary_name || 'Document';
  const pdfBlob = response.getBlob().setName(`TVC_Quotation_${nameVal}.pdf`);
  const file = DriveApp.createFile(pdfBlob);

  if (borderApplied) {
    range.setBorder(false, false, false, false, false, false);
  }

  return {
    success: true,
    url: file.getUrl(),
    name: file.getName()
  };
}
