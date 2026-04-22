function mainFormBuilder() {
  var html = HtmlService.createHtmlOutputFromFile('formBuilderUI') // ref: formBuilderUI.html
    .setTitle(' ')
    .setWidth(300)
  SpreadsheetApp.getUi().showSidebar(html);
}

function classAlreadyExists(classCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const normalized = (classCode || '').toString().trim().toLowerCase();
  if (!normalized) return false;

  // Check from Form Logger first (source of created forms)
  const formLogger = ss.getSheetByName('Form Logger');
  if (formLogger && formLogger.getLastRow() >= 2) {
    const values = formLogger.getRange(2, 3, formLogger.getLastRow() - 1, 1).getValues();
    const existsInLogger = values.some(row => (row[0] || '').toString().trim().toLowerCase() === normalized);
    if (existsInLogger) return true;
  }

  // Check from Class List as a secondary source
  const classList = ss.getSheetByName('Class List');
  if (classList && classList.getLastRow() >= 2) {
    const values = classList.getRange(2, 1, classList.getLastRow() - 1, 1).getValues();
    const existsInClassList = values.some(row => (row[0] || '').toString().trim().toLowerCase() === normalized);
    if (existsInClassList) return true;
  }

  return false;
}

// build the form from 'form_' sheet
function buildForm(subject, classCode, deadline, notes) {
  const normalizedClassCode = (classCode || '').toString().trim();
  if (!normalizedClassCode) {
    throw new Error('Class không được để trống. Vui lòng nhập Class hợp lệ.');
  }

  if (classAlreadyExists(normalizedClassCode)) {
    throw new Error(`Class "${normalizedClassCode}" đã tồn tại. Vui lòng dùng mã Class khác để tránh trùng dữ liệu ở Create Folder và Rules.`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("form_");
  const data = sheet.getDataRange().getValues();
  data.shift(); // remove header row

  // --- Create Form ---
  const formTitle = subject + " - " + normalizedClassCode;
  const form = FormApp.create(formTitle)
    .setDescription(notes + "\nDeadline: " + deadline);

  // --- Add questions ---
  data.forEach(row => {
    const question = row[0];
    const type = row[1];

    let item;
    switch (type.toLowerCase()) {
      case "text":
        item = form.addTextItem().setTitle(question);
        break;
      case "paragraph":
        item = form.addParagraphTextItem().setTitle(question);
        break;
      case "multiple choice":
        item = form.addMultipleChoiceItem()
          .setTitle(question)
          .setChoices([
            form.addMultipleChoiceItem().createChoice("Option 1"),
            form.addMultipleChoiceItem().createChoice("Option 2"),
          ]);
        break;
      default:
        item = form.addTextItem().setTitle(question);
        break;
    }
    item.setRequired(true);
  });

  const editUrl = form.getEditUrl();
  const publishUrl = form.getPublishedUrl();
  const formId = form.getId();

  // --- Create response sheet (R) ---
  const responseSs = SpreadsheetApp.create("Responses - " + formTitle);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, responseSs.getId());

  // --- Optional Drive organization ---
  // If DriveApp is not authorized for current user, keep form creation successful.
  let folderUrl = "";
  try {
    // --- Get parent folder of the app spreadsheet ---
    const appFile = DriveApp.getFileById(ss.getId());
    const parents = appFile.getParents();
    const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

    // --- Find or create "temp" folder ---
    const folders = parentFolder.getFoldersByName("temp");
    const tempFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder("temp");
    const formtempFolder = getOrCreateFolder(tempFolder, '_form_temp_');

    // --- Create a subfolder for this form ---
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
    const formFolder = formtempFolder.createFolder(formId + "_" + timestamp);

    // --- Move form into the subfolder ---
    const formFile = DriveApp.getFileById(formId);
    formFile.moveTo(formFolder);

    // Move response sheet into the same subfolder
    const responseFile = DriveApp.getFileById(responseSs.getId());
    responseFile.moveTo(formFolder);

    folderUrl = formFolder.getUrl();
  } catch (e) {
    Logger.log("Drive organization skipped for buildForm: " + e.message);
  }

  // --- Write log ---
  let logSheet = ss.getSheetByName("Form Logger");
  if (!logSheet) {
    logSheet = ss.insertSheet("Form Logger");
    logSheet.appendRow(["Timestamp", "Subject", "Class", "Publish URL", "Edit URL", "Folder URL", "Response Sheet"]);
    logSheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#d9ead3");
  }
  logSheet.appendRow([new Date(), subject, normalizedClassCode, publishUrl, editUrl, folderUrl, responseSs.getUrl()]);

  return publishUrl; // return form link
}

//////////////
// syncData
///////////////
function manualSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Ensure classList sheet exists ---
  let targetSheet = ss.getSheetByName("Class List");
  if (!targetSheet) {
    targetSheet = ss.insertSheet("Class List");
  } else {
    targetSheet.clear(); // optional: clear previous content
  }

  // --- Read response sheets from Form Logger (no DriveApp dependency) ---
  const logSheet = ss.getSheetByName("Form Logger");
  if (!logSheet || logSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("Không tìm thấy dữ liệu trong Form Logger để sync.");
    return;
  }

  const logData = logSheet.getDataRange().getValues();
  const seenSheetIds = new Set();
  const responseSheets = [];

  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];
    const classCode = (row[2] || "").toString().trim();
    const responseUrl = (row[6] || "").toString().trim();
    if (!responseUrl) continue;

    const match = responseUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) continue;

    const sheetId = match[1];
    if (seenSheetIds.has(sheetId)) continue;
    seenSheetIds.add(sheetId);

    responseSheets.push({
      classCode: classCode,
      sheetId: sheetId
    });
  }

  if (responseSheets.length === 0) {
    SpreadsheetApp.getUi().alert("Không tìm thấy response sheet hợp lệ trong Form Logger.");
    return;
  }

  let headersSet = false;
  let maxCols = 1;

  responseSheets.forEach(item => {
    try {
      const responseSs = SpreadsheetApp.openById(item.sheetId);
      const responseSheet = responseSs.getSheets()[0];
      const data = responseSheet.getDataRange().getValues();

      if (data.length <= 1) return;

      if (!headersSet) {
        const formHeaders = data[0].slice(1);
        targetSheet.appendRow(["Class Name", ...formHeaders]);
        maxCols = Math.max(maxCols, formHeaders.length + 1);
        headersSet = true;
      }

      for (let i = 1; i < data.length; i++) {
        const row = data[i].slice(1);
        const className = item.classCode || responseSs.getName().replace("Responses -", "").trim().split(" - ").pop();
        targetSheet.appendRow([className, ...row]);
        maxCols = Math.max(maxCols, row.length + 1);
      }
    } catch (e) {
      Logger.log(`Skip response sheet ${item.sheetId}: ${e.message}`);
    }
  });

  if (targetSheet.getLastRow() >= 1) {
    targetSheet.getRange(1, 1, 1, maxCols).setFontWeight("bold").setBackground("#d9ead3");
  }

  SpreadsheetApp.getUi().alert("Sync complete!");
}




/////////////////// downloalClassList
///////////////////
function downloalClassList() {
  var html = HtmlService.createHtmlOutputFromFile('downloadClassListUI')
    .setTitle('Download Class List')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Đọc Form Logger sheet và trả về danh sách lớp cho sidebar
 */
function getClassListFromFormLogger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Form Logger");

  if (!logSheet || logSheet.getLastRow() < 2) {
    return [];
  }

  const data = logSheet.getDataRange().getValues();
  const result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var responseUrl = (row[6] || "").toString().trim();

    // Trích xuất spreadsheet ID từ URL
    var match = responseUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) continue;

    result.push({
      subject: (row[1] || "").toString().trim(),
      classCode: (row[2] || "").toString().trim(),
      responseSheetId: match[1],
      timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0])
    });
  }

  Logger.log(result)

  return result;
}

/**
 * Tải file response dưới dạng CSV hoặc Excel
 * Trả về { filename, data (base64) } hoặc { error }
 */
function downloadClassListFile(responseSheetId, subject, classCode, format) {
  try {
    var ss = SpreadsheetApp.openById(responseSheetId);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();

    if (data.length < 1) {
      return { error: "Response sheet không có dữ liệu." };
    }

    var safeName = (subject + "_" + classCode).replace(/[^a-zA-Z0-9_\-\u00C0-\u024F\u1E00-\u1EFF]/g, "_");

    if (format === "csv") {
      var csvContent = convertToCSV(data);
      // Thêm BOM để Excel mở đúng UTF-8
      var bom = "\uFEFF";
      var csvBytes = Utilities.newBlob(bom + csvContent, "text/csv", safeName + ".csv").getBytes();
      return {
        filename: safeName + ".csv",
        data: Utilities.base64Encode(csvBytes)
      };
    } else {
      // Excel: export qua Drive export URL
      var url = "https://docs.google.com/spreadsheets/d/" + responseSheetId + "/export?format=xlsx";
      var blob = UrlFetchApp.fetch(url, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      }).getBlob();

      return {
        filename: safeName + ".xlsx",
        data: Utilities.base64Encode(blob.getBytes())
      };
    }
  } catch (e) {
    return { error: "Lỗi: " + e.toString() };
  }
}

/**
 * Chuyển mảng 2D thành CSV string
 */
function convertToCSV(data) {
  return data.map(function (row) {
    return row.map(function (cell) {
      var val = (cell === null || cell === undefined) ? "" : cell.toString();
      // Escape double quotes và wrap nếu chứa ký tự đặc biệt
      if (val.indexOf(",") > -1 || val.indexOf('"') > -1 || val.indexOf("\n") > -1) {
        val = '"' + val.replace(/"/g, '""') + '"';
      }
      return val;
    }).join(",");
  }).join("\r\n");
}
