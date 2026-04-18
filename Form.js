function mainFormBuilder () {
  var html = HtmlService.createHtmlOutputFromFile('formBuilderUI') // ref: formBuilderUI.html
    .setTitle (' ')
    .setWidth (300)
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

  // --- Get parent folder of the app spreadsheet ---
  const appFile = DriveApp.getFileById(ss.getId());
  const parents = appFile.getParents();
  let parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  // --- Find or create "temp" folder ---
  let tempFolder;
  const folders = parentFolder.getFoldersByName("temp");
  tempFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder("temp");

  const formtempFolder = getOrCreateFolder (tempFolder, '_form_temp_')

  // --- Create a subfolder for this form ---
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const formFolder = formtempFolder.createFolder(formId + "_" + timestamp);

  // --- Move form into the subfolder ---
  const formFile = DriveApp.getFileById(formId);
  formFile.moveTo(formFolder);

  // --- Create response sheet (R) ---
  const responseSs = SpreadsheetApp.create("Responses - " + formTitle);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, responseSs.getId());

  // Move response sheet into the same subfolder
  const responseFile = DriveApp.getFileById(responseSs.getId());
  responseFile.moveTo(formFolder);

  // --- Write log ---
  let logSheet = ss.getSheetByName("Form Logger");
  if (!logSheet) {
    logSheet = ss.insertSheet("Form Logger");
    logSheet.appendRow(["Timestamp", "Subject", "Class", "Publish URL", "Edit URL", "Folder URL", "Response Sheet"]);
    logSheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#d9ead3");
  }
  logSheet.appendRow([new Date(), subject, normalizedClassCode, publishUrl, editUrl, formFolder.getUrl(), responseSs.getUrl()]);

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

  // --- Find temp folder ---
  const appFile = DriveApp.getFileById(ss.getId());
  const parentFolder = appFile.getParents().hasNext() ? appFile.getParents().next() : DriveApp.getRootFolder();
  const tempFolder = getOrCreateFolder(parentFolder, "temp");
  const formFolder = getOrCreateFolder(tempFolder, "_form_temp_");

  const subfolders = formFolder.getFolders();

  let headersSet = false;

  while (subfolders.hasNext()) {
    const formFolder = subfolders.next();
    const files = formFolder.getFilesByType(MimeType.GOOGLE_SHEETS);

    while (files.hasNext()) {
      const responseFile = files.next();
      const responseSs = SpreadsheetApp.openById(responseFile.getId());
      const responseSheet = responseSs.getSheets()[0];
      let data = responseSheet.getDataRange().getValues();

      if (data.length > 1) {
        // --- Set header if not yet ---
        if (!headersSet) {
          const formHeaders = data[0].slice(1); // header from response sheet
          targetSheet.appendRow(["Class Name", ...formHeaders]);
          targetSheet.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#d9ead3");
          headersSet = true;
        }

        // --- Append data rows ---
        for (let i = 1; i < data.length; i++) {
          const row = data[i].slice(1);
          // Get class name from form title (response sheet name: "Responses - {ClassName}")
          const className = responseSs.getName().replace("Responses -", "").trim().split(" - ").pop();
          targetSheet.appendRow([className, ...row]);
        }
      }
    }
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

  Logger.log (result)

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
  return data.map(function(row) {
    return row.map(function(cell) {
      var val = (cell === null || cell === undefined) ? "" : cell.toString();
      // Escape double quotes và wrap nếu chứa ký tự đặc biệt
      if (val.indexOf(",") > -1 || val.indexOf('"') > -1 || val.indexOf("\n") > -1) {
        val = '"' + val.replace(/"/g, '""') + '"';
      }
      return val;
    }).join(",");
  }).join("\r\n");
}
