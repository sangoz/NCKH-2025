function UpdateDashBoard() {
  try {
    Logger.log("=== BẮT ĐẦU UPDATE DASHBOARD ===");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName("Dashboard");

    if (!dashboardSheet) {
      SpreadsheetApp.getUi().alert("Không tìm thấy sheet Dashboard!");
      return;
    }

    const rulesSheet = ss.getSheetByName("Rules");
    if (!rulesSheet) {
      SpreadsheetApp.getUi().alert("Không tìm thấy sheet Rules! Vui lòng tạo Rules trước.");
      return;
    }

    // Đọc Rules data
    const rulesData = getRulesDataForUpdate();
    Logger.log(`📋 Đã load ${Object.keys(rulesData).length} rules`);

    // Đọc Dashboard data (bỏ qua header)
    const dashboardData = dashboardSheet.getDataRange().getValues();
    if (dashboardData.length <= 1) {
      SpreadsheetApp.getUi().alert("Dashboard không có dữ liệu để cập nhật!");
      return;
    }

    const header = dashboardData[0];
    const classIdx = 0;
    const groupIdx = 1;
    const assignmentIdx = 2;
    const dueDayIdx = 3;
    const submissionStatusIdx = 4;
    const extensionReqIdx = 5;
    const lastSubmissionIdx = 6;
    const overdueIdx = 7;

    Logger.log(`📊 Đang xử lý ${dashboardData.length - 1} dòng trong Dashboard...`);

    // Khởi tạo mảng kết quả cho batch write (tránh ghi cell-by-cell)
    const numRows = dashboardData.length - 1;
    const updateValues = [];
    for (let r = 0; r < numRows; r++) {
      updateValues.push([
        dashboardData[r + 1][dueDayIdx] || '',
        dashboardData[r + 1][submissionStatusIdx] || '',
        dashboardData[r + 1][extensionReqIdx] || '',
        dashboardData[r + 1][lastSubmissionIdx] || '',
        dashboardData[r + 1][overdueIdx] || ''
      ]);
    }

    // Process từng dòng (bỏ qua header)
    for (let i = 1; i < dashboardData.length; i++) {
      const row = dashboardData[i];
      const classname = (row[classIdx] || "").toString().trim();
      const groupname = (row[groupIdx] || "").toString().trim();
      const assignment = (row[assignmentIdx] || "").toString().trim();

      if (!classname || !groupname || !assignment) continue;

      Logger.log(`\n🔍 Processing: ${classname} / ${groupname} / ${assignment}`);

      // Lấy rules cho assignment này TRƯỚC (tránh Drive lookup khi không cần)
      const ruleKey = `${classname}|${assignment}`;
      const rule = rulesData[ruleKey];

      if (!rule) {
        Logger.log(`⚠️ Không có rule cho: ${ruleKey}`);
        updateValues[i - 1] = ["", "No rule", "", "", ""];
        continue;
      }

      // Due day - format từ rule
      let dueDay = "";
      if (rule.dueDay) {
        try {
          const dueDateObj = new Date(rule.dueDay);
          if (!isNaN(dueDateObj.getTime())) {
            dueDay = dueDateObj;
          } else {
            dueDay = rule.dueDay;
          }
        } catch (e) {
          Logger.log(`⚠️ Invalid due day format: ${rule.dueDay}`);
          dueDay = rule.dueDay;
        }
      }

      // Tìm folder trong Drive
      const folderInfo = findGroupAssignmentFolder(classname, groupname, assignment);

      if (!folderInfo) {
        Logger.log(`⚠️ Không tìm thấy folder: ${classname}/${groupname}/${assignment}`);
        updateValues[i - 1] = [dueDay, "Folder not found", "", "", "☐"];
        continue;
      }

      // Đếm file trong folder
      const files = folderInfo.folder.getFiles();
      const fileList = [];
      while (files.hasNext()) {
        const file = files.next();
        fileList.push({
          name: file.getName(),
          lastUpdated: file.getDateCreated(),
          mimeType: file.getMimeType()
        });
      }

      Logger.log(`📁 Tìm thấy ${fileList.length} files trong folder`);

      // Kiểm tra requirements
      const requiredCount = rule.numberOfFiles || 0;
      const actualCount = fileList.length;

      // Extension requirement - chỉ kiểm tra khi có file
      const requiredExtensions = rule.fileTypes || [];
      let extensionStatus = "";
      let isExtensionValid = true;

      if (actualCount > 0 && requiredExtensions.length > 0) {
        extensionStatus = "✓";
        // Đếm số lượng file theo từng loại yêu cầu
        const requiredFileCount = {};
        requiredExtensions.forEach(ext => {
          const cleanExt = ext.toLowerCase().replace('*', '');
          requiredFileCount[cleanExt] = (requiredFileCount[cleanExt] || 0) + 1;
        });

        // Đếm số lượng file thực tế theo từng loại
        const actualFileCount = {};
        fileList.forEach(f => {
          const fileName = f.name.toLowerCase();
          Object.keys(requiredFileCount).forEach(ext => {
            if (fileName.endsWith(ext)) {
              actualFileCount[ext] = (actualFileCount[ext] || 0) + 1;
            }
          });
        });

        // So sánh và tìm missing/extra
        const missing = [];
        const extra = [];

        // Kiểm tra thiếu
        Object.keys(requiredFileCount).forEach(ext => {
          const required = requiredFileCount[ext];
          const actual = actualFileCount[ext] || 0;
          if (actual < required) {
            const count = required - actual;
            const extName = ext.replace('.', '');
            missing.push(`${count} file ${extName}`);
          }
        });

        // Kiểm tra thừa (file không nằm trong danh sách yêu cầu)
        fileList.forEach(f => {
          const fileName = f.name.toLowerCase();
          const hasMatchingExt = Object.keys(requiredFileCount).some(ext => fileName.endsWith(ext));
          if (!hasMatchingExt) {
            extra.push(f.name);
          }
        });

        // Kiểm tra thừa trong từng loại
        Object.keys(actualFileCount).forEach(ext => {
          const actual = actualFileCount[ext];
          const required = requiredFileCount[ext] || 0;
          if (actual > required) {
            const count = actual - required;
            const extName = ext.replace('.', '');
            extra.push(`${count} file ${extName} thừa`);
          }
        });

        if (missing.length > 0 || extra.length > 0) {
          isExtensionValid = false;
          const issues = [];
          if (missing.length > 0) issues.push(`Thiếu: ${missing.join(', ')}`);
          if (extra.length > 0) issues.push(`Thừa: ${extra.join(', ')}`);
          extensionStatus = issues.join(' | ');
        }
      }

      // Submission status
      // Chỉ "Complete" khi: có nộp, đúng số lượng yêu cầu và không lỗi extension.
      let submissionStatus = "";
      if (actualCount === 0) {
        submissionStatus = "Not submitted";
      } else {
        const isCountValid = actualCount === requiredCount;
        if (isCountValid && isExtensionValid) {
          submissionStatus = `Complete (${actualCount}/${requiredCount})`;
        } else {
          submissionStatus = `Incomplete (${actualCount}/${requiredCount})`;
        }
      }

      // Last submission - file mới nhất
      let lastSubmission = "";
      let lastSubmissionDate = null;
      if (fileList.length > 0) {
        const latestFile = fileList.reduce((latest, file) => {
          return file.lastUpdated > latest.lastUpdated ? file : latest;
        }, fileList[0]);
        lastSubmissionDate = latestFile.lastUpdated;
        lastSubmission = lastSubmissionDate;
      }

      // Overdue - logic mới
      let overdue = "";
      if (rule.dueDay && lastSubmissionDate) {
        try {
          const dueDate = new Date(rule.dueDay);

          if (isExtensionValid) {
            // Đã tích extension requirement
            if (lastSubmissionDate <= dueDate) {
              // Due day > last submission -> bỏ trắng
              overdue = "";
            } else {
              // Due day < last submission -> checkbox đã check
              overdue = "☑";
            }
          } else {
            // Chưa tích extension requirement -> checkbox chưa check
            overdue = "☐";
          }
        } catch (e) {
          Logger.log(`⚠️ Invalid due day format: ${rule.dueDay}`);
          overdue = "Invalid date";
        }
      } else if (!lastSubmissionDate) {
        // Chưa có submission
        overdue = "☐";
      }

      // Cập nhật mảng kết quả
      updateValues[i - 1] = [dueDay, submissionStatus, extensionStatus, lastSubmission, overdue];

      Logger.log(`✅ Processed row ${i + 1}: ${submissionStatus}`);
    }

    // Batch write - ghi tất cả một lần (nhanh hơn nhiều so với cell-by-cell)
    if (numRows > 0) {
      dashboardSheet.getRange(2, dueDayIdx + 1, numRows, 5).setValues(updateValues);
      // Ép format hiển thị cho cột ngày tháng (tránh bị Sheets parse sai theo locale)
      dashboardSheet.getRange(2, dueDayIdx + 1, numRows, 1).setNumberFormat("dd/MM/yyyy HH:mm");
      dashboardSheet.getRange(2, lastSubmissionIdx + 1, numRows, 1).setNumberFormat("dd/MM/yyyy HH:mm");
    }

    Logger.log("=== HOÀN TẤT UPDATE DASHBOARD ===");
    SpreadsheetApp.getUi().alert("✅ Dashboard đã được cập nhật thành công!");

  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert("Lỗi khi cập nhật Dashboard: " + error.toString());
  }
}

/**
 * Tìm folder của một assignment trong group
 */
function findGroupAssignmentFolder(classname, groupname, assignmentName) {
  try {
    const rootFolder = DriveApp.getFoldersByName("userprofile");
    if (!rootFolder.hasNext()) {
      Logger.log("⚠️ Không tìm thấy folder 'userprofile'");
      return null;
    }

    const profileFolder = rootFolder.next();
    const classIter = profileFolder.getFoldersByName(classname);

    if (!classIter.hasNext()) {
      Logger.log(`⚠️ Không tìm thấy class folder: ${classname}`);
      return null;
    }

    const classFolder = classIter.next();
    const groupIter = classFolder.getFoldersByName(groupname);

    if (!groupIter.hasNext()) {
      Logger.log(`⚠️ Không tìm thấy group folder: ${groupname}`);
      return null;
    }

    const groupFolder = groupIter.next();

    // Tìm assignment folder (có thể nested)
    const found = findFolderRecursive(groupFolder, assignmentName);

    if (found) {
      Logger.log(`✓ Tìm thấy folder: ${found.getName()} (ID: ${found.getId()})`);
      return { folder: found, path: found.getName() };
    }

    return null;

  } catch (error) {
    Logger.log(`❌ Error finding folder: ${error.toString()}`);
    return null;
  }
}

/**
 * Tìm folder theo tên trong cây thư mục (đệ quy)
 */
function findFolderRecursive(parentFolder, targetName) {
  const folders = parentFolder.getFolders();

  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName() === targetName) {
      return folder;
    }

    // Tìm trong subfolder
    const found = findFolderRecursive(folder, targetName);
    if (found) return found;
  }

  return null;
}

/**
 * Hiển thị UI biểu đồ thống kê
 * Đọc toàn bộ dữ liệu Dashboard rồi inject vào HTML template
 * (tránh lỗi authorization khi dùng google.script.run từ modal dialog)
 */
function showDashboardChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("Dashboard");

  if (!dashboardSheet || dashboardSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("Không có dữ liệu trong sheet Dashboard!");
    return;
  }

  // Đọc toàn bộ dữ liệu dưới dạng text (tránh lỗi Date serialize)
  const data = dashboardSheet.getDataRange().getDisplayValues();
  const allRows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const classname = (row[0] || "").trim();
    if (!classname) continue;

    allRows.push({
      classname: classname,
      groupname: (row[1] || "").trim(),
      assignment: (row[2] || "").trim(),
      dueDay: (row[3] || "").trim(),
      submissionStatus: (row[4] || "").trim(),
      extensionRequirement: (row[5] || "").trim(),
      lastSubmission: (row[6] || "").trim(),
      overdue: (row[7] || "").trim()
    });
  }

  // Inject data vào HTML template
  const template = HtmlService.createTemplateFromFile('dashboardChartUI');
  template.injectedData = JSON.stringify(allRows);

  const html = template.evaluate()
    .setWidth(1000)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Biểu Đồ Thống Kê Submission');
}

/**
 * Đọc Rules sheet và trả về object mapping
 * Key: "classname|foldername"
 * Value: { numberOfFiles, fileTypes[], dueDay }
 */
function getRulesDataForUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName("Rules");

  if (!rulesSheet || rulesSheet.getLastRow() < 2) {
    return {};
  }

  const data = rulesSheet.getDataRange().getValues();
  const header = data[0].map(h => (h || "").toString().trim().toLowerCase());

  const classIdx = header.indexOf("class name");
  const folderIdx = header.indexOf("folder");
  const numberIdx = header.indexOf("number of file");

  if (classIdx === -1 || folderIdx === -1) {
    Logger.log("⚠️ Rules sheet thiếu column 'Class name' hoặc 'Folder'");
    return {};
  }

  const result = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const classname = (row[classIdx] || "").toString().trim();
    const folder = (row[folderIdx] || "").toString().trim();
    const numberOfFiles = numberIdx !== -1 ? parseInt(row[numberIdx]) || 0 : 0;

    if (!classname || !folder) continue;

    const key = `${classname}|${folder}`;

    // Tìm các cột file type và due day
    const fileTypes = [];
    let dueDay = "";

    if (numberOfFiles > 0) {
      for (let j = 0; j < numberOfFiles; j++) {
        const fileTypeCol = numberIdx + 1 + j * 3;
        const dueDayCol = numberIdx + 1 + j * 3 + 2;

        if (fileTypeCol < row.length && row[fileTypeCol]) {
          fileTypes.push(row[fileTypeCol].toString().trim());
        }

        if (dueDayCol < row.length && row[dueDayCol] && !dueDay) {
          dueDay = row[dueDayCol].toString().trim();
        }
      }
    }

    result[key] = {
      numberOfFiles: numberOfFiles,
      fileTypes: fileTypes,
      dueDay: dueDay
    };
  }

  Logger.log(`📋 Loaded ${Object.keys(result).length} rules`);
  return result;
}



