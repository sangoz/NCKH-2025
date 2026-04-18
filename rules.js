function showRulesSidebar() {
  // Tạo HTML
  const html = HtmlService.createHtmlOutputFromFile('rulesSidebar')
    .setTitle('Quản lý chi tiết Rules')
    .setWidth(1000) // ⬅️ Đặt chiều rộng (ví dụ 700px)
    .setHeight(800); // ⬅️ Đặt chiều cao (ví dụ 600px)

  // Hiển thị ra giữa
  SpreadsheetApp.getUi().showModalDialog(html, 'Quản lý chi tiết Rules');
}
/** Hàm lấy danh sách lớp */
function getClassList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rules');
  if (!sheet) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues(); // cột CLASSNAME
  const classes = [...new Set(data.map(r => r[0]).filter(c => c))];
  return classes;
}

/**
 * [ĐÃ CẬP NHẬT]
 * Lấy dữ liệu chi tiết trên sheet "Rules" theo classname
 * Trả về mảng {row, classname, folder, number, files: [...]}
 */
function getRulesDataByClass(classname) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rules');
  if (!sheet || sheet.getLastRow() < 2) return [];

  // Lấy cả giá trị (để check) và giá trị hiển thị (để giữ nguyên format date)
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const values = dataRange.getValues();
  const displayValues = dataRange.getDisplayValues(); // Lấy text y như trên sheet

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => (h || '').toString().trim().toLowerCase());

  const classIdx = header.indexOf('class name'); // Sửa 'classname' thành 'class name'
  const folderIdx = header.indexOf('folder');
  const numberIdx = header.indexOf('number of file'); // Sửa 'number of files' thành 'number of file'

  if (classIdx === -1 || folderIdx === -1 || numberIdx === -1) {
    throw new Error("Không tìm thấy các cột 'Class name', 'Folder', hoặc 'Number of file'. Vui lòng kiểm tra header.");
  }

  const out = [];

  for (let i = 0; i < values.length; i++) {
    const rowNum = i + 2;
    const row = values[i];
    const displayRow = displayValues[i];
    const classnameVal = row[classIdx];

    if (classnameVal !== classname) continue;

    const number = Number(row[numberIdx]) || 0;
    const files = [];

    if (number > 0) {
      for (let j = 1; j <= number; j++) {
        const fileColIdx = 3 + (j - 1) * 3; // File type 1 (col D) là index 3
        const reqColIdx = fileColIdx + 1;
        const dueColIdx = fileColIdx + 2;

        // Kiểm tra xem cột có tồn tại không
        if (dueColIdx >= sheet.getMaxColumns()) break;
        files.push({
          type: displayRow[fileColIdx] || '',
          req: displayRow[reqColIdx] || '',
          due: displayRow[dueColIdx] || ''
        });
      }
    }

    out.push({
      row: rowNum,
      classname: classnameVal.toString(),
      folder: row[folderIdx].toString(),
      number: number,
      files: files // Trả về mảng chi tiết các file
    });
  }

  return out;
}

/**
 * [ĐÃ CẬP NHẬT]
 * Hàm này dùng để trả về danh sách file extensions cho sidebar.
 */
function getFileExtensions() {
  return required_file_extension;
}
/**
 * [HÀM MỚI - THAY THẾ updateRulesNumberAndCreateColumns]
 * Nhận payload đầy đủ từ sidebar và cập nhật/thêm/xóa dữ liệu trên sheet.
 * Bao gồm: Cập nhật dữ liệu, xóa dữ liệu thừa trong hàng, xóa cột thừa, và xóa tất cả hàng rỗng (toàn bộ hàng).
 */
function saveRulesData(payload) {
  if (!payload || payload.length === 0) return { success: false, message: 'No data' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Rules');
  if (!sheet) throw new Error("Không tìm thấy sheet 'Rules'");

  // --- 1. Chuẩn hóa header và validation ---
  let header = sheet.getRange(1, 1, 1, Math.max(3, sheet.getLastColumn())).getValues()[0];
  let headerL = header.map(h => (h || '').toString().trim().toLowerCase());

  let classCol = headerL.indexOf('class name');
  let folderCol = headerL.indexOf('folder');
  let numberCol = headerL.indexOf('number of file');

  if (classCol === -1 || folderCol === -1 || numberCol === -1) {
    header = ['Class name', 'Folder', 'Number of file'];
    sheet.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#d9ead3');
    numberCol = 2; // index 2
  }

  const fileValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(required_file_extension, true)
    .setAllowInvalid(false)
    .build();
  const dateValidation = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true) // Cho phép nhập text, sheet sẽ tự báo lỗi
    .build();

  // --- 2. Tạo thêm cột nếu cần ---
  const maxNeed = payload.reduce((acc, p) => Math.max(acc, Number(p.number) || 0), 0);
  const requiredTotalCols = 3 + maxNeed * 3;
  const currentMaxCols = sheet.getMaxColumns();
  if (currentMaxCols < requiredTotalCols) {
    sheet.insertColumnsAfter(currentMaxCols, requiredTotalCols - currentMaxCols);
  }

  // Set header cho các cột mới (nếu có)
  for (let i = 1; i <= maxNeed; i++) {
    const col = 4 + (i - 1) * 3;
    // Chỉ set header nếu cột đó chưa có header
    if (!header[col - 1]) {
      sheet.getRange(1, col, 1, 3).setValues([[`File type ${i}`, `Requirements ${i}`, `Due day ${i}`]])
        .setFontWeight('bold').setBackground('#d9ead3');
    }
  }

  // --- 3. Ghi dữ liệu (Update/Create/Delete) ---
  payload.forEach(item => {
    const row = Number(item.row);
    const count = Math.max(0, Number(item.number) || 0);
    sheet.getRange(row, numberCol + 1).setValue(count); // Cập nhật số lượng file

    const files = item.files || [];

    // Duyệt qua *tất cả* các cột file có thể có (lên đến maxNeed)
    // để CẬP NHẬT (nếu i <= count) hoặc XÓA (nếu i > count)
    for (let i = 1; i <= maxNeed; i++) {
      const fileCol = 4 + (i - 1) * 3;
      const reqCol = fileCol + 1;
      const dueCol = fileCol + 2;

      // Đảm bảo không ghi đè lên cột không tồn tại
      if (dueCol > sheet.getMaxColumns()) continue;

      // Nếu i nằm trong số lượng file mới -> GHI DỮ LIỆU
      if (i <= count && files[i - 1]) {
        const fileData = files[i - 1];
        sheet.getRange(row, fileCol).setDataValidation(fileValidation).setValue(fileData.type || required_file_extension[0]);
        sheet.getRange(row, reqCol).setValue(fileData.req || '');
        sheet.getRange(row, dueCol).setDataValidation(dateValidation)
          .setValue(fileData.due || '') // Ghi giá trị text từ sidebar
          .setNumberFormat("dd/MM/yyyy hh:mm");
      }
      // Nếu i VƯỢT quá số lượng file mới -> XÓA DỮ LIỆU CŨ
      else {
        // Chỉ xóa nếu ô đó có nội dung, để tránh các lệnh ghi không cần thiết
        const rangeToClear = sheet.getRange(row, fileCol, 1, 3);
        if (rangeToClear.getDisplayValue() !== "") {
          rangeToClear.clearContent().clearDataValidations();
        }
      }
    }
  });

  // --- 4. XÓA CỘT THỪA TOÀN SHEET (SAU KHI CẬP NHẬT TẤT CẢ) ---
  // Tìm "Number of file" lớn nhất từ tất cả rows (từ hàng 2 trở đi)
  const allRows = sheet.getRange(2, numberCol + 1, sheet.getLastRow() - 1, 1).getValues();
  const maxNumberAcrossSheet = Math.max(...allRows.map(r => Number(r[0]) || 0));

  // Tính số cột cần: 3 cố định + maxNumber * 3
  const requiredColsAfterUpdate = 3 + maxNumberAcrossSheet * 3;
  const currentColsAfterUpdate = sheet.getMaxColumns();

  // Nếu có cột thừa, xóa chúng (bao gồm header)
  if (currentColsAfterUpdate > requiredColsAfterUpdate) {
    const colsToDelete = currentColsAfterUpdate - requiredColsAfterUpdate;
    // Xóa từ cột cuối, nhưng đảm bảo không xóa cột cố định (chỉ xóa từ cột 4 trở đi)
    const startDeleteCol = Math.max(4, requiredColsAfterUpdate + 1);
    if (startDeleteCol <= currentColsAfterUpdate) {
      sheet.deleteColumns(startDeleteCol, colsToDelete);
    }
  }

  // --- 5. XÓA TẤT CẢ HÀNG RỖNG (TOÀN BỘ HÀNG, TỪ HÀNG 2 TRỞ ĐI) ---
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const rowsToDelete = [];

    for (let row = 2; row <= lastRow; row++) {
      // Lấy toàn bộ giá trị của hàng (tất cả cột)
      const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      // Kiểm tra nếu tất cả ô trong hàng đều rỗng
      const isEmpty = rowValues.every(cell => !cell || cell.toString().trim() === '');
      if (isEmpty) {
        rowsToDelete.push(row);
      }
    }

    // Xóa từ dưới lên để không làm lệch index
    rowsToDelete.reverse().forEach(rowNum => {
      sheet.deleteRow(rowNum);
    });
  }

  return { success: true, message: 'Đã cập nhật dữ liệu và dọn dẹp sheet (xóa cột/hàng thừa) thành công!' };
}
