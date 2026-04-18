function lockAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // Lấy CHỦ SỞ HỮU (Owner) của file. Cách này không yêu cầu quyền đặc biệt.
  const owner = ss.getOwner(); 
  
  let lockedCount = 0;

  if (!owner) {
    // Trường hợp này hiếm gặp, ví dụ file nằm trong Shared Drive không có chủ sở hữu rõ ràng
    SpreadsheetApp.getUi().alert('Không thể xác định chủ sở hữu file. Không thể khóa sheet.');
    return;
  }

  sheets.forEach(sheet => {
    try {
      // Áp dụng bảo vệ cho toàn bộ sheet
      const protection = sheet.protect();
      
      // Xóa TẤT CẢ editor hiện tại khỏi quyền chỉnh sửa protection này
      const editors = protection.getEditors();
      protection.removeEditors(editors);
      
      // Chỉ thêm CHỦ SỞ HỮU (Owner) vào danh sách được phép chỉnh sửa
      protection.addEditor(owner);
      
      // (Tùy chọn) Ngăn không cho editor trong cùng domain chỉnh sửa
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
      
      protection.setDescription('Đã khóa bởi chủ sở hữu file');
      lockedCount++;

    } catch (e) {
      Logger.log(`Không thể khóa sheet: ${sheet.getName()}. Lỗi: ${e.message}`);
    }
  });
  
  // Thông báo hoàn tất
  SpreadsheetApp.getUi().alert(`Đã khóa thành công ${lockedCount} / ${sheets.length} sheets. Chỉ chủ sở hữu file mới có thể chỉnh sửa.`);
}
/**
 * Hàm này lặp qua tất cả các sheet và GỠ BỎ mọi chế độ bảo vệ (khóa) 
 * đang được áp dụng trên toàn bộ sheet.
 * Sau khi chạy, sheet sẽ quay về quyền chỉnh sửa mặc định của file.
 */
function unlockAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let unlockedCount = 0;

  sheets.forEach(sheet => {
    try {
      // Lấy TẤT CẢ các chế độ bảo vệ đang áp dụng cho sheet này
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      
      if (protections.length > 0) {
        // Lặp qua và gỡ bỏ từng cái một
        protections.forEach(protection => {
          protection.remove();
        });
        
        unlockedCount++;
        Logger.log(`Đã mở khóa sheet: ${sheet.getName()}`);
      }
      
    } catch (e) {
      Logger.log(`Không thể mở khóa sheet: ${sheet.getName()}. Lỗi: ${e.message}`);
    }
  });
  
  // Thông báo hoàn tất
  SpreadsheetApp.getUi().alert(`Đã mở khóa thành công ${unlockedCount} / ${sheets.length} sheets. Mọi người có quyền "Editor" đều có thể sửa lại.`);
}