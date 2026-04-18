function onOpen () {
  // get the current spread sheet app
  let ui = SpreadsheetApp.getUi ()
  
  // whole menu
  ui.createMenu ('🖥️ToolScript')
    .addSubMenu (ui.createMenu ("Form") // ref: Form.gs
                  .addItem ('Form Creation', 'mainFormBuilder') 
                  .addItem ('Sync Result', 'manualSync')
                  .addItem ('Download class list', 'downloalClassList')) 
    .addSeparator ()

    .addSubMenu (ui.createMenu ("Folder") // ref: Folder.gs
                  .addItem('Open Folder Creator', 'showFolderCreatorSidebar')
                  .addItem ('Change Folder Structure', 'showChangeFolderStructureSidebar'))
    .addSeparator ()

    .addItem('Create Rules','showRulesSidebar')
    .addSeparator ()

    .addItem ('Update Permissions', 'updatePerrmissions') // ref: Application.js
    .addSeparator ()
    
    // 📧 Thêm Email Notification
    .addItem ('📧 Email Notifications', 'showEmailNotificationSidebar')
    .addSeparator ()
    
    //DashBoard
    .addSubMenu (ui.createMenu ('📊 Dashboard')
                  .addItem ('Update Dashboard', 'UpdateDashBoard')
                  .addItem ('View the Dashboard', 'showDashboardChart'))
    .addSeparator ()

    // khoa sheet
    .addItem('🔒 Khóa tất cả Sheet', 'lockAllSheets')
    .addSeparator()
 
    .addItem('🔑 Mở khóa tất cả Sheet', 'unlockAllSheets')
    .addSeparator()
    // ======================================
    // hihi
    // end create menu process
    .addToUi ()
   
}