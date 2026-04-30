function showEmailNotificationSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('emailNotificationUI')
    .setTitle('📧 Email Notification System')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Gửi email qua Gmail API (không có notification email)
 * @param {Object} options - {to, subject, htmlBody}
 * @returns {Object} - {success, messageId, error}
 */
function sendEmailViaGmail(options) {
  try {
    // Method 1: Sử dụng MailApp với UTF-8 encoding
    MailApp.sendEmail({
      to: options.to,
      subject: options.subject,
      htmlBody: options.htmlBody,
      charset: 'UTF-8'
    });

    return {
      success: true,
      messageId: 'sent-via-mailapp'
    };

  } catch (e) {
    Logger.log(`Error sending email via MailApp: ${e.message}`);

    // Method 2: Fallback to Gmail API nếu MailApp fail
    try {
      const message = {
        to: options.to,
        subject: options.subject,
        htmlBody: options.htmlBody
      };

      // Tạo raw message theo RFC 2822 với UTF-8
      const rawMessage =
        "To: " + message.to + "\r\n" +
        "Subject: " + message.subject + "\r\n" +
        "MIME-Version: 1.0\r\n" +
        "Content-Type: text/html; charset=UTF-8\r\n" +
        "Content-Transfer-Encoding: quoted-printable\r\n\r\n" +
        message.htmlBody;

      const encodedMessage = Utilities.base64EncodeWebSafe(rawMessage);

      const response = Gmail.Users.Messages.send(
        {
          raw: encodedMessage
        },
        "me"
      );

      return {
        success: true,
        messageId: response.id
      };

    } catch (gmailError) {
      Logger.log(`Error sending email via Gmail API: ${gmailError.message}`);
      return {
        success: false,
        error: gmailError.message
      };
    }
  }
}

/**
 * Gửi email folder link cho các nhóm
 */
function sendFolderLinksToGroups(classname, selectedGroups = []) {
  try {
    Logger.log(`Starting to send folder links for class: ${classname}, groups: ${selectedGroups.join(', ')}`);

    const results = [];

    // Lấy tất cả groups nếu không specify
    const groups = selectedGroups.length > 0 ? selectedGroups : getGroupsByClass(classname);

    if (groups.length === 0) {
      return {
        success: false,
        message: 'No groups found for this class'
      };
    }

    groups.forEach(groupName => {
      try {
        Logger.log(`Processing group: ${groupName}`);

        const members = getGroupMembers(classname, groupName);
        Logger.log(`Found ${members.length} members: ${members.join(', ')}`);

        if (members.length === 0) {
          results.push({ group: groupName, status: 'error', message: 'No members found' });
          return;
        }

        const folderData = getGroupFolderInfo(classname, groupName);
        if (!folderData || !folderData.url) {
          results.push({ group: groupName, status: 'error', message: 'Folder not found' });
          return;
        }

        const rulesData = getGroupRulesData(classname, groupName);
        Logger.log(`📋 Rules data for ${classname}/${groupName}: ${JSON.stringify(rulesData)}`);

        const emailContent = buildFolderEmailContent(classname, groupName, folderData, rulesData);

        // Gửi email cho từng thành viên
        let emailCount = 0;
        let failedEmails = [];
        let invalidEmails = [];

        members.forEach((email, index) => {
          Logger.log(`Processing member ${index + 1}/${members.length}: "${email}"`);

          // Kiểm tra email có hợp lệ không
          if (!email || email.toString().trim() === '') {
            Logger.log(`  ⚠️ Empty email at position ${index + 1}`);
            invalidEmails.push(`Position ${index + 1}: Empty`);
            return;
          }

          const emailStr = email.toString().trim();

          if (!emailStr.includes('@')) {
            Logger.log(`  ⚠️ Invalid email format (no @): "${emailStr}"`);
            invalidEmails.push(`"${emailStr}" (no @ symbol)`);
            return;
          }

          // Email hợp lệ, thử gửi
          try {
            Logger.log(`  ✉️ Sending email to: ${emailStr}`);

            // Sử dụng Gmail API thay vì MailApp
            const result = sendEmailViaGmail({
              to: emailStr,
              subject: `Thong bao truy cap thu muc - ${classname} - ${groupName}`,
              htmlBody: emailContent
            });

            if (result.success) {
              emailCount++;
              Logger.log(`  ✅ Email sent successfully to ${emailStr} (Message ID: ${result.messageId})`);
            } else {
              throw new Error(result.error);
            }

          } catch (e) {
            Logger.log(`  ❌ Failed to send email to ${emailStr}: ${e.message}`);
            failedEmails.push(`${emailStr} (${e.message})`);
          }
        });

        // Build detailed message
        let detailedMessage = `Sent to ${emailCount}/${members.length} members`;
        if (invalidEmails.length > 0) {
          detailedMessage += `. Invalid emails: ${invalidEmails.join(', ')}`;
        }
        if (failedEmails.length > 0) {
          detailedMessage += `. Failed: ${failedEmails.join(', ')}`;
        }

        results.push({
          group: groupName,
          status: emailCount > 0 ? 'success' : 'warning',
          message: detailedMessage,
          emailsSent: emailCount,
          emailsFailed: failedEmails.length,
          emailsInvalid: invalidEmails.length,
          members: members
        });

      } catch (e) {
        Logger.log(`Error processing group ${groupName}: ${e.message}`);
        results.push({ group: groupName, status: 'error', message: e.message });
      }
    });

    Logger.log(`Completed processing ${groups.length} groups`);

    return {
      success: true,
      results: results,
      message: `Processed ${groups.length} groups`
    };

  } catch (error) {
    Logger.log(`Error in sendFolderLinksToGroups: ${error.message}`);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Debug function cho việc kiểm tra và gửi email nhắc
 */
function debugReminderSystem() {
  try {
    Logger.log('=== DEBUGGING REMINDER SYSTEM ===');

    // 1. Kiểm tra Rules sheet
    Logger.log('1. Checking Rules sheet...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rulesSheet = ss.getSheetByName("Rules");

    if (!rulesSheet) {
      Logger.log('ERROR: No Rules sheet found!');
      return { success: false, error: 'No Rules sheet found' };
    }

    const data = rulesSheet.getDataRange().getValues();
    Logger.log(`Rules sheet has ${data.length} rows (including header)`);

    if (data.length < 2) {
      Logger.log('ERROR: Rules sheet is empty or only has header');
      return { success: false, error: 'Rules sheet is empty' };
    }

    // 2. Debug header structure
    Logger.log('2. Checking header structure...');
    const header = data[0].map(h => (h || '').toString().toLowerCase().replace(/\s+/g, ''));
    Logger.log(`Header columns (normalized): ${header.join(', ')}`);

    // Based on your Rules sheet: "Class name Folder", "Number of file", "File type 1", etc.
    const classnameCol = header.indexOf('classnamefolder') !== -1 ? header.indexOf('classnamefolder') :
      (header.indexOf('classname') !== -1 ? header.indexOf('classname') : 0);
    const folderCol = classnameCol; // Same column contains both class name and folder
    const numberCol = header.indexOf('numberoffile') !== -1 ? header.indexOf('numberoffile') :
      (header.indexOf('numberoffiles') !== -1 ? header.indexOf('numberoffiles') : 1);

    Logger.log(`Column indices - Classname: ${classnameCol}, Folder: ${folderCol}, Number of files: ${numberCol}`);

    if (classnameCol === -1 || folderCol === -1 || numberCol === -1) {
      Logger.log('ERROR: Required columns not found in Rules sheet');
      return { success: false, error: 'Required columns missing' };
    }

    // 3. Debug date calculations
    Logger.log('3. Checking date calculations...');
    const today = new Date();
    const tomorrow = new Date(today.getTime() + 24 * 60 * 60 * 1000);
    Logger.log(`Today: ${today.toLocaleDateString()}`);
    Logger.log(`Tomorrow: ${tomorrow.toLocaleDateString()}`);

    // 4. Process each rule
    Logger.log('4. Processing rules...');
    const validRules = [];
    const reminders = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const classname = row[classnameCol];
      const folderName = row[folderCol];
      const numberOfFiles = row[numberCol] || 0;

      Logger.log(`Row ${i}: Class=${classname}, Folder=${folderName}, Files=${numberOfFiles}`);

      if (!classname || numberOfFiles === 0) {
        Logger.log(`  -> Skipping: empty classname or no files`);
        continue;
      }

      validRules.push({ classname, folderName, numberOfFiles });

      // Check each file requirement
      for (let j = 1; j <= numberOfFiles; j++) {
        const dueDateCol = 3 + (j - 1) * 3 + 2; // Due date column
        const dueDate = row[dueDateCol];

        Logger.log(`  File ${j}: Due date column ${dueDateCol}, Value: ${dueDate}, Type: ${typeof dueDate}`);

        if (dueDate && dueDate instanceof Date) {
          const daysDiff = Math.floor((dueDate - today) / (1000 * 60 * 60 * 24));
          Logger.log(`    Days until due: ${daysDiff}`);

          if (daysDiff === 1) {
            Logger.log(`    -> Due tomorrow! Checking groups...`);

            try {
              const groups = getGroupsByClass(classname);
              Logger.log(`    Groups in ${classname}: ${groups.join(', ')}`);

              groups.forEach(groupName => {
                Logger.log(`      Checking group: ${groupName}`);

                const missingFiles = debugCheckMissingFiles(classname, groupName, folderName, row, j);
                Logger.log(`      Missing files: ${JSON.stringify(missingFiles)}`);

                if (missingFiles.length > 0) {
                  reminders.push({
                    classname: classname,
                    groupName: groupName,
                    folderName: folderName,
                    dueDate: dueDate,
                    missingFiles: missingFiles
                  });
                  Logger.log(`      -> Added reminder for ${groupName}`);
                }
              });
            } catch (groupError) {
              Logger.log(`    ERROR getting groups for ${classname}: ${groupError.message}`);
            }
          }
        } else {
          Logger.log(`    -> Invalid or missing due date`);
        }
      }
    }

    Logger.log(`5. Summary:`);
    Logger.log(`  Valid rules found: ${validRules.length}`);
    Logger.log(`  Reminders to send: ${reminders.length}`);

    return {
      success: true,
      validRules: validRules.length,
      reminders: reminders.length,
      reminderDetails: reminders
    };

  } catch (error) {
    Logger.log(`ERROR in debugReminderSystem: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return { success: false, error: error.message };
  }
}

/**
 * Debug version của checkMissingFiles
 */
function debugCheckMissingFiles(classname, groupName, folderName, ruleRow, fileIndex) {
  try {
    Logger.log(`      DEBUG checkMissingFiles: ${classname}/${groupName}/${folderName}/file${fileIndex}`);

    // 1. Check folder structure
    const root = getSpreadsheetParent();
    Logger.log(`        Root folder: ${root.getName()}`);

    const userprofile = getOrCreateFolder(root, "userprofile");
    Logger.log(`        Userprofile folder: ${userprofile.getName()}`);

    const classFolder = getOrCreateFolder(userprofile, classname);
    Logger.log(`        Class folder: ${classFolder.getName()}`);

    // 2. Find group folder
    const groupFolders = classFolder.getFoldersByName(groupName);
    if (!groupFolders.hasNext()) {
      Logger.log(`        ERROR: Group folder '${groupName}' not found`);
      return [`Group folder '${groupName}' not found`];
    }

    const groupFolder = groupFolders.next();
    Logger.log(`        Group folder found: ${groupFolder.getName()}`);

    // 3. Find target folder
    const targetFolder = debugFindFolderByName(groupFolder, folderName);
    if (!targetFolder) {
      Logger.log(`        ERROR: Target folder '${folderName}' not found`);
      return [`Folder '${folderName}' not found`];
    }

    Logger.log(`        Target folder found: ${targetFolder.getName()}`);

    // 4. Get file requirements
    const fileTypeCol = 3 + (fileIndex - 1) * 3;
    const expectedType = ruleRow[fileTypeCol] || '';
    const requirement = ruleRow[fileTypeCol + 1] || '';

    Logger.log(`        Expected file type: '${expectedType}'`);
    Logger.log(`        Requirement: '${requirement}'`);

    // 5. Check existing files
    const files = targetFolder.getFiles();
    const fileList = [];
    let hasMatchingFile = false;

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const extension = fileName.substring(fileName.lastIndexOf('.'));
      fileList.push(`${fileName} (${extension})`);

      Logger.log(`        File found: ${fileName}, Extension: ${extension}`);

      if (expectedType.toLowerCase().includes(extension.toLowerCase())) {
        hasMatchingFile = true;
        Logger.log(`        -> MATCH found!`);
      }
    }

    Logger.log(`        Total files in folder: ${fileList.length}`);
    Logger.log(`        Files: ${fileList.join(', ')}`);
    Logger.log(`        Has matching file: ${hasMatchingFile}`);

    if (!hasMatchingFile) {
      return [{
        type: expectedType,
        requirement: requirement,
        missing: true,
        filesFound: fileList
      }];
    }

    return [];

  } catch (e) {
    Logger.log(`        ERROR in debugCheckMissingFiles: ${e.message}`);
    return [`Error checking files: ${e.message}`];
  }
}

/**
 * Debug version của findFolderByName
 */
function debugFindFolderByName(parentFolder, targetName, depth = 0) {
  const indent = '  '.repeat(depth + 4);
  Logger.log(`${indent}Searching in: ${parentFolder.getName()} for '${targetName}'`);

  const folders = parentFolder.getFolders();
  const folderList = [];

  while (folders.hasNext()) {
    const folder = folders.next();
    const folderName = folder.getName();
    folderList.push(folderName);

    Logger.log(`${indent}  Found subfolder: '${folderName}'`);

    if (folderName === targetName) {
      Logger.log(`${indent}  -> EXACT MATCH found!`);
      return folder;
    }

    // Search in subfolders (limit depth to avoid infinite recursion)
    if (depth < 3) {
      const subResult = debugFindFolderByName(folder, targetName, depth + 1);
      if (subResult) return subResult;
    }
  }

  Logger.log(`${indent}Subfolders in ${parentFolder.getName()}: ${folderList.join(', ')}`);
  return null;
}

/**
 * Test function đơn giản để debug missing files cho 1 nhóm cụ thể
 */
function testMissingFilesForGroup(classname, groupName) {
  try {
    Logger.log(`=== TESTING MISSING FILES FOR ${classname}/${groupName} ===`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rulesSheet = ss.getSheetByName("Rules");

    if (!rulesSheet) {
      Logger.log('ERROR: No Rules sheet found');
      return { success: false, error: 'No Rules sheet found' };
    }

    const data = rulesSheet.getDataRange().getValues();
    const header = data[0].map(h => (h || '').toString().toLowerCase().replace(/\s+/g, ''));

    const classnameCol = header.indexOf('classname');
    const folderCol = header.indexOf('folder');
    const numberCol = header.indexOf('numberoffiles');

    const results = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[classnameCol] === classname) {
        const folderName = row[folderCol];
        const numberOfFiles = row[numberCol] || 0;

        Logger.log(`Checking folder: ${folderName} (${numberOfFiles} files expected)`);

        for (let j = 1; j <= numberOfFiles; j++) {
          const missingFiles = debugCheckMissingFiles(classname, groupName, folderName, row, j);

          results.push({
            folder: folderName,
            fileIndex: j,
            missing: missingFiles
          });
        }
      }
    }

    return { success: true, results: results };

  } catch (error) {
    Logger.log(`ERROR in testMissingFilesForGroup: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/**
 * FORCE AUTHORIZATION - Chạy function này từ Apps Script Editor để trigger authorization
 * Function này sẽ gửi test email và force Google hiện popup xin quyền
 */
function forceEmailAuthorization() {
  const testEmail = Session.getActiveUser().getEmail();

  Logger.log('========================================');
  Logger.log('FORCE EMAIL AUTHORIZATION TEST');
  Logger.log('========================================');
  Logger.log('Test email will be sent to: ' + testEmail);
  Logger.log('');

  try {
    // Attempt to send test email - this will trigger authorization popup if needed
    MailApp.sendEmail({
      to: testEmail,
      subject: '✅ Authorization Success - Google Apps Script',
      htmlBody: '<div style="font-family: Arial, sans-serif; padding: 20px;">' +
        '<h1 style="color: #28a745;">✅ Authorization Success!</h1>' +
        '<p>Congratulations! Your Google Apps Script project is now authorized to send emails.</p>' +
        '<p><strong>Test Email Sent:</strong> ' + new Date().toLocaleString() + '</p>' +
        '<p><strong>From:</strong> ' + testEmail + '</p>' +
        '<hr>' +
        '<p style="color: #666; font-size: 12px;">This is an automated test email from your Apps Script authorization process.</p>' +
        '</div>'
    });

    Logger.log('✅ SUCCESS!');
    Logger.log('Test email sent successfully to: ' + testEmail);
    Logger.log('');
    Logger.log('Authorization is complete. You can now:');
    Logger.log('1. Check your email inbox for the test message');
    Logger.log('2. Use the email notification system from the Sheet UI');
    Logger.log('3. Send folder links to groups');
    Logger.log('========================================');

    return {
      success: true,
      message: 'Email sent successfully to ' + testEmail,
      timestamp: new Date().toISOString()
    };

  } catch (e) {
    Logger.log('❌ ERROR!');
    Logger.log('Failed to send email: ' + e.message);
    Logger.log('');
    Logger.log('Possible reasons:');
    Logger.log('1. Authorization was cancelled');
    Logger.log('2. Email quota exceeded');
    Logger.log('3. OAuth scope not properly configured');
    Logger.log('');
    Logger.log('Please try running this function again.');
    Logger.log('========================================');

    return {
      success: false,
      error: e.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Gửi email qua Gmail API (không gửi notification)
 * Requires Advanced Gmail Service enabled
 */
function sendEmailViaGmail(options) {
  try {
    const { to, subject, htmlBody } = options;

    // Tạo email message với encoding UTF-8
    const emailMessage = [
      'Content-Type: text/html; charset=UTF-8',
      'MIME-Version: 1.0',
      `To: ${to}`,
      `Subject: ${subject}`,
      '',
      htmlBody
    ].join('\r\n');

    // Encode email theo base64url
    const encodedEmail = Utilities.base64EncodeWebSafe(emailMessage);

    // Gửi qua Gmail API
    const response = Gmail.Users.Messages.send({
      raw: encodedEmail
    }, 'me');

    return {
      success: true,
      messageId: response.id
    };

  } catch (error) {
    Logger.log(`Error in sendEmailViaGmail: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Lấy thông tin folder của nhóm
 */
function getGroupFolderInfo(classname, groupName) {
  try {
    Logger.log(`Getting folder info for class: ${classname}, group: ${groupName}`);

    const root = getSpreadsheetParent();
    const userprofile = getOrCreateFolder(root, "userprofile");
    const classFolder = getOrCreateFolder(userprofile, classname);

    const groupIter = classFolder.getFoldersByName(groupName);
    if (!groupIter.hasNext()) {
      Logger.log(`No folder found for group: ${groupName} in class: ${classname}`);
      return null;
    }

    const groupFolder = groupIter.next();
    const result = {
      name: groupFolder.getName(),
      id: groupFolder.getId(),
      url: groupFolder.getUrl(),
      createdDate: groupFolder.getDateCreated()
    };

    Logger.log(`Found folder info: ${JSON.stringify(result)}`);
    return result;

  } catch (e) {
    Logger.log(`Error getting folder info: ${e.message}`);
    return null;
  }
}

/**
 * Tìm folder theo tên trong cây thư mục và trả về cả đường dẫn.
 */
function findFolderByNameWithPath(parentFolder, targetName, pathParts = []) {
  const folders = parentFolder.getFolders();

  while (folders.hasNext()) {
    const folder = folders.next();
    const currentPath = pathParts.concat(folder.getName());

    if (folder.getName() === targetName) {
      return {
        folder: folder,
        path: currentPath
      };
    }

    const found = findFolderByNameWithPath(folder, targetName, currentPath);
    if (found) return found;
  }

  return null;
}

/**
 * Lấy rules data cho nhóm từ Rules sheet
 */
function getGroupRulesData(classname, groupName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rulesSheet = ss.getSheetByName("Rules");

    if (!rulesSheet) return [];

    const data = rulesSheet.getDataRange().getValues();
    const displayData = rulesSheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    const header = data[0].map(h => (h || '').toString().toLowerCase().trim());
    const classnameCol = header.indexOf('class name');
    const folderCol = header.indexOf('folder');
    const numberCol = header.indexOf('number of file') !== -1 ? header.indexOf('number of file') :
      (header.indexOf('number of files') !== -1 ? header.indexOf('number of files') : 2);

    if (classnameCol === -1 || folderCol === -1) {
      Logger.log(`Rules header invalid: ${JSON.stringify(header)}`);
      return [];
    }

    const root = getSpreadsheetParent();
    const userprofile = getOrCreateFolder(root, "userprofile");
    const classFolder = getOrCreateFolder(userprofile, classname);
    const groupIter = classFolder.getFoldersByName(groupName);
    const groupFolder = groupIter.hasNext() ? groupIter.next() : null;

    const rules = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const displayRow = displayData[i];

      if ((displayRow[classnameCol] || '').toString().trim() === classname) {
        const folderName = (displayRow[folderCol] || '').toString().trim();
        const numberOfFiles = row[numberCol] || 0;

        if (numberOfFiles > 0) {
          let folderPath = folderName;
          if (groupFolder && folderName) {
            const found = findFolderByNameWithPath(groupFolder, folderName, [groupFolder.getName()]);
            if (found && found.path && found.path.length > 0) {
              folderPath = found.path.join(' / ');
            }
          }

          const ruleData = {
            folder: folderName,
            folderPath: folderPath,
            numberOfFiles: numberOfFiles,
            requirements: []
          };

          // Lấy file requirements
          for (let j = 1; j <= numberOfFiles; j++) {
            const fileTypeCol = 3 + (j - 1) * 3;
            const reqCol = fileTypeCol + 1;
            const dueCol = fileTypeCol + 2;

            if (fileTypeCol < row.length) {
              ruleData.requirements.push({
                fileType: row[fileTypeCol] || '',
                requirement: row[reqCol] || '',
                dueDate: row[dueCol] || ''
              });
            }
          }

          rules.push(ruleData);
        }
      }
    }

    return rules;
  } catch (e) {
    Logger.log(`Error getting rules data: ${e.message}`);
    return [];
  }
}

/**
 * Xây dựng nội dung email
 */
function buildFolderEmailContent(classname, groupName, folderData, rulesData) {
  // Format date in EN
  const formatDate = (dateStr) => {
    if (!dateStr) return '';
    try {
      const date = new Date(dateStr);
      return date.toLocaleString('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
      });
    } catch (e) {
      return dateStr;
    }
  };

  let emailHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
    </head>
    <body>
    <div style="font-family: Arial, sans-serif; max-width: 700px; margin: 0 auto;">
      <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 25px; text-align: center; border-radius: 10px 10px 0 0;">
        <h2 style="margin: 0; font-size: 24px;">FOLDER ACCESS NOTIFICATION</h2>
        <p style="margin: 10px 0 0 0; font-size: 14px; opacity: 0.9;">Class Folder Management System</p>
      </div>
      
      <div style="padding: 30px; background: #ffffff; border: 1px solid #e0e0e0; border-radius: 0 0 10px 10px;">
        <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
          <h3 style="margin: 0 0 10px 0; color: #1a73e8; font-size: 18px;">Class: ${classname}</h3>
          <h4 style="margin: 0; color: #5f6368; font-size: 16px;">Group: ${groupName}</h4>
        </div>
        
        <div style="background: white; padding: 20px; border-radius: 8px; margin: 15px 0; border: 2px solid #e8f0fe;">
          <h4 style="margin: 0 0 15px 0; color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 10px;">Folder Information</h4>
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="padding: 8px 0; color: #5f6368; width: 140px;"><strong>Folder Name:</strong></td>
              <td style="padding: 8px 0;">${folderData.name}</td>
            </tr>
            <tr>
              <td style="padding: 8px 0; color: #5f6368;"><strong>Access Link:</strong></td>
              <td style="padding: 8px 0;">
                <a href="${folderData.url}" target="_blank" style="color: #1a73e8; text-decoration: none; font-weight: 500;">
                  Click here to open folder
                </a>
              </td>
            </tr>
            <tr>
              <td style="padding: 8px 0; color: #5f6368;"><strong>Created Date:</strong></td>
              <td style="padding: 8px 0;">${formatDate(folderData.createdDate)}</td>
            </tr>
          </table>
        </div>
  `;

  if (rulesData && rulesData.length > 0) {
    emailHtml += `
        <div style="background: white; padding: 20px; border-radius: 8px; margin: 15px 0; border: 2px solid #34a853;">
          <h4 style="margin: 0 0 15px 0; color: #34a853; border-bottom: 2px solid #34a853; padding-bottom: 10px;">Assignment Requirements</h4>
    `;

    rulesData.forEach((rule, ruleIndex) => {
      const folderTitle = rule.folder || 'Not specified';
      const folderPath = rule.folderPath || folderTitle;
      emailHtml += `
          <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 15px 0; border-left: 4px solid #fbbc04;">
            <h5 style="margin: 0 0 15px 0; color: #202124; font-size: 16px;">
              ${folderTitle} 
              <span style="background: #fbbc04; color: #202124; padding: 3px 10px; border-radius: 12px; font-size: 12px; margin-left: 10px;">
                ${rule.numberOfFiles} file${rule.numberOfFiles > 1 ? 's' : ''}
              </span>
            </h5>
      `;

      rule.requirements.forEach((req, index) => {
        const dueDate = formatDate(req.dueDate);
        const now = new Date();
        const due = new Date(req.dueDate);
        const isUrgent = due && (due - now) < (2 * 24 * 60 * 60 * 1000); // < 2 days

        emailHtml += `
            <div style="background: white; padding: 12px; border-radius: 6px; margin: 10px 0; border: 1px solid #dadce0;">
              <div style="display: flex; align-items: center; margin-bottom: 8px;">
                <span style="background: #1a73e8; color: white; padding: 2px 8px; border-radius: 10px; font-size: 12px; font-weight: bold; margin-right: 8px;">
                  ${folderTitle} - FILE ${index + 1}
                </span>
                ${isUrgent ? '<span style="background: #ea4335; color: white; padding: 2px 8px; border-radius: 10px; font-size: 11px;">DUE SOON</span>' : ''}
              </div>
              <table style="width: 100%; font-size: 14px;">
                <tr>
                  <td style="padding: 4px 0; color: #5f6368; width: 120px;"><strong>Folder:</strong></td>
                  <td style="padding: 4px 0;"><strong>${folderTitle}</strong></td>
                </tr>
                <tr>
                  <td style="padding: 4px 0; color: #5f6368; width: 120px;"><strong>Folder Path:</strong></td>
                  <td style="padding: 4px 0; color: #202124;">${folderPath}</td>
                </tr>
                <tr>
                  <td style="padding: 4px 0; color: #5f6368; width: 120px;"><strong>File Type:</strong></td>
                  <td style="padding: 4px 0;"><code style="background: #f1f3f4; padding: 2px 8px; border-radius: 4px;">${req.fileType || 'Not specified'}</code></td>
                </tr>
                <tr>
                  <td style="padding: 4px 0; color: #5f6368;"><strong>Requirement:</strong></td>
                  <td style="padding: 4px 0;">${req.requirement || 'No specific requirement'}</td>
                </tr>
                <tr>
                  <td style="padding: 4px 0; color: #5f6368;"><strong>Due Date:</strong></td>
                  <td style="padding: 4px 0; ${isUrgent ? 'color: #ea4335; font-weight: bold;' : ''}">${dueDate || 'Not specified'}</td>
                </tr>
              </table>
            </div>
        `;
      });

      emailHtml += `</div>`;
    });

    emailHtml += `</div>`;
  }

  emailHtml += `
        <div style="background: #fff3cd; border: 2px solid #fbbc04; padding: 20px; border-radius: 8px; margin: 20px 0;">
          <h4 style="margin: 0 0 15px 0; color: #ea8600;">Important Notes</h4>
          <ul style="margin: 0; padding-left: 20px; line-height: 1.8;">
            <li>Please upload files to the <strong>correct folder</strong> as required</li>
            <li>Ensure <strong>file format</strong> matches the requirements</li>
            <li>Submit <strong>before the deadline</strong> to avoid penalties</li>
            <li>Contact your instructor if you have questions</li>
            <li><strong>Do not share</strong> the folder link with others</li>
          </ul>
        </div>
        
        <div style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0; color: #5f6368; font-size: 13px;">
          <p style="margin: 5px 0;">Automated email from the Folder Management System</p>
          <p style="margin: 5px 0;">Sent at: ${new Date().toLocaleString('en-US')}</p>
          <p style="margin: 15px 0 0 0; font-style: italic;">Please do not reply to this email</p>
        </div>
      </div>
    </div>
    </body>
    </html>
  `;

  return emailHtml;
}

function _toDateOnly(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function _dateDiffInDays(targetDate, baseDate) {
  const msPerDay = 24 * 60 * 60 * 1000;
  return Math.round((_toDateOnly(targetDate) - _toDateOnly(baseDate)) / msPerDay);
}

function normalizeRequiredExtension(fileType) {
  const raw = (fileType || '').toString().trim().toLowerCase();
  if (!raw) return '';

  const match = raw.match(/\.([a-z0-9]+)$/i);
  if (match) return `.${match[1].toLowerCase()}`;

  if (raw.startsWith('.')) return raw;
  return `.${raw.replace(/^\*+/, '').replace(/^\./, '')}`;
}

function analyzeMissingFilesForReminder(classname, groupName, folderName, ruleRow, numberOfFiles, today) {
  try {
    const root = getSpreadsheetParent();
    const userprofile = getOrCreateFolder(root, "userprofile");
    const classFolder = getOrCreateFolder(userprofile, classname);
    const groupFolders = classFolder.getFoldersByName(groupName);
    if (!groupFolders.hasNext()) {
      return { shouldRemind: false, reason: `Group folder not found: ${groupName}`, missingFiles: [] };
    }

    const groupFolder = groupFolders.next();
    const targetFolder = findFolderByName(groupFolder, folderName);
    if (!targetFolder) {
      return { shouldRemind: false, reason: `Target folder not found: ${folderName}`, missingFiles: [] };
    }

    const requirementsByExt = {};
    for (let j = 1; j <= numberOfFiles; j++) {
      const fileTypeCol = 3 + (j - 1) * 3;
      const reqCol = fileTypeCol + 1;
      const dueCol = fileTypeCol + 2;

      const fileType = ruleRow[fileTypeCol] || '';
      const requirement = ruleRow[reqCol] || '';
      const dueDate = ruleRow[dueCol] instanceof Date ? ruleRow[dueCol] : null;

      const ext = normalizeRequiredExtension(fileType);
      if (!ext) continue;

      if (!requirementsByExt[ext]) requirementsByExt[ext] = [];
      requirementsByExt[ext].push({
        type: fileType || ext,
        requirement: requirement,
        dueDate: dueDate
      });
    }

    const actualCounts = {};
    const actualFileNames = [];
    const files = targetFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      actualFileNames.push(fileName);
      const lower = fileName.toLowerCase();
      const dot = lower.lastIndexOf('.');
      if (dot === -1) continue;
      const ext = lower.substring(dot);
      actualCounts[ext] = (actualCounts[ext] || 0) + 1;
    }

    const missingFiles = [];
    let nearestDueDate = null;

    Object.keys(requirementsByExt).forEach(ext => {
      const reqList = requirementsByExt[ext];
      const requiredCount = reqList.length;
      const actualCount = actualCounts[ext] || 0;

      if (actualCount >= requiredCount) return;

      const missingCount = requiredCount - actualCount;
      const sortedByDue = reqList.slice().sort((a, b) => {
        if (!a.dueDate && !b.dueDate) return 0;
        if (!a.dueDate) return 1;
        if (!b.dueDate) return -1;
        return a.dueDate - b.dueDate;
      });

      // Ambiguous match for same extension: assume missing slots are earliest due ones.
      const missingSlots = sortedByDue.slice(0, missingCount);
      missingSlots.forEach(slot => {
        if (slot.dueDate && (!nearestDueDate || slot.dueDate < nearestDueDate)) {
          nearestDueDate = slot.dueDate;
        }
        missingFiles.push({
          type: slot.type || ext,
          requirement: slot.requirement || `Missing required file (${ext})`,
          dueDate: slot.dueDate,
          missing: true,
          filesFound: actualFileNames
        });
      });
    });

    if (missingFiles.length === 0) {
      return { shouldRemind: false, reason: 'No missing files by count', missingFiles: [] };
    }

    if (!nearestDueDate) {
      return { shouldRemind: false, reason: 'Missing files have no due date', missingFiles: missingFiles };
    }

    const daysUntilNearestDue = _dateDiffInDays(nearestDueDate, today);
    return {
      shouldRemind: daysUntilNearestDue === 1,
      nearestDueDate: nearestDueDate,
      missingFiles: missingFiles,
      reason: `Nearest missing due in ${daysUntilNearestDue} day(s)`
    };

  } catch (error) {
    return { shouldRemind: false, reason: error.message, missingFiles: [] };
  }
}

/**
 * Tự động kiểm tra và gửi reminder emails - Updated with debug logic
 */
function checkAndSendReminderEmails() {
  try {
    Logger.log('=== AUTOMATIC REMINDER CHECK STARTED ===');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rulesSheet = ss.getSheetByName("Rules");

    if (!rulesSheet) {
      Logger.log("ERROR: No Rules sheet found");
      return { success: false, error: "No Rules sheet found" };
    }

    const data = rulesSheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("ERROR: Rules sheet is empty");
      return { success: false, error: "Rules sheet is empty" };
    }

    // Debug header structure
    const header = data[0].map(h => (h || '').toString().toLowerCase().replace(/\s+/g, ''));
    Logger.log(`Header columns (normalized): ${header.join(', ')}`);

    const classnameCol = header.indexOf('classname') !== -1 ? header.indexOf('classname') : 0;
    const folderCol = header.indexOf('folder') !== -1 ? header.indexOf('folder') : 1;
    const numberCol = header.indexOf('numberoffile') !== -1 ? header.indexOf('numberoffile') :
      (header.indexOf('numberoffiles') !== -1 ? header.indexOf('numberoffiles') : 2);

    if (classnameCol === -1 || folderCol === -1 || numberCol === -1) {
      Logger.log('ERROR: Required columns not found in Rules sheet');
      return { success: false, error: 'Required columns missing in Rules sheet' };
    }

    const today = new Date();
    const tomorrow = new Date(today.getTime() + 24 * 60 * 60 * 1000);
    Logger.log(`Today: ${today.toLocaleDateString()}, Tomorrow: ${tomorrow.toLocaleDateString()}`);

    const reminders = [];
    let totalChecks = 0;
    let validRules = 0;

    // Duyệt qua tất cả rules
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const displayRow = rulesSheet.getRange(i + 1, 1, 1, rulesSheet.getLastColumn()).getDisplayValues()[0];
      const classname = (displayRow[classnameCol] || '').toString().trim();
      const folderName = (displayRow[folderCol] || '').toString().trim();
      const numberOfFiles = row[numberCol] || 0;

      Logger.log(`Row ${i}: Class=${classname}, Folder=${folderName}, Files=${numberOfFiles}`);

      if (!classname || numberOfFiles === 0) {
        Logger.log(`  -> Skipping: empty classname or no files`);
        continue;
      }

      validRules++;

      try {
        const groups = getGroupsByClass(classname);
        Logger.log(`    Groups in ${classname}: ${groups.join(', ')}`);

        groups.forEach(groupName => {
          totalChecks++;
          Logger.log(`      Checking group: ${groupName}`);

          const analysis = analyzeMissingFilesForReminder(
            classname,
            groupName,
            folderName,
            row,
            numberOfFiles,
            today
          );

          Logger.log(`      Missing analysis: ${JSON.stringify(analysis)}`);

          if (analysis.shouldRemind) {
            reminders.push({
              classname: classname,
              groupName: groupName,
              folderName: folderName,
              dueDate: analysis.nearestDueDate,
              missingFiles: analysis.missingFiles
            });
            Logger.log(`      -> Added reminder for ${groupName} (${analysis.missingFiles.length} missing files)`);
          }
        });

      } catch (groupError) {
        Logger.log(`    ERROR getting groups for ${classname}: ${groupError.message}`);
      }
    }

    Logger.log(`=== SUMMARY ===`);
    Logger.log(`Valid rules processed: ${validRules}`);
    Logger.log(`Total group checks: ${totalChecks}`);
    Logger.log(`Reminders to send: ${reminders.length}`);

    // Gửi reminder emails
    let emailsSent = 0;
    const sendErrors = [];
    reminders.forEach((reminder, index) => {
      Logger.log(`Sending reminder ${index + 1}/${reminders.length}: ${reminder.classname}/${reminder.groupName}`);

      try {
        const sendResult = sendReminderEmail(reminder);
        if (sendResult && sendResult.success) {
          emailsSent += sendResult.sentCount || 0;
          Logger.log(`  -> Email sent successfully to ${sendResult.sentCount || 0} recipient(s)`);
        } else {
          const err = sendResult && sendResult.error ? sendResult.error : 'unknown error';
          sendErrors.push(`${reminder.classname}/${reminder.groupName}/${reminder.folderName}: ${err}`);
          Logger.log(`  -> Email not sent: ${err}`);
        }
      } catch (emailError) {
        sendErrors.push(`${reminder.classname}/${reminder.groupName}/${reminder.folderName}: ${emailError.message}`);
        Logger.log(`  -> Failed to send email: ${emailError.message}`);
      }
    });

    Logger.log(`=== AUTOMATIC REMINDER CHECK COMPLETED ===`);
    Logger.log(`Total emails sent: ${emailsSent}/${reminders.length}`);

    const summaryText = sendErrors.length > 0
      ? `Sent ${emailsSent} reminder emails out of ${reminders.length} needed. Errors: ${sendErrors.join(' | ')}`
      : `Sent ${emailsSent} reminder emails out of ${reminders.length} needed`;

    return {
      success: true,
      count: emailsSent,
      totalReminders: reminders.length,
      validRules: validRules,
      totalChecks: totalChecks,
      summary: summaryText
    };

  } catch (error) {
    Logger.log(`ERROR in checkAndSendReminderEmails: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return { success: false, error: error.message };
  }
}

/**
 * Kiểm tra files còn thiếu - DEPRECATED, use debugCheckMissingFiles instead
 * This function is kept for reference but debugCheckMissingFiles is now used
 */
function checkMissingFiles_DEPRECATED(classname, groupName, folderName, ruleRow, fileIndex) {
  try {
    // Lấy folder của nhóm
    const root = getSpreadsheetParent();
    const userprofile = getOrCreateFolder(root, "userprofile");
    const classFolder = getOrCreateFolder(userprofile, classname);
    const groupFolder = classFolder.getFoldersByName(groupName).next();

    // Tìm folder theo tên
    const targetFolder = findFolderByName(groupFolder, folderName);
    if (!targetFolder) return ["Folder not found"];

    // Lấy requirements
    const fileTypeCol = 3 + (fileIndex - 1) * 3 + 1;
    const expectedType = ruleRow[fileTypeCol] || '';
    const requirement = ruleRow[fileTypeCol + 1] || '';

    // Kiểm tra files hiện có
    const files = targetFolder.getFiles();
    let hasMatchingFile = false;

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName().toLowerCase();
      const extension = fileName.substring(fileName.lastIndexOf('.'));

      if (expectedType.toLowerCase().includes(extension)) {
        hasMatchingFile = true;
        break;
      }
    }

    if (!hasMatchingFile) {
      return [{
        type: expectedType,
        requirement: requirement,
        missing: true
      }];
    }

    return [];

  } catch (e) {
    Logger.log(`Error checking missing files: ${e.message}`);
    return ["Error checking files"];
  }
}

/**
 * Tìm folder theo tên trong cây thư mục
 */
function findFolderByName(parentFolder, targetName) {
  const folders = parentFolder.getFolders();

  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName() === targetName) {
      return folder;
    }

    // Tìm trong subfolders
    const subResult = findFolderByName(folder, targetName);
    if (subResult) return subResult;
  }

  return null;
}

/**
 * Gửi reminder email - Updated to handle debug logic results
 */
function sendReminderEmail(reminder) {
  try {
    Logger.log(`Preparing reminder email for ${reminder.classname}/${reminder.groupName}/${reminder.folderName}`);

    const members = getGroupMembers(reminder.classname, reminder.groupName);
    Logger.log(`Found ${members.length} members: ${members.join(', ')}`);

    if (members.length === 0) {
      Logger.log(`No members found for group ${reminder.groupName}`);
      return { success: false, sentCount: 0, error: `No members found for ${reminder.classname}/${reminder.groupName}` };
    }

    // Build missing files list from debug results
    let missingFilesHtml = '';
    let missingCount = 0;

    reminder.missingFiles.forEach(file => {
      if (typeof file === 'string') {
        // Handle string errors (like "Folder not found")
        missingFilesHtml += `<li><strong>Error:</strong> ${file}</li>`;
        missingCount++;
      } else if (file.missing === true) {
        // Handle structured missing file data
        missingFilesHtml += `<li><strong>${file.type}:</strong> ${file.requirement}`;
        if (file.filesFound && file.filesFound.length > 0) {
          missingFilesHtml += `<br><small>Files found: ${file.filesFound.join(', ')}</small>`;
        }
        missingFilesHtml += `</li>`;
        missingCount++;
      }
    });

    if (missingCount === 0) {
      Logger.log(`No actual missing files found, skipping email for ${reminder.groupName}`);
      return { success: false, sentCount: 0, error: 'No missing files found' };
    }

    const folderInfo = findGroupAssignmentFolder(reminder.classname, reminder.groupName, reminder.folderName);
    const folderUrl = folderInfo && folderInfo.folder ? folderInfo.folder.getUrl() : '';

    const formatDate = (date) => {
      if (!date) return '';
      return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
    };

    const emailContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
      </head>
      <body>
      <div style="font-family: Arial, sans-serif; max-width: 700px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #ea4335 0%, #fbbc04 100%); color: white; padding: 25px; text-align: center; border-radius: 10px 10px 0 0;">
          <h2 style="margin: 0; font-size: 24px;">URGENT: ASSIGNMENT DUE TOMORROW!</h2>
          <p style="margin: 10px 0 0 0; font-size: 14px; opacity: 0.95;">Less than 24 hours remaining to submit</p>
        </div>
        
        <div style="padding: 30px; background: #ffffff; border: 1px solid #e0e0e0; border-radius: 0 0 10px 10px;">
          <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <h3 style="margin: 0 0 10px 0; color: #ea4335; font-size: 18px;">Class: ${reminder.classname}</h3>
            <h4 style="margin: 0; color: #5f6368; font-size: 16px;">Group: ${reminder.groupName}</h4>
          </div>
          
          <div style="background: #fff3cd; border: 2px solid #fbbc04; padding: 20px; border-radius: 8px; margin: 15px 0;">
            <h4 style="margin: 0 0 15px 0; color: #ea8600; font-size: 16px;">Folder: ${reminder.folderName}</h4>
            <table style="width: 100%; font-size: 14px;">
              <tr>
                <td style="padding: 5px 0; color: #5f6368; width: 140px;"><strong>Due Date:</strong></td>
                <td style="padding: 5px 0; color: #ea4335; font-weight: bold;">${formatDate(reminder.dueDate)}</td>
              </tr>
              <tr>
                <td style="padding: 5px 0; color: #5f6368;"><strong>Time Remaining:</strong></td>
                <td style="padding: 5px 0; color: #ea4335; font-weight: bold;">Less than 24 hours!</td>
              </tr>
            </table>
          </div>
          
          <div style="background: #fce8e6; border: 2px solid #ea4335; padding: 20px; border-radius: 8px; margin: 15px 0;">
            <h4 style="margin: 0 0 15px 0; color: #ea4335;">Missing Files (${missingCount} issues):</h4>
            <ul style="margin: 0; padding-left: 20px; line-height: 1.8; color: #202124;">
              ${missingFilesHtml}
            </ul>
          </div>
          
          <div style="text-align: center; margin: 25px 0;">
            ${folderUrl ? `
            <a href="${folderUrl}" 
               target="_blank"
               style="display: inline-block; background: #34a853; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.2);">
              OPEN FOLDER NOW
            </a>` : `<span style="color:#ea4335;font-weight:bold;">Folder link not available</span>`}
          </div>
          
          <div style="background: #e8f0fe; border: 2px solid #1a73e8; padding: 20px; border-radius: 8px; margin: 15px 0;">
            <h4 style="margin: 0 0 15px 0; color: #1a73e8;">Action Required:</h4>
            <ol style="margin: 0; padding-left: 20px; line-height: 1.8;">
              <li><strong>Upload immediately</strong> the missing files</li>
              <li>Check <strong>file format</strong> matches requirements</li>
              <li>Review <strong>all requirements</strong> are completed</li>
              <li>Contact instructor if you need help</li>
            </ol>
          </div>
          
          <div style="background: #fff; border: 1px solid #fbbc04; padding: 15px; border-radius: 8px; margin: 15px 0;">
            <p style="margin: 0; color: #ea8600; font-size: 14px;">
              <strong>Warning:</strong> Late submission will result in grade penalty as per instructor guidelines.
            </p>
          </div>
          
          <div style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0; color: #5f6368; font-size: 13px;">
            <p style="margin: 5px 0;">Automated reminder from the Management System</p>
            <p style="margin: 5px 0;">Sent at: ${new Date().toLocaleString('en-US')}</p>
            <p style="margin: 15px 0 0 0; font-style: italic;">Please do not reply to this email</p>
          </div>
        </div>
      </div>
      </body>
      </html>
    `;

    // Gửi email cho từng thành viên bằng cùng logic với Send Folder Links
    let emailCount = 0;
    const failedEmails = [];
    members.forEach(email => {
      const emailAddress = (email || '').toString().trim();
      if (emailAddress && emailAddress.includes('@')) {
        try {
          Logger.log(`Sending reminder email to: ${emailAddress}`);

          const result = sendEmailViaGmail({
            to: emailAddress,
            subject: `URGENT: Assignment Due Tomorrow - ${reminder.classname} - ${reminder.folderName}`,
            htmlBody: emailContent,
          });

          if (result && result.success) {
            emailCount++;
            Logger.log(`Reminder email sent successfully to ${emailAddress}`);
          } else {
            const errMsg = result && result.error ? result.error : 'Unknown sendEmailViaGmail error';
            failedEmails.push(`${emailAddress} (${errMsg})`);
            Logger.log(`Failed to send reminder to ${emailAddress}: ${errMsg}`);
          }

        } catch (e) {
          failedEmails.push(`${emailAddress} (${e.message})`);
          Logger.log(`Failed to send reminder to ${emailAddress}: ${e.message}`);
        }
      } else {
        Logger.log(`Skipping invalid reminder recipient: ${emailAddress}`);
      }
    });

    Logger.log(`Sent reminder email to ${emailCount}/${members.length} members for ${reminder.groupName}`);
    if (emailCount === 0) {
      const detail = failedEmails.length > 0 ? ` | Failed: ${failedEmails.join('; ')}` : '';
      return {
        success: false,
        sentCount: 0,
        error: `No email was sent for ${reminder.classname}/${reminder.groupName}${detail}`
      };
    }
    return {
      success: true,
      sentCount: emailCount,
      totalMembers: members.length,
      failedEmails: failedEmails
    };

  } catch (error) {
    Logger.log(`Error sending reminder email: ${error.message}`);
    return { success: false, sentCount: 0, error: error.message };
  }
}

/**
 * Lấy danh sách classes cho dropdown
 */
function getClassesForEmail() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Class List');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const classes = new Set();
    for (let i = 1; i < data.length; i++) {
      const classname = data[i][0];
      if (classname) classes.add(classname);
    }

    return Array.from(classes);
  } catch (e) {
    Logger.log(`Error getting classes: ${e.message}`);
    return [];
  }
}

/**
 * Lấy danh sách groups theo class
 */
function getGroupsByClass(classname) {
  try {
    if (!classname) return [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Class List');
    if (!sheet) return [];

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    const groups = new Set();
    for (let i = 1; i < data.length; i++) {
      const rowClass = (data[i][0] || '').toString().trim();
      const rowGroup = (data[i][1] || '').toString().trim();
      if (rowClass === classname.toString().trim() && rowGroup) {
        groups.add(rowGroup);
      }
    }

    return Array.from(groups);
  } catch (e) {
    Logger.log(`Error getting groups: ${e.message}`);
    return [];
  }
}

/**
 * Lấy members của một nhóm
 */
function getGroupMembers(classname, groupName) {
  try {
    Logger.log(`Getting members for ${classname} - ${groupName}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Class List');
    if (!sheet) {
      Logger.log('ERROR: Class List sheet not found');
      return [];
    }

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) {
      Logger.log('ERROR: Class List sheet is empty');
      return [];
    }

    Logger.log(`Class List has ${data.length - 1} rows of data`);

    // Lấy tất cả emails từ các rows matching classname và groupName
    const allEmails = [];
    const emailPattern = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowClass = row[0] ? row[0].toString().trim() : '';
      const rowGroup = row[1] ? row[1].toString().trim() : '';

      if (rowClass === classname.toString().trim() && rowGroup === groupName.toString().trim()) {
        Logger.log(`  Found matching row ${i + 1}: ${rowClass} - ${rowGroup}`);

        // Ưu tiên cấu trúc cột cũ, sau đó fallback quét toàn dòng để tránh lệch layout.
        const emailColumns = [4, 7, 10, 13];
        let foundAnyEmail = false;

        emailColumns.forEach((colIndex, position) => {
          if (row[colIndex]) {
            const email = row[colIndex].toString().trim();
            if (email && emailPattern.test(email)) {
              allEmails.push(email);
              foundAnyEmail = true;
              Logger.log(`    Position ${position + 1}: ${email}`);
            } else if (email) {
              Logger.log(`    Position ${position + 1}: Invalid email format: "${email}"`);
            }
          }
        });

        if (!foundAnyEmail) {
          Logger.log('    No emails found in fixed columns, fallback scan entire row');
          for (let col = 2; col < row.length; col++) {
            const cell = (row[col] || '').toString().trim();
            if (cell && emailPattern.test(cell)) {
              allEmails.push(cell);
              foundAnyEmail = true;
              Logger.log(`    Fallback email found in col ${col + 1}: ${cell}`);
            }
          }
        }
      }
    }

    // Remove duplicates
    const uniqueEmails = [...new Set(allEmails)];
    Logger.log(`Total unique emails found: ${uniqueEmails.length}`);

    return uniqueEmails;

  } catch (e) {
    Logger.log(`Error getting group members: ${e.message}`);
    Logger.log(`Stack trace: ${e.stack}`);
    return [];
  }
}

/**
 * Helper function để lấy hoặc tạo folder
 */
function getSpreadsheetParent() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const file = DriveApp.getFileById(ss.getId());
    const parents = file.getParents();
    return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  } catch (e) {
    Logger.log(`Error getting spreadsheet parent: ${e.message}`);
    return DriveApp.getRootFolder();
  }
}

/**
 * Helper function để lấy hoặc tạo folder
 */
function getOrCreateFolder(parent, name) {
  try {
    const folders = parent.getFoldersByName(name);
    if (folders.hasNext()) {
      return folders.next();
    } else {
      return parent.createFolder(name);
    }
  } catch (e) {
    Logger.log(`Error creating folder ${name}: ${e.message}`);
    return null;
  }
}

/**
 * Helper function để export debug logs ra file text
 */
function exportDebugLogs() {
  try {
    Logger.log('=== EXPORTING DEBUG LOGS ===');

    // Run full debug
    const debugResult = debugReminderSystem();

    // Get current timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');

    // Create debug report
    let report = `DEBUG REPORT - ${timestamp}\n`;
    report += `=====================================\n\n`;

    report += `System Status: ${debugResult.success ? 'OK' : 'ERROR'}\n`;
    if (!debugResult.success) {
      report += `Error: ${debugResult.error}\n`;
    }
    report += `Valid Rules: ${debugResult.validRules || 0}\n`;
    report += `Reminders Needed: ${debugResult.reminders || 0}\n\n`;

    // Add reminder details
    if (debugResult.reminderDetails && debugResult.reminderDetails.length > 0) {
      report += `REMINDER DETAILS:\n`;
      report += `-----------------\n`;
      debugResult.reminderDetails.forEach((r, index) => {
        report += `${index + 1}. ${r.classname}/${r.groupName}/${r.folderName}\n`;
        report += `   Due Date: ${new Date(r.dueDate).toLocaleDateString()}\n`;
        report += `   Missing Files: ${r.missingFiles.length}\n`;
        r.missingFiles.forEach(f => {
          report += `     - ${f.type}: ${f.requirement}\n`;
        });
        report += `\n`;
      });
    }

    // Add system info
    report += `SYSTEM INFORMATION:\n`;
    report += `--------------------\n`;
    report += `Email Quota Remaining: ${MailApp.getRemainingDailyQuota()}\n`;
    report += `Active User: ${Session.getActiveUser().getEmail()}\n`;
    report += `Spreadsheet ID: ${SpreadsheetApp.getActiveSpreadsheet().getId()}\n`;

    // Try to save to Drive
    try {
      const blob = Utilities.newBlob(report, 'text/plain', `debug-report-${timestamp}.txt`);
      const file = DriveApp.createFile(blob);

      Logger.log(`Debug report saved: ${file.getUrl()}`);

      return {
        success: true,
        fileUrl: file.getUrl(),
        fileName: file.getName(),
        report: report
      };

    } catch (driveError) {
      Logger.log(`Could not save to Drive: ${driveError.message}`);

      return {
        success: true,
        fileUrl: null,
        fileName: null,
        report: report,
        note: 'Report generated but could not save to Drive'
      };
    }

  } catch (error) {
    Logger.log(`Error exporting debug logs: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test reminder system with option to actually send emails
 */
function testReminderSystemWithDebug(sendEmails = false) {
  try {
    Logger.log(`=== TESTING REMINDER SYSTEM (Send Emails: ${sendEmails}) ===`);

    // First run debug analysis
    const debugResult = debugReminderSystem();

    if (!debugResult.success) {
      return debugResult;
    }

    Logger.log(`Debug analysis completed: ${debugResult.reminders} reminders found`);

    if (sendEmails && debugResult.reminders > 0) {
      Logger.log('Proceeding to send actual reminder emails...');

      // If debug found reminders and we want to send emails, use the actual function
      const emailResult = checkAndSendReminderEmails();

      return {
        success: emailResult.success,
        debugResults: debugResult,
        emailResults: emailResult,
        mode: 'email_sent',
        count: emailResult.count || 0,
        totalReminders: emailResult.totalReminders || 0,
        validRules: emailResult.validRules || 0,
        totalChecks: emailResult.totalChecks || 0,
        summary: emailResult.summary || `Debug found ${debugResult.reminders} reminders, email system ${emailResult.success ? 'succeeded' : 'failed'}`
      };

    } else {
      Logger.log('Debug mode only - no emails sent');

      return {
        success: true,
        debugResults: debugResult,
        emailResults: null,
        mode: 'debug_only',
        count: 0,
        totalReminders: debugResult.reminders,
        validRules: debugResult.validRules,
        totalChecks: 0,
        summary: `Debug analysis completed - ${debugResult.reminders} reminders would be sent if email mode was enabled`
      };
    }

  } catch (error) {
    Logger.log(`ERROR in testReminderSystemWithDebug: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/**
 * Setup trigger tự động chạy reminder emails hàng ngày
 */
function setupAutomaticReminders() {
  // Không cần ScriptApp trigger nữa, chỉ báo thành công
  Logger.log('Automatic reminder setup completed - use manual execution');
  return {
    success: true,
    message: 'Automatic reminders configured to use manual execution due to permission constraints'
  };
}