function showFolderCreatorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('folderCreationUI') // ref: folderCreationUI
    .setTitle('Folder Creator');
  SpreadsheetApp.getUi().showSidebar(html);
}
function writeClassTree(classname, obj) {
  if (!obj || !Array.isArray(obj) || obj.length === 0) {
    Logger.log("writeClassTree: empty obj, nothing to write.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Create Folder');
  if (!sheet) sheet = ss.insertSheet('Create Folder');

  // ====== TÍNH MAX DEPTH ======
  let maxDepth = 0;
  (function getMaxDepth(nodes, depth = 1) {
    nodes.forEach(node => {
      if (!node) return;
      maxDepth = Math.max(maxDepth, depth);
      if (node.children && node.children.length > 0)
        getMaxDepth(node.children, depth + 1);
    });
  })(obj);

  // ====== HEADER ======
  const header = ["Class name", "Group name"];
  for (let i = 1; i <= maxDepth; i++) header.push(`LEVEL ${i}`);

  // đảm bảo header tồn tại và có đủ cột
  const lastCol = Math.max(sheet.getLastColumn(), header.length);
  const existingHeader = (sheet.getLastRow() >= 1)
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    : [];

  if (existingHeader.length < header.length || sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length)
      .setValues([header])
      .setBackground("#d9ead3")
      .setFontWeight("bold");
  }

  const startRow = sheet.getLastRow() + 1;
  const rows = [];

  // ====== GHI CÂY DỮ LIỆU ======
  let writtenAny = false;
  function traverse(nodes, level = 1) {
    if (!nodes || nodes.length === 0) return;
    nodes.forEach(node => {
      if (!node || !node.name || node.name.toString().trim() === "") {
        // skip nodes without a valid name
        if (node && node.children && node.children.length > 0) {
          // still traverse children in case they have names
          traverse(node.children, level + 1);
        }
        return;
      }

      const row = Array(maxDepth + 2).fill("");
      // write classname only once (first written row)
      if (!writtenAny) {
        row[0] = classname;
        row[1] = "{}";
        writtenAny = true;
      }

      row[1 + level] = node.name.toString().trim();
      rows.push(row);

      if (node.children && node.children.length > 0) {
        traverse(node.children, level + 1);
      }
    });
  }

  traverse(obj);

  if (!rows.length) {
    Logger.log("writeClassTree: no valid rows after traversal, skipping sheet write.");
    return;
  }

  // ====== GHI DỮ LIỆU VÀO SHEET (an toàn) ======
  const writeRange = sheet.getRange(startRow, 1, rows.length, maxDepth + 2);
  // Force text format so values like "4.1" are not auto-converted to dates.
  writeRange.setNumberFormat('@');
  writeRange.setValues(rows);
  sheet.autoResizeColumns(1, maxDepth + 2);

  // ====== THU THẬP TẤT CẢ NODE (kể cả node cha) an toàn ======
  const allNodes = new Set();
  (function collectAll(nodes) {
    nodes.forEach(node => {
      if (!node || !node.name) return;
      const n = node.name.toString().trim();
      if (n) allNodes.add(n);
      if (node.children && node.children.length > 0) collectAll(node.children);
    });
  })(obj);

  const lessons = Array.from(allNodes);
  Logger.log("📘 Node lá được ghi sang Rules: " + JSON.stringify(lessons));

  if (lessons.length === 0) {
    Logger.log("⚠️ Không có node lá nào để ghi sang Rules.");
    return;
  }

  // ====== GHI SANG SHEET RULES (chỉ khi có rows) ======
  let ruleSheet = ss.getSheetByName("Rules");
  if (!ruleSheet) {
    ruleSheet = ss.insertSheet("Rules");
    ruleSheet.getRange(1, 1, 1, 2).setValues([['Class name', 'Folder']]);
    ruleSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    ruleSheet.getRange(1, 1, 1, 2).setBackground('#d4e6f1');
    ruleSheet.autoResizeColumns(1, 2);
  }

  const existing = (ruleSheet.getLastRow() > 1)
    ? ruleSheet.getRange(2, 1, ruleSheet.getLastRow() - 1, 2).getValues()
    : [];
  const existingSet = new Set(existing.map(r => (r[0] || "") + "|" + (r[1] || "")));

  const newRows = [];
  lessons.forEach(name => {
    const key = classname + "|" + name;
    if (!existingSet.has(key)) newRows.push([classname, name]);
  });

  if (newRows.length > 0) {
    const start = ruleSheet.getLastRow() + 1;
    const rulesRange = ruleSheet.getRange(start, 1, newRows.length, 2);
    // Keep Class/Folder columns as text to avoid values like "4.1" becoming dates.
    rulesRange.setNumberFormat('@');
    rulesRange.setValues(newRows);
    ruleSheet.autoResizeColumns(1, 2);
    Logger.log(`✅ Đã thêm ${newRows.length} node lá vào Rules`);
  } else {
    Logger.log("ℹ️ Không có node mới nào để thêm vào Rules");
  }
}


/**
 * Lấy danh sách group name theo class name
 */
function getGroupNamesByClass(classname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = ss.getSheetByName('Class List');
  if (!classSheet) return [];

  const values = classSheet.getDataRange().getValues();
  if (values.length < 2) return [];

  // Chuẩn hóa header
  const header = values[0].map(h => (h || "").toString().trim().toLowerCase());
  const classIdx = header.indexOf("class name");
  const groupIdx = header.indexOf("group name");

  if (classIdx === -1 || groupIdx === -1) {
    Logger.log("❌ Không tìm thấy cột 'Class Name' hoặc 'Group Name' trong sheet 'Class List'");
    return [];
  }

  // Lọc theo class name
  const groups = values.slice(1)
    .filter(r => (r[classIdx] || "").toString().trim() === classname)
    .map(r => (r[groupIdx] || "").toString().trim())
    .filter(g => g !== "");

  // Loại bỏ trùng lặp
  const uniqueGroups = [...new Set(groups)];

  Logger.log(`✅ Class ${classname} có ${uniqueGroups.length} group: ${uniqueGroups.join(", ")}`);

  return uniqueGroups;
}




function _syncCreateFolderToRuleRange(sourceSheet, startRow, numRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ruleSheet = ss.getSheetByName("Rules");

  // Nếu chưa có Rules thì tạo mới với 2 cột (không có NUMBER OF FILES)
  if (!ruleSheet) {
    ruleSheet = ss.insertSheet("Rules");
    ruleSheet.getRange(1, 1, 1, 2).setValues([['Class name', 'Folder']]);
    ruleSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    ruleSheet.getRange(1, 1, 1, 2).setBackground('#d4e6f1');
    ruleSheet.autoResizeColumns(1, 2);
    SpreadsheetApp.getUi().alert('Đã tạo sheet Rules mới với 2 cột cơ bản.');

  }

  // Kiểm tra header (không phân biệt hoa thường)
  const ruleHeader = (ruleSheet.getLastRow() >= 1 && ruleSheet.getLastColumn() >= 1)
    ? ruleSheet.getRange(1, 1, 1, ruleSheet.getLastColumn())
      .getValues()[0]
      .map(h => (h || '').toString().toLowerCase())
    : [];

  const classCol = ruleHeader.indexOf('classname');
  const folderCol = ruleHeader.indexOf('folder');

  if (classCol === -1 || folderCol === -1) {
    throw new Error("Sheet Rules phải có cột 'Class name' và 'Folder'");
  }

  // lấy dữ liệu mới từ Create Folder
  const range = sourceSheet.getRange(startRow, 1, numRows, sourceSheet.getLastColumn());
  const values = range.getValues();

  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const classIdx = headers.map(h => (h || '').toString().toLowerCase()).indexOf('Class name');

  const result = [];
  let lastClass = "";

  values.forEach(row => {
    const classname = row[classIdx] || lastClass;
    if (classname) lastClass = classname;

    for (let j = 2; j < row.length; j++) { // LEVEL bắt đầu từ cột index 2
      const folder = row[j];
      if (folder) {
        result.push([classname, folder]); // ❌ không có NUMBER OF FILES nữa
      }
    }
  });

  if (result.length > 0) {
    ruleSheet.getRange(ruleSheet.getLastRow() + 1, 1, result.length, 2).setValues(result);
  }
}
/**
 * Kiểm tra xem class đã có folder chưa
 * Trả về object: { hasFolder: boolean, message: string }
 */
function checkClassHasFolder(classname) {
  if (!classname || classname.trim() === '') {
    return { hasFolder: false, message: 'Class name không hợp lệ' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Kiểm tra trong sheet "Create Folder"
  const createFolderSheet = ss.getSheetByName('Create Folder');
  if (createFolderSheet && createFolderSheet.getLastRow() > 1) {
    const data = createFolderSheet.getDataRange().getValues();
    const header = data[0];
    const classnameColIdx = header.map(h => (h || '').toString().toLowerCase()).indexOf('class name');

    if (classnameColIdx !== -1) {
      for (let i = 1; i < data.length; i++) {
        const rowClass = (data[i][classnameColIdx] || '').toString().trim();
        if (rowClass === classname) {
          return {
            hasFolder: true,
            message: `Class "${classname}" đã có cấu trúc folder trong sheet "Create Folder". Vui lòng sử dụng "Change Structure" để chỉnh sửa.`
          };
        }
      }
    }
  }

  // Kiểm tra trên Drive
  try {
    const parentFolder = getSpreadsheetParent();
    const tempFolder = parentFolder.getFoldersByName("temp");

    if (tempFolder.hasNext()) {
      const temp = tempFolder.next();
      const classFolder = temp.getFoldersByName(classname);

      if (classFolder.hasNext()) {
        return {
          hasFolder: true,
          message: `Class "${classname}" đã có folder trên Drive. Vui lòng sử dụng "Change Structure" để chỉnh sửa.`
        };
      }
    }
  } catch (e) {
    Logger.log("Error checking Drive folder: " + e.message);
  }

  return { hasFolder: false, message: '' };
}

function getAllClasses(sheetName = 'Class List') {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetName);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  const cleanedData = data.map(row => row[0]).filter(val => val && val.toString().trim() !== '');

  const distinctSet = new Set(cleanedData);

  const res = Array.from(distinctSet);

  res.forEach(item => Logger.log(item));

  return res;
}

function getAllGroupOfClass(classname) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Class List');

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  const cleanedData = data
    .filter(row => row[0] === classname)
    .map(row => row[1])
    .filter(val => val);

  const distinctSet = new Set(cleanedData);
  const res = Array.from(distinctSet);
  return res;
}

function getGroupMembers(classname = 'COMP1314', groupname = 'Nhóm kia tào lao á') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class List');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  const emailsGroup = data
    .filter(row => row[0] == classname && row[1] == groupname)
    .map(row => [
      row[4],  // leader
      row[7],  // member 1
      row[10], // member 2
      row[13]  // member 3
    ])
    .flat(); // gộp thành 1 array thay vì [[..]]

  const cleaned = emailsGroup.filter(e => e && e.toString().trim() !== "");

  Logger.log(`${classname} - ${groupname}: ${JSON.stringify(cleaned)}`);
  Logger.log(cleaned)

  return cleaned;
}


function createClassDrive(inputClassname, obj) {

  const classname = inputClassname;
  const data = obj

  const groups = getAllGroupOfClass(classname);

  const parentFolder = getSpreadsheetParent();

  // === Step 1: temp/classA/_template ===
  let tempFolder = getOrCreateFolder(parentFolder, "temp");
  let classFolder = getOrCreateFolder(tempFolder, classname);
  let templateFolder = getOrCreateFolder(classFolder, "_template");

  // Build template once if empty
  if (!templateFolder.getFolders().hasNext() && !templateFolder.getFiles().hasNext()) {
    createFolders(data, templateFolder);
  }

  // === Step 2: userprofile/groups ===
  const failedGroups = [];
  groups.forEach(group => {
    try {
      let ids = addGroupToClass(classname, group);
      Logger.log(JSON.stringify(ids, null, 2)); // log ID tree of group
    } catch (e) {
      failedGroups.push({ group: group, error: e.message });
      Logger.log(`❌ addGroupToClass failed for ${classname}/${group}: ${e.message}`);
    }
  });

  if (failedGroups.length > 0) {
    Logger.log(`⚠️ Some groups failed while generating class ${classname}: ${JSON.stringify(failedGroups)}`);
  }
}

/**
 * Add a new group into an existing class
 * - classname: e.g. "classA"
 * - groupName: e.g. "group05"
 */
function addGroupToClass(classname, groupName) {
  const parentFolder = getSpreadsheetParent();

  // === Find template: temp/classA/_template ===
  let tempFolder = getOrCreateFolder(parentFolder, "temp");
  let classFolder = getOrCreateFolder(tempFolder, classname);
  let templateFolders = classFolder.getFoldersByName("_template");
  if (!templateFolders.hasNext()) {
    throw new Error("No _template folder found for " + classname);
  }
  let templateFolder = templateFolders.next();

  // === userprofile/groupName ===
  let userProfileFolder = getOrCreateFolder(parentFolder, "userprofile");
  let classProfileFolder = getOrCreateFolder(userProfileFolder, classname)
  let groupFolder = getOrCreateFolder(classProfileFolder, groupName);

  // If empty, copy template into it
  if (!groupFolder.getFolders().hasNext() && !groupFolder.getFiles().hasNext()) {
    copyContents(templateFolder, groupFolder);
    Logger.log("Created new group: " + groupName);
  } else {
    Logger.log("Group " + groupName + " already exists, skipped copy");
  }

  // === Apply permissions for this group ===
  const members = getGroupMembers(classname, groupName);
  try {
    applyGroupPermissions(groupFolder, members);
  } catch (e) {
    // Không chặn luồng tạo folder nếu có lỗi quyền ở một group.
    Logger.log(`Permission apply failed for ${classname}/${groupName}: ${e.message}`);
  }

  // Return ID tree
  const ids = collectFolderIds(groupFolder)
  Logger.log(ids)
  return ids;
}
/**
 * === PERMISSIONS HANDLING ===
 * Sử dụng DriveApp (gửi email) thay vì Drive API để tránh lỗi
 */
function driveWithRetry(operation, label, maxAttempts = 3) {
  let lastError = null;
  for (let i = 1; i <= maxAttempts; i++) {
    try {
      return operation();
    } catch (e) {
      lastError = e;
      Logger.log(`Drive attempt ${i}/${maxAttempts} failed [${label}]: ${e.message}`);
      if (i < maxAttempts) {
        Utilities.sleep(i * 250);
      }
    }
  }
  throw lastError;
}

function applyGroupPermissions(folder, groupMembers) {
  // Không reset sharing ở đây để tránh lỗi Drive service trên một số môi trường Drive/Shared Drive.

  // 1. Remove old permissions (không dùng getFileById cho folder)
  let editors = [];
  try {
    editors = driveWithRetry(() => folder.getEditors(), `getEditors:${folder.getId()}`) || [];
  } catch (e) {
    Logger.log("Cannot list editors for " + folder.getName() + ": " + e.message);
  }

  editors.forEach(user => {
    try {
      driveWithRetry(() => {
        folder.removeEditor(user);
        return true;
      }, `removeEditor:${folder.getId()}:${user.getEmail()}`);
    } catch (e) {
      Logger.log("Cannot remove editor " + user.getEmail() + ": " + e.message);
    }
  });

  // 2. Remove all viewers
  let viewers = [];
  try {
    viewers = driveWithRetry(() => folder.getViewers(), `getViewers:${folder.getId()}`) || [];
  } catch (e) {
    Logger.log("Cannot list viewers for " + folder.getName() + ": " + e.message);
  }

  viewers.forEach(user => {
    try {
      driveWithRetry(() => {
        folder.removeViewer(user);
        return true;
      }, `removeViewer:${folder.getId()}:${user.getEmail()}`);
    } catch (e) {
      Logger.log("Cannot remove viewer " + user.getEmail() + ": " + e.message);
    }
  });

  // 3. Grant fresh permissions - mặc định viewer
  groupMembers.forEach(email => {
    if (email) {
      try {
        driveWithRetry(() => {
          folder.addViewer(email);
          return true;
        }, `addViewer:${folder.getId()}:${email}`);
      } catch (e) {
        Logger.log("Failed to add " + email + ": " + e.message);
      }
    }
  });

  // 4. Recurse for subfolders
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const child = subfolders.next();
    try {
      applyGroupPermissions(child, groupMembers);
    } catch (e) {
      Logger.log(`Permission recursion failed at ${child.getName()}: ${e.message}`);
    }
  }
}

/**
 * Remove all editors and viewers except the owner
 */
function removeNonOwnerPermissions(folderId) {
  const folder = driveWithRetry(() => DriveApp.getFolderById(folderId), `getFolderById:${folderId}`);

  let ownerEmail = "";
  try {
    const owner = driveWithRetry(() => folder.getOwner(), `getOwner:${folderId}`);
    ownerEmail = owner ? owner.getEmail() : "";
  } catch (e) {
    Logger.log(`Cannot read owner for ${folderId}: ${e.message}`);
  }

  let editors = [];
  try {
    editors = driveWithRetry(() => folder.getEditors(), `getEditors:${folderId}`) || [];
  } catch (e) {
    Logger.log(`Cannot list editors for ${folderId}: ${e.message}`);
  }

  editors.forEach(user => {
    const email = user.getEmail();
    if (email && email !== ownerEmail) {
      try {
        driveWithRetry(() => {
          folder.removeEditor(email);
          return true;
        }, `removeEditor:${folderId}:${email}`);
      } catch (e) {
        Logger.log("Cannot remove editor " + email + ": " + e.message);
      }
    }
  });

  let viewers = [];
  try {
    viewers = driveWithRetry(() => folder.getViewers(), `getViewers:${folderId}`) || [];
  } catch (e) {
    Logger.log(`Cannot list viewers for ${folderId}: ${e.message}`);
  }

  viewers.forEach(user => {
    const email = user.getEmail();
    if (email && email !== ownerEmail) {
      try {
        driveWithRetry(() => {
          folder.removeViewer(email);
          return true;
        }, `removeViewer:${folderId}:${email}`);
      } catch (e) {
        Logger.log("Cannot remove viewer " + email + ": " + e.message);
      }
    }
  });
}

/**
 * Add editor without sending email
 */
function addEditorNoEmail(fileId, email) {
  const folder = driveWithRetry(() => DriveApp.getFolderById(fileId), `getFolderById:${fileId}`);
  driveWithRetry(() => {
    folder.addEditor(email);
    return true;
  }, `addEditor:${fileId}:${email}`);
}

/**
 * Add viewer without sending email
 */
function addViewerNoEmail(fileId, email) {
  const folder = driveWithRetry(() => DriveApp.getFolderById(fileId), `getFolderById:${fileId}`);
  driveWithRetry(() => {
    folder.addViewer(email);
    return true;
  }, `addViewer:${fileId}:${email}`);
}

/**
 * Update permissions for a specific folder
 */
function updateDrivePermissions(folderId, emails, permission) {
  try {
    if (!folderId || !emails || emails.length === 0) return;

    const folderRole = normalizePermissionValue(permission);
    if (!folderRole) {
      throw new Error(`Permission không hợp lệ: ${permission}`);
    }

    const token = ScriptApp.getOAuthToken();
    const currentPermissions = listFolderPermissionsByApi(folderId, token);
    const ownerEmails = new Set();

    currentPermissions.forEach(p => {
      if (p.type === 'user' && p.role === 'owner' && p.emailAddress) {
        ownerEmails.add(p.emailAddress.toLowerCase());
      }
    });

    const wantedEmails = new Set(
      emails
        .map(email => (email || '').toString().trim().toLowerCase())
        .filter(email => email)
    );

    currentPermissions.forEach(p => {
      if (p.type !== 'user' || !p.emailAddress || !p.id) return;
      const email = p.emailAddress.toLowerCase();
      if (ownerEmails.has(email)) return;
      if (!wantedEmails.has(email)) {
        deletePermissionByApi(folderId, p.id, token);
      }
    });

    wantedEmails.forEach(email => {
      const existing = currentPermissions.filter(p => p.type === 'user' && p.emailAddress && p.emailAddress.toLowerCase() === email);
      existing.forEach(p => {
        if (p.role !== 'owner') {
          deletePermissionByApi(folderId, p.id, token);
        }
      });
      createPermissionByApi(folderId, email, folderRole, token);
    });

    Logger.log(`Updated permissions for folder ${folderId}: ${Array.from(wantedEmails).join(', ')} as ${folderRole}`);
  } catch (error) {
    Logger.log(`Error updating drive permissions: ${error.message}`);
    throw error;
  }
}

function listFolderPermissionsByApi(folderId, token) {
  const url = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(folderId)}/permissions?supportsAllDrives=true&fields=permissions(id,emailAddress,role,type)`;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error(`List permissions failed (${response.getResponseCode()}): ${response.getContentText()}`);
  }

  const body = JSON.parse(response.getContentText() || '{}');
  return body.permissions || [];
}

function deletePermissionByApi(folderId, permissionId, token) {
  const url = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(folderId)}/permissions/${encodeURIComponent(permissionId)}?supportsAllDrives=true`;
  const response = UrlFetchApp.fetch(url, {
    method: 'delete',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error(`Delete permission failed (${response.getResponseCode()}): ${response.getContentText()}`);
  }
}

function createPermissionByApi(folderId, email, permission, token) {
  const url = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(folderId)}/permissions?supportsAllDrives=true&sendNotificationEmail=false`;
  const role = permission === 'editor' ? 'writer' : permission === 'viewer' ? 'reader' : permission;
  const payload = {
    type: 'user',
    role: role,
    emailAddress: email
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error(`Create permission failed (${response.getResponseCode()}): ${response.getContentText()}`);
  }
}

function clearDirectPermissionForEmail(folderId, email) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    try {
      folder.removeEditor(email);
    } catch (e) {
      // Ignore if not an editor.
    }

    try {
      folder.removeViewer(email);
    } catch (e) {
      // Ignore if not a viewer.
    }

    // commenter is only available via Advanced Drive Service; keep safe fallback.
    if (typeof Drive !== 'undefined' && Drive.Permissions && Drive.Permissions.list) {
      try {
        const permissionList = Drive.Permissions.list(folderId, { fields: 'permissions(id,emailAddress,role,type)' });
        const permissions = (permissionList && permissionList.permissions) || [];
        permissions.forEach(p => {
          if (p.type === 'user' && p.emailAddress && p.emailAddress.toLowerCase() === email.toLowerCase()) {
            try {
              Drive.Permissions.remove(folderId, p.id);
            } catch (e) {
              Logger.log(`Cannot remove direct permission for ${email}: ${e.message}`);
            }
          }
        });
      } catch (e) {
        Logger.log(`Cannot inspect advanced permissions for ${email}: ${e.message}`);
      }
    }
  } catch (error) {
    Logger.log(`Error clearing direct permission for ${email}: ${error.message}`);
  }
}
// Cập nhật function addCommenterNoEmail mới
function addCommenterNoEmail(fileId, email) {
  const token = ScriptApp.getOAuthToken();
  createPermissionByApi(fileId, email, 'commenter', token);
}

// Cập nhật function writePermissionsSheet - UPDATE thay vì ghi đè
function writePermissionsSheet(classname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Permissions");
  if (!sheet) {
    sheet = ss.insertSheet("Permissions");
    // Tạo header cho sheet mới (KHÔNG có Folder ID)
    const initHeader = ["Class name", "Group name", "Role", "Emails", "Permission"];
    sheet.getRange(1, 1, 1, initHeader.length).setValues([initHeader])
      .setBackground("#d9ead3").setFontWeight("bold");
  }

  const root = getSpreadsheetParent();
  const userprofile = getOrCreateFolder(root, "userprofile");
  const classFolder = getOrCreateFolder(userprofile, classname);

  // 1) Lấy nhóm có folder trong Drive
  const groupsFromSheet = getAllGroupOfClass(classname);
  const groupsInfo = [];
  let globalMaxDepth = 0;

  groupsFromSheet.forEach(group => {
    const groupIter = classFolder.getFoldersByName(group);
    if (!groupIter.hasNext()) {
      Logger.log("Skip group (no folder): " + group);
      return;
    }
    const groupFolder = groupIter.next();
    const tree = collectFolderIds(groupFolder);
    const depth = getMaxDepth(tree);
    globalMaxDepth = Math.max(globalMaxDepth, depth);
    const members = getGroupMembers(classname, group);
    groupsInfo.push({ name: group, members: members, tree: tree });
  });

  if (groupsInfo.length === 0) {
    Logger.log("⚠️ No groups with folders found for class: " + classname);
    return;
  }

  // 2) Xác định header cần thiết (KHÔNG có Folder ID)
  const levelsCount = Math.max(0, globalMaxDepth - 1);
  const newHeader = ["Class name", "Group name", "Role", "Emails", "Permission"];
  for (let i = 1; i <= levelsCount; i++) {
    newHeader.push(`LEVEL ${i}`, "Emails", "Permission");
  }

  // 3) Đọc dữ liệu hiện có để UPDATE
  const existingData = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
  const existingHeader = existingData.length > 0 ? existingData[0] : [];

  // 4) Chuẩn hóa header: luôn đúng schema hiện tại (không giữ cột legacy).
  const currentLastCol = Math.max(sheet.getLastColumn(), 1);
  const currentHeader = sheet.getRange(1, 1, 1, currentLastCol).getValues()[0];
  const sameHeader = currentHeader.length === newHeader.length &&
    newHeader.every((h, idx) => (currentHeader[idx] || "") === h);

  if (!sameHeader) {
    sheet.getRange(1, 1, 1, newHeader.length).setValues([newHeader])
      .setBackground("#d9ead3").setFontWeight("bold");

    if (currentLastCol > newHeader.length) {
      sheet.deleteColumns(newHeader.length + 1, currentLastCol - newHeader.length);
    }

    Logger.log(`📝 Đã chuẩn hóa header Permissions về ${newHeader.length} cột`);
  }

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  function levelBaseIndex(level) {
    return 5 + (level - 1) * 3; // 0-based, mỗi level có 3 cột: Name, Emails, Permission
  }

  // 5) XÓA tất cả rows cũ của class này (sẽ ghi lại toàn bộ)
  const classnameIdx = existingHeader.indexOf("Class name");
  const rowsToDelete = [];

  for (let i = 1; i < existingData.length; i++) {
    const cn = (existingData[i][classnameIdx] || "").toString().trim();
    if (cn === classname) {
      rowsToDelete.push(i + 1); // Sheet row (1-based)
    }
  }

  // Xóa từ dưới lên để không lệch index
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  Logger.log(`🗑️ Đã xóa ${rowsToDelete.length} rows cũ của class ${classname}`);

  // 6) Build ALL rows mới hoàn toàn
  const allNewRows = [];
  const folderLinksToSet = []; // Lưu thông tin để set link sau

  groupsInfo.forEach(info => {
    const members = info.members || [];
    const leader = members.length > 0 ? members[0] : "";
    const otherMembers = members.slice(1);

    function traverse(node, level = 0) {
      const leaderRow = Array(header.length).fill("");
      const membersRow = Array(header.length).fill("");

      if (level === 0) {
        // Root level - Group folder
        leaderRow[0] = classname;
        leaderRow[1] = info.name;
        leaderRow[2] = "Leader";
        leaderRow[3] = leader;
        leaderRow[4] = "editor";

        membersRow[0] = classname;
        membersRow[1] = info.name;
        membersRow[2] = "Members";
        membersRow[3] = otherMembers.join(", ");
        membersRow[4] = "viewer";

        const rowIndex = allNewRows.length;
        allNewRows.push(leaderRow);
        allNewRows.push(membersRow);

        // Lưu thông tin để set link cho Group name (cột B)
        folderLinksToSet.push({
          row: rowIndex,
          col: 1, // Group name column (0-based)
          name: info.name,
          folderId: node.id
        });
        folderLinksToSet.push({
          row: rowIndex + 1,
          col: 1, // Group name column (0-based)
          name: info.name,
          folderId: node.id
        });

      } else {
        // Sub-level folders
        const base = levelBaseIndex(level);
        if (base + 2 < header.length) {
          leaderRow[0] = classname;
          leaderRow[1] = info.name;
          leaderRow[2] = "Leader";
          leaderRow[base] = node.name || "";
          leaderRow[base + 1] = leader;
          leaderRow[base + 2] = "editor";

          membersRow[0] = classname;
          membersRow[1] = info.name;
          membersRow[2] = "Members";
          membersRow[base] = node.name || "";
          membersRow[base + 1] = otherMembers.join(", ");
          membersRow[base + 2] = "viewer";

          const rowIndex = allNewRows.length;
          allNewRows.push(leaderRow);
          allNewRows.push(membersRow);

          // Lưu thông tin để set link cho LEVEL name
          folderLinksToSet.push({
            row: rowIndex,
            col: base, // LEVEL column (0-based)
            name: node.name,
            folderId: node.id
          });
          folderLinksToSet.push({
            row: rowIndex + 1,
            col: base, // LEVEL column (0-based)
            name: node.name,
            folderId: node.id
          });
        }
      }

      if (node.children && node.children.length > 0) {
        node.children.forEach(child => traverse(child, level + 1));
      }
    }

    traverse(info.tree, 0);
  });

  // 7) Insert tất cả rows mới vào sheet
  if (allNewRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, allNewRows.length, header.length).setValues(allNewRows);

    // 8) Set links cho folder names
    folderLinksToSet.forEach(linkInfo => {
      const cell = sheet.getRange(startRow + linkInfo.row, linkInfo.col + 1);
      const folderUrl = `https://drive.google.com/drive/folders/${linkInfo.folderId}`;
      const formula = `=HYPERLINK("${folderUrl}"; "${linkInfo.name}")`;
      cell.setFormula(formula);
    });

    // 9) Áp dụng data validation cho Permission columns (KHÔNG cho LEVEL columns)
    const permRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["owner", "editor", "commenter", "viewer"], true)
      .setAllowInvalid(true)
      .build();

    for (let i = 0; i < header.length; i++) {
      // Chỉ áp dụng validation cho cột Permission, KHÔNG cho LEVEL
      if (header[i] === "Permission") {
        sheet.getRange(startRow, i + 1, allNewRows.length).setDataValidation(permRule);
      }
    }

    Logger.log(`✅ Đã ghi ${allNewRows.length} rows cho class ${classname} vào Permissions sheet`);
    Logger.log(`✅ Đã set ${folderLinksToSet.length} folder links`);
  }

  Logger.log(`✅ Hoàn tất UPDATE Permissions sheet cho class ${classname}`);
}

/**
 * Tạo sheet Dashboard với thông tin submission
 * Được gọi khi Generate hoặc Change folder structure
 */
function writeDashboardSheet(classname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Dashboard");

  // Tạo hoặc cập nhật header
  const expectedHeader = ["Class name", "Group name", "Assignment", "Due day", "Submission status", "Extension requirement", "Last submission", "Overdue"];

  if (!sheet) {
    sheet = ss.insertSheet("Dashboard");
  }

  // Đảm bảo header luôn tồn tại và đúng
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, expectedHeader.length).setValues([expectedHeader])
      .setBackground("#c9daf8").setFontWeight("bold");
  } else {
    // Kiểm tra và cập nhật header nếu cần
    const currentHeader = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), expectedHeader.length)).getValues()[0];
    if (currentHeader.length < expectedHeader.length || currentHeader[0] !== expectedHeader[0]) {
      sheet.getRange(1, 1, 1, expectedHeader.length).setValues([expectedHeader])
        .setBackground("#c9daf8").setFontWeight("bold");
    }
  }

  const root = getSpreadsheetParent();
  const userprofile = getOrCreateFolder(root, "userprofile");
  const classFolder = getOrCreateFolder(userprofile, classname);

  // 1) Lấy nhóm có folder trong Drive
  const groupsFromSheet = getAllGroupOfClass(classname);
  const groupsInfo = [];

  groupsFromSheet.forEach(group => {
    const groupIter = classFolder.getFoldersByName(group);
    if (!groupIter.hasNext()) {
      Logger.log("Skip group (no folder): " + group);
      return;
    }
    const groupFolder = groupIter.next();
    const tree = collectFolderIds(groupFolder);
    const members = getGroupMembers(classname, group);
    groupsInfo.push({ name: group, members: members, tree: tree });
  });

  if (groupsInfo.length === 0) {
    Logger.log("⚠️ No groups with folders found for class: " + classname);
    return;
  }

  // Lấy header sau khi đã đảm bảo tồn tại
  const header = expectedHeader;

  // 2) XÓA tất cả rows cũ của class này
  const existingData = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
  const classnameIdx = 0; // Class name ở cột đầu tiên
  const rowsToDelete = [];

  for (let i = 1; i < existingData.length; i++) {
    const cn = (existingData[i][classnameIdx] || "").toString().trim();
    if (cn === classname) {
      rowsToDelete.push(i + 1);
    }
  }

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  Logger.log(`🗑️ Đã xóa ${rowsToDelete.length} rows cũ của class ${classname} trong Dashboard`);

  // 3) Build rows mới
  const allNewRows = [];
  const assignmentLinksToSet = [];
  const groupLinksToSet = [];

  groupsInfo.forEach(info => {
    // Lấy ID của group folder
    const groupFolderIter = classFolder.getFoldersByName(info.name);
    const groupFolderId = groupFolderIter.hasNext() ? groupFolderIter.next().getId() : null;

    function traverse(node, level = 0) {
      // Lấy TẤT CẢ các folders (trừ root level)
      if (level > 0) {
        const row = Array(header.length).fill("");
        row[0] = classname;
        row[1] = info.name;
        row[2] = node.name; // Assignment name
        row[3] = ""; // Due day - để trống
        row[4] = ""; // Submission status - để trống
        row[5] = ""; // Extension requirement - để trống
        row[6] = ""; // Last submission - để trống
        row[7] = ""; // Overdue - để trống

        const rowIndex = allNewRows.length;
        allNewRows.push(row);

        // Lưu thông tin để set link cho Group name
        if (groupFolderId) {
          groupLinksToSet.push({
            row: rowIndex,
            col: 1, // Group name column (0-based)
            name: info.name,
            folderId: groupFolderId
          });
        }

        // Lưu thông tin để set link cho Assignment
        assignmentLinksToSet.push({
          row: rowIndex,
          col: 2, // Assignment column (0-based)
          name: node.name,
          folderId: node.id
        });
      }

      // Đệ quy vào children
      if (node.children && node.children.length > 0) {
        node.children.forEach(child => traverse(child, level + 1));
      }
    }

    traverse(info.tree, 0);
  });

  // 4) Insert rows mới vào sheet
  if (allNewRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, allNewRows.length, header.length).setValues(allNewRows);

    // 5) Set hyperlink cho Group name column
    groupLinksToSet.forEach(linkInfo => {
      const cell = sheet.getRange(startRow + linkInfo.row, linkInfo.col + 1);
      const folderUrl = `https://drive.google.com/drive/folders/${linkInfo.folderId}`;
      const formula = `=HYPERLINK("${folderUrl}"; "${linkInfo.name}")`;
      cell.setFormula(formula);
    });

    // 6) Set hyperlink cho Assignment column
    assignmentLinksToSet.forEach(linkInfo => {
      const cell = sheet.getRange(startRow + linkInfo.row, linkInfo.col + 1);
      const folderUrl = `https://drive.google.com/drive/folders/${linkInfo.folderId}`;
      const formula = `=HYPERLINK("${folderUrl}"; "${linkInfo.name}")`;
      cell.setFormula(formula);
    });

    Logger.log(`✅ Đã ghi ${allNewRows.length} rows cho class ${classname} vào Dashboard sheet`);
  }

  Logger.log(`✅ Hoàn tất UPDATE Dashboard sheet cho class ${classname}`);
}



/**
 * onEdit đã được disable - không cần multi-select email nữa
 */
function onEdit(e) {
  // Đã disable tính năng này
  return;
}

/**
 * Tính độ sâu tối đa của cây
 */
function getMaxDepth(node, depth = 1) {
  let maxDepth = depth;
  if (node.children && node.children.length > 0) {
    node.children.forEach(child => {
      maxDepth = Math.max(maxDepth, getMaxDepth(child, depth + 1));
    });
  }
  return maxDepth;
}

/**
 * Thu thập ID + details thư mục
 */
function collectFolderIdsWithDetails(folder) {
  const children = [];
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    children.push(collectFolderIdsWithDetails(sub));
  }
  return { name: folder.getName(), id: folder.getId(), children: children };
}


////////////////////////////////////////////////////////////////////////////


/**
 * Recursively collect IDs of a folder and its subfolders
 * returns { name, id, children: [] }
 */
function collectFolderIds(folder) {
  let children = [];
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    children.push(collectFolderIds(sub));
  }
  return { name: folder.getName(), id: folder.getId(), children: children };
}


/**
 * Utility: get parent folder of spreadsheet (or root)
 */
function getSpreadsheetParent() {
  const ssFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = ssFile.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

/**
 * Utility: get or create subfolder
 */
function getOrCreateFolder(parent, name) {
  let folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

/**
 * Create folder structure from JSON
 */
function createFolders(nodes, parent) {
  nodes.forEach(node => {
    const folderName = (node && node.name ? node.name : '').toString().trim();

    // Bỏ qua node rỗng để tránh Drive service error khi createFolder("").
    if (!folderName) {
      if (node && node.children && node.children.length > 0) {
        createFolders(node.children, parent);
      }
      return;
    }

    let folder = driveWithRetry(() => parent.createFolder(folderName), `createFolder:${folderName}`);
    if (node.children && node.children.length > 0) {
      createFolders(node.children, folder);
    }
  });
}

/**
 * Copy folder contents recursively
 */
function copyContents(source, target) {
  // Copy files
  let files;
  try {
    files = driveWithRetry(() => source.getFiles(), `getFiles:${source.getName()}`);
  } catch (e) {
    Logger.log(`Cannot read files in ${source.getName()}: ${e.message}`);
    files = null;
  }

  if (files) {
    while (files.hasNext()) {
      const file = files.next();
      try {
        driveWithRetry(() => file.makeCopy(file.getName(), target), `makeCopy:${file.getName()}`);
      } catch (e) {
        // Một file lỗi không nên làm hỏng toàn bộ generate.
        Logger.log(`Skip copy file ${file.getName()}: ${e.message}`);
      }
    }
  }

  // Copy subfolders
  let subfolders;
  try {
    subfolders = driveWithRetry(() => source.getFolders(), `getFolders:${source.getName()}`);
  } catch (e) {
    Logger.log(`Cannot read subfolders in ${source.getName()}: ${e.message}`);
    subfolders = null;
  }

  if (subfolders) {
    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      try {
        const newSub = driveWithRetry(
          () => target.createFolder(subfolder.getName()),
          `createSubFolder:${subfolder.getName()}`
        );
        copyContents(subfolder, newSub);
      } catch (e) {
        Logger.log(`Skip subfolder ${subfolder.getName()}: ${e.message}`);
      }
    }
  }
}

/////////////////////////////////////////////
function showChangeFolderStructureSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('changeFolderStructureUI') // ref: changeFolderStructureUI.html
    .setTitle('Change Folder Structure')

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Cập nhật cấu trúc trong sheet "Create Folder" cho class đã có
 * Xóa các dòng cũ của class và ghi mới
 */
function updateClassTreeSheet(classname, obj) {
  if (!classname || !obj || obj.length === 0) {
    throw new Error("Class name hoặc cấu trúc folder không hợp lệ");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Create Folder');
  if (!sheet) {
    throw new Error("Không tìm thấy sheet 'Create Folder'");
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error("Sheet 'Create Folder' trống");
  }

  const header = data[0];
  const classnameColIdx = header.map(h => (h || '').toString().toLowerCase()).indexOf('class name');
  const groupnameColIdx = header.map(h => (h || '').toString().toLowerCase()).indexOf('group name');

  if (classnameColIdx === -1 || groupnameColIdx === -1) {
    throw new Error("Không tìm thấy cột 'Class name' hoặc 'Group name' trong sheet");
  }

  // Tìm block liên tục của class này và xóa TOÀN BỘ block
  let startDeleteRow = -1;
  let endDeleteRow = -1;
  let inClassBlock = false;

  for (let i = 1; i < data.length; i++) {
    const cn = (data[i][classnameColIdx] || "").toString().trim();

    // Nếu gặp class name khớp → bắt đầu block
    if (cn === classname) {
      if (startDeleteRow === -1) {
        startDeleteRow = i + 1; // Sheet row (1-based)
        Logger.log(`� Bắt đầu block class "${classname}" tại row ${startDeleteRow}`);
      }
      inClassBlock = true;
      endDeleteRow = i + 1;
    }
    // Nếu đang trong block và gặp class name KHÁC (không rỗng) → kết thúc block
    else if (inClassBlock && cn !== "" && cn !== classname) {
      Logger.log(`🛑 Kết thúc block tại row ${i + 1}, gặp class "${cn}"`);
      break;
    }
    // Nếu đang trong block và gặp row trống (Class name rỗng) → vẫn thuộc block
    else if (inClassBlock && cn === "") {
      endDeleteRow = i + 1;
    }
  }

  // Nếu block chạy đến cuối sheet
  if (inClassBlock && endDeleteRow === data.length) {
    Logger.log(`📍 Block chạy đến cuối sheet (row ${endDeleteRow})`);
  }

  // Xóa toàn bộ block nếu tìm thấy
  if (startDeleteRow !== -1 && endDeleteRow !== -1) {
    const numRowsToDelete = endDeleteRow - startDeleteRow + 1;
    Logger.log(`🗑️ Xóa block từ row ${startDeleteRow} đến ${endDeleteRow} (${numRowsToDelete} rows)`);

    // Xóa toàn bộ block một lần (từ trên xuống)
    for (let i = 0; i < numRowsToDelete; i++) {
      sheet.deleteRow(startDeleteRow); // Luôn xóa row đầu tiên của block
    }

    Logger.log(`✅ Đã xóa ${numRowsToDelete} dòng liên tục của class ${classname}`);
  } else {
    Logger.log(`⚠️ Không tìm thấy block nào của class ${classname}`);
  }

  // ====== GHI LẠI CẤU TRÚC MỚI (FORMAT GIỐNG writeClassTree) ======
  // Tính max depth từ template structure
  let maxDepth = 0;
  (function getMaxDepth(nodes, depth = 1) {
    nodes.forEach(node => {
      if (!node) return;
      maxDepth = Math.max(maxDepth, depth);
      if (node.children && node.children.length > 0) {
        getMaxDepth(node.children, depth + 1);
      }
    });
  })(obj);

  // Kiểm tra header có đủ cột không
  const requiredCols = 2 + maxDepth; // Class name, Group name, LEVEL 1, LEVEL 2, ...
  const existingHeader = header;

  if (existingHeader.length < requiredCols) {
    // Thêm cột thiếu
    const newHeader = ["Class name", "Group name"];
    for (let i = 1; i <= maxDepth; i++) newHeader.push(`LEVEL ${i}`);
    sheet.getRange(1, 1, 1, requiredCols).setValues([newHeader]);
  }

  // Build rows với format: Class name chỉ 1 lần, Group name = "{}"
  const allRows = [];
  let isFirstRow = true;

  function traverse(nodes, level = 1) {
    nodes.forEach(node => {
      if (!node || !node.name || node.name.toString().trim() === "") return;

      const row = Array(requiredCols).fill("");

      // Chỉ ghi Class name và Group name ở dòng đầu tiên
      if (isFirstRow) {
        row[0] = classname;
        row[1] = "{}";
        isFirstRow = false;
      }

      row[1 + level] = node.name.toString().trim();
      allRows.push(row);

      if (node.children && node.children.length > 0) {
        traverse(node.children, level + 1);
      }
    });
  }

  traverse(obj, 1);

  // Ghi tất cả rows vào sheet
  if (allRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    const writeRange = sheet.getRange(startRow, 1, allRows.length, requiredCols);
    // Keep folder names as plain text (avoid Date coercion for names like "4.1").
    writeRange.setNumberFormat('@');
    writeRange.setValues(allRows);
    Logger.log(`✅ Đã ghi ${allRows.length} dòng mới cho class ${classname}`);
  }

  Logger.log(`✅ Hoàn tất cập nhật Create Folder sheet cho class ${classname}`);
  return { success: true, message: "Đã cập nhật sheet thành công" };
}

/**
 * Cập nhật sheet Rules cho class
 * GIỮ NGUYÊN dữ liệu cũ, chỉ thêm mới hoặc xóa folder đã thay đổi
 */
function updateRulesForClass(classname, obj) {
  if (!classname || !obj || obj.length === 0) {
    throw new Error("Class name hoặc cấu trúc folder không hợp lệ");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ruleSheet = ss.getSheetByName("Rules");

  if (!ruleSheet) {
    ruleSheet = ss.insertSheet("Rules");
    ruleSheet.getRange(1, 1, 1, 3).setValues([['Class name', 'Folder', 'Number of file']]);
    ruleSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    ruleSheet.getRange(1, 1, 1, 3).setBackground('#d4e6f1');
    ruleSheet.autoResizeColumns(1, 3);
  } else {
    // Ensure header exists if sheet was created without it
    if (ruleSheet.getLastRow() === 0) {
      ruleSheet.getRange(1, 1, 1, 3).setValues([['Class name', 'Folder', 'Number of file']]);
      ruleSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      ruleSheet.getRange(1, 1, 1, 3).setBackground('#d4e6f1');
      ruleSheet.autoResizeColumns(1, 3);
    }
  }

  // Thu thập tất cả node từ cấu trúc mới
  const newFolderNames = new Set();
  (function collectAll(nodes) {
    nodes.forEach(node => {
      if (!node || !node.name) return;
      const n = node.name.toString().trim();
      if (n) newFolderNames.add(n);
      if (node.children && node.children.length > 0) collectAll(node.children);
    });
  })(obj);

  Logger.log(`📂 Cấu trúc mới có ${newFolderNames.size} folders: ${Array.from(newFolderNames).join(', ')}`);

  // Đọc dữ liệu cũ của class này
  const maxCols = ruleSheet.getMaxColumns();
  const allData = ruleSheet.getLastRow() > 1
    ? ruleSheet.getRange(2, 1, ruleSheet.getLastRow() - 1, maxCols).getValues()
    : [];

  const existingRulesMap = new Map(); // key: folderName, value: {row, data}
  const rowsToDelete = [];

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const cn = (row[0] || '').toString().trim();
    const folderName = (row[1] || '').toString().trim();

    if (cn === classname) {
      if (newFolderNames.has(folderName)) {
        // Folder vẫn tồn tại - lưu lại để giữ nguyên
        existingRulesMap.set(folderName, {
          rowIndex: i + 2, // Sheet row (1-based)
          data: row
        });
        Logger.log(`✓ Giữ nguyên rule: ${folderName}`);
      } else {
        // Folder đã bị xóa - đánh dấu để xóa
        rowsToDelete.push(i + 2);
        Logger.log(`✗ Xóa rule (folder không còn): ${folderName}`);
      }
    }
  }

  // Xóa các dòng không còn trong structure (từ dưới lên)
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    ruleSheet.deleteRow(rowsToDelete[i]);
  }
  Logger.log(`🗑️ Đã xóa ${rowsToDelete.length} rules không còn trong structure`);

  // Tìm các folder mới cần thêm
  const foldersToAdd = [];
  newFolderNames.forEach(folderName => {
    if (!existingRulesMap.has(folderName)) {
      foldersToAdd.push(folderName);
      Logger.log(`➕ Thêm mới rule: ${folderName}`);
    }
  });

  // Thêm các folder mới vào cuối
  if (foldersToAdd.length > 0) {
    const newRows = foldersToAdd.map(name => [classname, name, ""]);
    const start = ruleSheet.getLastRow() + 1;
    const writeRange = ruleSheet.getRange(start, 1, newRows.length, 3);
    // Force first 2 columns as text (Class, Folder), keep Number of file as general.
    writeRange.offset(0, 0, newRows.length, 2).setNumberFormat('@');
    writeRange.setValues(newRows);
    ruleSheet.autoResizeColumns(1, 3);
    Logger.log(`✅ Đã thêm ${newRows.length} rules mới`);
  }

  Logger.log(`✅ Hoàn tất update Rules: Giữ nguyên ${existingRulesMap.size}, Xóa ${rowsToDelete.length}, Thêm mới ${foldersToAdd.length}`);
  return { success: true, message: "Đã cập nhật Rules thành công (giữ nguyên dữ liệu cũ)" };
}

/**
 * Cập nhật cấu trúc folder trên Drive
 * KHÔNG xóa folder cũ - chỉ thêm mới và đổi tên (giữ nguyên data)
 */
function updateClassFolderStructure(classname, obj) {
  if (!classname || !obj || obj.length === 0) {
    throw new Error("Class name hoặc cấu trúc folder không hợp lệ");
  }

  const parentFolder = getSpreadsheetParent();
  let tempFolder = getOrCreateFolder(parentFolder, "temp");
  let classFolder = getOrCreateFolder(tempFolder, classname);

  // Lấy hoặc tạo template
  let templateFolder;
  let oldTemplateStructure = null;
  const oldTemplates = classFolder.getFoldersByName("_template");
  if (oldTemplates.hasNext()) {
    templateFolder = oldTemplates.next();
    // Lấy cấu trúc cũ VỚI ID để detect rename chính xác
    oldTemplateStructure = collectFolderIds(templateFolder).children || [];
    Logger.log("✅ Tìm thấy template hiện có");
  } else {
    templateFolder = classFolder.createFolder("_template");
    Logger.log("✅ Tạo template mới");
  }

  // Sync cấu trúc template
  try {
    syncFolderStructure(templateFolder, obj, oldTemplateStructure);
    Logger.log("✅ Đã sync template");
  } catch (e) {
    Logger.log(`❌ Sync template failed for ${classname}: ${e.message}`);
    throw e;
  }

  // Cập nhật tất cả các group folder
  const groups = getAllGroupOfClass(classname);
  let userProfileFolder = getOrCreateFolder(parentFolder, "userprofile");
  let classProfileFolder = getOrCreateFolder(userProfileFolder, classname);

  groups.forEach(groupName => {
    const groupFolders = classProfileFolder.getFoldersByName(groupName);
    if (!groupFolders.hasNext()) {
      Logger.log(`⚠️ Group ${groupName} không có folder, bỏ qua`);
      return;
    }

    const groupFolder = groupFolders.next();

    // Lấy cấu trúc cũ của group VỚI ID
    const oldGroupStructure = collectFolderIds(groupFolder).children || [];

    // Sync cấu trúc
    try {
      syncFolderStructure(groupFolder, obj, oldGroupStructure);

      // Áp dụng lại permissions cho folder mới
      const members = getGroupMembers(classname, groupName);
      applyGroupPermissionsToNewFolders(groupFolder, obj, members);

      Logger.log(`✅ Đã cập nhật folder cho group ${groupName}`);
    } catch (e) {
      Logger.log(`❌ Sync group folder failed for ${groupName}: ${e.message}`);
    }
  });

  return { success: true, message: "Đã cập nhật cấu trúc folder trên Drive thành công" };
}

/**
 * Sync cấu trúc folder: thêm mới, đổi tên, và XÓA folder không còn trong structure
 * @param {Folder} parentFolder - Folder cha trên Drive
 * @param {Array} newStructure - Cấu trúc mới từ UI (chỉ có name, children)
 * @param {Array} oldStructureWithIds - Cấu trúc cũ từ Drive (có name, id, children)
 */
function syncFolderStructure(parentFolder, newStructure, oldStructureWithIds) {
  if (!newStructure || newStructure.length === 0) {
    // Nếu structure mới rỗng → XÓA tất cả folder con
    Logger.log(`⚠️ Structure mới rỗng, xóa tất cả folder trong "${parentFolder.getName()}"`);
    let subfolders = null;
    try {
      subfolders = driveWithRetry(() => parentFolder.getFolders(), `getFolders:${parentFolder.getId()}`);
    } catch (e) {
      Logger.log(`Cannot list subfolders in ${parentFolder.getName()}: ${e.message}`);
    }

    if (subfolders) {
      while (subfolders.hasNext()) {
        const folder = subfolders.next();
        try {
          Logger.log(`🗑️ Xóa folder: "${folder.getName()}"`);
          driveWithRetry(() => {
            folder.setTrashed(true);
            return true;
          }, `trashFolder:${folder.getId()}`);
        } catch (e) {
          Logger.log(`Cannot trash folder ${folder.getName()}: ${e.message}`);
        }
      }
    }
    return;
  }

  // Lấy danh sách folder hiện có trên Drive
  const existingFolders = []; // [{name, id, folder}]
  const folderById = {}; // id → Folder object
  const folderByName = {}; // name → Folder object

  let subfolders = null;
  try {
    subfolders = driveWithRetry(() => parentFolder.getFolders(), `getFolders:${parentFolder.getId()}`);
  } catch (e) {
    Logger.log(`Cannot list folders in ${parentFolder.getName()}: ${e.message}`);
  }

  if (subfolders) {
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      try {
        const name = folder.getName();
        const id = folder.getId();

        existingFolders.push({ name, id, folder });
        folderById[id] = folder;
        folderByName[name] = folder;
      } catch (e) {
        Logger.log(`Skip unreadable folder under ${parentFolder.getName()}: ${e.message}`);
      }
    }
  }

  // Build map: tên mới → old ID (để detect rename)
  const oldIdByName = {};
  const oldChildrenById = {};
  if (oldStructureWithIds && oldStructureWithIds.length > 0) {
    oldStructureWithIds.forEach(oldNode => {
      if (oldNode && oldNode.name && oldNode.id) {
        oldIdByName[oldNode.name] = oldNode.id;
        oldChildrenById[oldNode.id] = oldNode.children || [];
      }
    });
  }

  // Build set tên folder trong structure mới
  const newFolderNames = new Set(newStructure.map(n => n.name ? n.name.trim() : null).filter(n => n));

  // Track các folder ID đã được xử lý
  const processedIds = new Set();

  // Duyệt qua cấu trúc mới
  newStructure.forEach((node, newIndex) => {
    if (!node || !node.name || node.name.trim() === '') return;

    const nodeName = node.name.trim();
    let currentFolder = null;
    let folderId = null;
    let oldChildrenStructure = null;

    // Case 1: Folder với tên này đã tồn tại → GIỮ NGUYÊN
    if (folderByName[nodeName]) {
      currentFolder = folderByName[nodeName];
      folderId = currentFolder.getId();
      processedIds.add(folderId);
      Logger.log(`📁 Giữ nguyên folder: "${nodeName}" (ID: ${folderId})`);

      // Lấy children từ oldStructure nếu folder này có ID match
      if (oldIdByName[nodeName]) {
        oldChildrenStructure = oldChildrenById[oldIdByName[nodeName]];
      }
    }
    // Case 2: Detect RENAME - tên mới chưa tồn tại
    else {
      // Tìm xem có folder cũ nào ở cùng vị trí index không
      let oldNodeAtSamePos = null;
      if (oldStructureWithIds && newIndex < oldStructureWithIds.length) {
        oldNodeAtSamePos = oldStructureWithIds[newIndex];
      }

      // Nếu có old node ở vị trí này VÀ folder với ID đó vẫn tồn tại VÀ tên khác
      if (oldNodeAtSamePos && oldNodeAtSamePos.id && folderById[oldNodeAtSamePos.id] &&
        oldNodeAtSamePos.name !== nodeName) {
        // ĐÂY LÀ RENAME!
        currentFolder = folderById[oldNodeAtSamePos.id];
        folderId = oldNodeAtSamePos.id;
        const oldName = currentFolder.getName();
        currentFolder.setName(nodeName);
        processedIds.add(folderId);
        Logger.log(`✏️ Đổi tên folder: "${oldName}" → "${nodeName}" (ID: ${folderId})`);

        // Lấy children từ old structure
        oldChildrenStructure = oldChildrenById[folderId];
      }
      // Case 3: TẠO folder mới
      else {
        try {
          currentFolder = driveWithRetry(() => parentFolder.createFolder(nodeName), `createFolder:${parentFolder.getId()}:${nodeName}`);
          folderId = currentFolder.getId();
          processedIds.add(folderId);
          Logger.log(`✨ Tạo folder mới: "${nodeName}" (ID: ${folderId})`);
        } catch (e) {
          Logger.log(`❌ Cannot create folder "${nodeName}" in ${parentFolder.getName()}: ${e.message}`);
          return;
        }
      }
    }

    // Đệ quy với các folder con
    if (currentFolder) {
      if (node.children && node.children.length > 0) {
        try {
          syncFolderStructure(currentFolder, node.children, oldChildrenStructure);
        } catch (e) {
          Logger.log(`❌ Sync children failed at ${currentFolder.getName()}: ${e.message}`);
        }
      } else if (node.children && node.children.length === 0 && oldChildrenStructure) {
        // Nếu children mới rỗng nhưng cũ có children → XÓA tất cả subfolder
        try {
          syncFolderStructure(currentFolder, [], oldChildrenStructure);
        } catch (e) {
          Logger.log(`❌ Clear children failed at ${currentFolder.getName()}: ${e.message}`);
        }
      }
    }
  });

  // XÓA các folder không còn trong structure mới (dựa vào ID đã xử lý)
  existingFolders.forEach(({ name, id, folder }) => {
    if (!processedIds.has(id)) {
      Logger.log(`🗑️ Xóa folder không còn trong structure: "${name}" (ID: ${id})`);
      try {
        driveWithRetry(() => {
          folder.setTrashed(true);
          return true;
        }, `trashFolder:${id}`);
      } catch (e) {
        Logger.log(`Cannot trash obsolete folder ${name} (${id}): ${e.message}`);
      }
    }
  });

  Logger.log(`✅ Sync hoàn tất cho "${parentFolder.getName()}"`);
}

/**
 * Lấy cấu trúc folder từ Drive (để so sánh)
 */
function getFolderStructureFromDrive(folder) {
  const children = [];
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    children.push({
      name: sub.getName(),
      children: getFolderStructureFromDrive(sub)
    });
  }
  return children;
}

/**
 * Áp dụng permissions chỉ cho các folder MỚI được tạo
 * Không làm gì với folder cũ (giữ nguyên permissions)
 */
function applyGroupPermissionsToNewFolders(parentFolder, structure, members) {
  if (!structure || structure.length === 0) return;

  structure.forEach(node => {
    if (!node.name || node.name.trim() === '') return;

    const nodeName = node.name.trim();
    const folders = parentFolder.getFoldersByName(nodeName);

    if (folders.hasNext()) {
      const folder = folders.next();

      // Kiểm tra xem folder này có phải mới tạo không
      // (cách đơn giản: kiểm tra created date gần đây)
      const createdDate = folder.getDateCreated();
      const now = new Date();
      const diffMinutes = (now - createdDate) / (1000 * 60);

      // Nếu folder được tạo trong vòng 5 phút → là folder mới
      if (diffMinutes < 5) {
        applyGroupPermissions(folder, members);
        Logger.log(`🔒 Đã set permissions cho folder mới: "${nodeName}"`);
      } else {
        Logger.log(`⏭️ Bỏ qua folder cũ: "${nodeName}" (giữ nguyên permissions)`);
      }

      // Đệ quy với folder con
      if (node.children && node.children.length > 0) {
        applyGroupPermissionsToNewFolders(folder, node.children, members);
      }
    }
  });
}

function sheetToObjectByClassname(classname = 'IE107') {
  if (!classname) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Create Folder");
  if (!sheet) return [];

  // Use display values to avoid Date objects being returned for text-like folder names.
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];

  const header = data[0];
  const classnameCol = header.indexOf("Class name");
  const levelCols = header.map((h, i) => h.startsWith("LEVEL") ? i : -1).filter(i => i !== -1);

  const tree = [];
  const stack = [];
  let currentClass = null;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (row.every(cell => !cell || cell.toString().trim() === "")) continue;

    if (row[classnameCol] && row[classnameCol].toString().trim() !== "") {
      currentClass = row[classnameCol].toString().trim();
      stack.length = 0;
    }

    if (currentClass !== classname) continue;

    let level = -1;
    let nodeName = null;
    for (let j = 0; j < levelCols.length; j++) {
      const val = row[levelCols[j]];
      if (val && val.toString().trim() !== "") {
        level = j;
        nodeName = val.toString().trim();
        break;
      }
    }
    if (level === -1) continue;

    const node = { name: nodeName, children: [] };

    if (level === 0) {
      tree.push(node);
      stack.length = 0;
      stack.push(node);
    } else {
      const parent = stack[level - 1];
      if (parent) {
        parent.children.push(node);
        stack[level] = node;
      }
    }
  }

  Logger.log(JSON.stringify(tree, null, 2));
  return tree;
}

/**
 * Xử lý sheet Permissions
 */
function updatePerrmissions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const permSheet = ss.getSheetByName("Permissions");

    if (!permSheet) {
      SpreadsheetApp.getUi().alert('Không tìm thấy sheet Permissions. Vui lòng tạo permissions trước.');
      return;
    }

    const range = permSheet.getDataRange();
    const data = range.getValues();
    const formulas = range.getFormulas();

    if (data.length === 0) {
      SpreadsheetApp.getUi().alert('Sheet Permissions trống.');
      return;
    }

    const hasHeader = (data[0][0] || "").toString().trim() === "Class name";
    const startRow = hasHeader ? 1 : 0;

    // Gom theo folderId -> email -> permission theo đúng thứ tự dòng.
    const folderPermissionMap = {};

    for (let rowIdx = startRow; rowIdx < data.length; rowIdx++) {
      const row = data[rowIdx];
      const rowFormulas = formulas[rowIdx] || [];

      for (let folderCol = 0; folderCol < rowFormulas.length; folderCol++) {
        const formula = rowFormulas[folderCol];
        const folderId = extractFolderIdFromFormula(formula);
        if (!folderId) continue;

        // Layout thực tế:
        // - Group name ở cột B(1), Emails ở D(3), Permission ở E(4)
        // - LEVEL n ở cột F/I/L..., Emails ngay sau đó, Permission kế tiếp
        const emailCol = folderCol === 1 ? 3 : folderCol + 1;
        const permCol = folderCol === 1 ? 4 : folderCol + 2;

        if (emailCol >= row.length || permCol >= row.length) continue;

        const emailsStr = row[emailCol];
        const permission = normalizePermissionValue(row[permCol]);
        if (!emailsStr || !permission) continue;

        const emails = emailsStr
          .toString()
          .split(/[;,\n]/)
          .map(e => e.trim().toLowerCase())
          .filter(e => e);

        if (emails.length === 0) continue;

        if (!folderPermissionMap[folderId]) {
          folderPermissionMap[folderId] = {};
        }

        emails.forEach(email => {
          // Tôn trọng giá trị Permission theo đúng dòng người dùng chỉnh sửa.
          // Nếu email xuất hiện nhiều lần trong cùng folder, dòng dưới sẽ ghi đè dòng trên.
          folderPermissionMap[folderId][email] = permission;
        });
      }
    }

    let appliedCount = 0;
    let folderCount = 0;

    Object.keys(folderPermissionMap).forEach(folderId => {
      const emailPermission = folderPermissionMap[folderId];
      const emails = Object.keys(emailPermission);
      if (emails.length === 0) return;

      try {
        folderCount++;
        removeNonOwnerPermissions(folderId);

        emails.forEach(email => {
          const permission = emailPermission[email];
          try {
            switch (permission) {
              case 'owner':
                // Không thể set owner trực tiếp bằng create permission, fallback editor.
                addEditorNoEmail(folderId, email);
                break;
              case 'editor':
                addEditorNoEmail(folderId, email);
                break;
              case 'commenter':
                addCommenterNoEmail(folderId, email);
                break;
              case 'viewer':
                addViewerNoEmail(folderId, email);
                break;
              default:
                return;
            }
            appliedCount++;
          } catch (e) {
            Logger.log(`Failed to apply ${permission} for ${email} on ${folderId}: ${e.message}`);
          }
        });
      } catch (folderError) {
        Logger.log(`Skip folder ${folderId} due to Drive error: ${folderError.message}`);
      }
    });

    SpreadsheetApp.getUi().alert(`✅ Đã cập nhật ${appliedCount} quyền trên ${folderCount} folder.`);
    Logger.log(`✅ Updated ${appliedCount} permissions across ${folderCount} folders`);

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Lỗi cập nhật permissions: ' + error.toString());
    Logger.log('Error updating permissions: ' + error);
  }
}

function normalizePermissionValue(permission) {
  if (!permission) return null;
  const p = permission.toString().trim().toLowerCase();
  if (p === 'commentor') return 'commenter';
  if (p === 'owner' || p === 'editor' || p === 'commenter' || p === 'viewer') return p;
  return null;
}

/**
 * Helper function: Extract Folder ID từ hyperlink formula
 * Formula format: =HYPERLINK("https://drive.google.com/drive/folders/[FOLDER_ID]"; "[NAME]")
 */
function extractFolderIdFromFormula(formula) {
  if (!formula || typeof formula !== 'string') return null;
  
  // Pattern: /folders/[FOLDER_ID]
  const match = formula.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

