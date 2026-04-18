/**
 * Helper functions cho Permission management
 */

function getPermissionForCurrentUser() {
  const currentUser = Session.getActiveUser().getEmail();
  return checkPermission(currentUser);
}

function hasAdminPermission() {
  const permission = getPermissionForCurrentUser();
  return permission.hasPermission && permission.permissionLevel === 'Admin';
}

function hasEditorPermission() {
  const permission = getPermissionForCurrentUser();
  return permission.hasPermission &&
    (permission.permissionLevel === 'Admin' || permission.permissionLevel === 'Editor');
}

function validatePermissionAccess(requiredLevel = 'Viewer') {
  const permission = getPermissionForCurrentUser();

  if (!permission.hasPermission) {
    throw new Error('Bạn không có quyền truy cập hệ thống này');
  }

  const levels = ['Viewer', 'Contributor', 'Editor', 'Admin'];
  const userLevelIndex = levels.indexOf(permission.permissionLevel);
  const requiredLevelIndex = levels.indexOf(requiredLevel);

  if (userLevelIndex < requiredLevelIndex) {
    throw new Error(`Bạn cần quyền ${requiredLevel} để thực hiện thao tác này`);
  }

  return true;
}
