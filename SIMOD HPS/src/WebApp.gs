function doGet(e) {
  var page = sanitizeText(e && e.parameter && e.parameter.page).toLowerCase();
  if (page === 'admin') {
    return renderPage_('Admin', 'SIMOD HPS Admin');
  }
  return renderPage_('Index', 'SIMOD HPS Pengelola Dokumen');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function renderPage_(fileName, title) {
  var template = HtmlService.createTemplateFromFile(fileName);
  var appBaseUrl = ScriptApp.getService().getUrl() || '';
  template.appBaseUrl = appBaseUrl;
  template.userUiUrl = appBaseUrl || '';
  template.adminUiUrl = appBaseUrl ? appBaseUrl + '?page=admin' : '?page=admin';

  return template.evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getBootstrapData(sessionToken) {
  var session = requireApprovedSession_(sessionToken);
  var config = getProjectConfig();
  var response = {
    profile: {
      email: session.email || 'unknown',
      name: session.name || '',
      picture: session.picture || ''
    },
    configured: config.configured,
    stats: {
      totalEvents: 0,
      totalHps: 0,
      readyHps: 0,
      draftHps: 0
    },
    notifications: [],
    unreadNotificationCount: 0,
    warning: '',
    events: [],
    packages: []
  };

  var spreadsheetState = getSpreadsheetState();
  if (!spreadsheetState.ok) {
    response.configured = false;
    response.warning = spreadsheetState.message;
    return response;
  }

  if (spreadsheetState.recovered) {
    response.warning = spreadsheetState.message;
  }

  var driveState = getDriveFolderState();
  response.configured = !!(spreadsheetState.ok && driveState.ok);
  if (!driveState.ok) {
    response.warning = driveState.message;
  } else if (driveState.recovered) {
    response.warning = driveState.message;
  }

  setupSheets();

  var events = listEducationEvents();
  var packages = listHpsPackages({});

  response.events = events;
  response.packages = packages;
  response.notifications = listNotifications_(25, {
    audience: 'USER',
    recipientEmail: session.email
  });
  response.unreadNotificationCount = response.notifications.filter(function (item) {
    return !item.isRead;
  }).length;
  response.stats = getStats(events, packages);
  return response;
}

function addEducation(sessionToken, eventName) {
  var session = requireAdminSession_(sessionToken);
  return addEducationRecord(eventName, session.email);
}

function updateEducationStatus(sessionToken, eventId, nextStatus) {
  var session = requireAdminSession_(sessionToken);
  return updateEducationStatusRecord(eventId, nextStatus, session.email);
}

function deleteEducation(sessionToken, eventId) {
  var session = requireAdminSession_(sessionToken);
  return deleteEducationRecord(eventId, session.email);
}

function createHps(sessionToken, payload) {
  var session = requireApprovedSession_(sessionToken);
  payload = payload || {};
  if (sanitizeText(payload.noPesanan)) {
    throw new Error('No. Surat Pesanan hanya dapat diisi oleh admin.');
  }
  return createHpsRecord(payload, session.email);
}

function updateHps(sessionToken, payload) {
  requireApprovedSession_(sessionToken);
  payload = payload || {};
  if (sanitizeText(payload.noPesanan)) {
    throw new Error('No. Surat Pesanan hanya dapat diubah oleh admin.');
  }
  return updateHpsRecord(payload);
}

function uploadHpsFiles(sessionToken, payload) {
  var session = requireApprovedSession_(sessionToken);
  payload = payload || {};
  payload.files = filterFilesByRole_(payload.files || {}, CONFIG.USER_UPLOAD_KEYS, 'admin');
  return uploadHpsFilesRecord(payload, session.email);
}

function adminSaveRestrictedData(sessionToken, payload) {
  var session = requireAdminSession_(sessionToken);
  payload = payload || {};
  payload.files = filterFilesByRole_(payload.files || {}, CONFIG.ADMIN_UPLOAD_KEYS, 'pengguna');
  return updateAdminRestrictedRecord(payload, session.email);
}

function deleteHpsDocument(sessionToken, packageId, fileKey) {
  requireAdminSession_(sessionToken);
  return deleteHpsDocumentRecord(packageId, fileKey);
}

function startUserSession(email, displayName) {
  return authenticateEmail(email, displayName);
}

function startAdminSession(adminCode) {
  return authenticateAdmin(adminCode);
}

function deleteAccess(sessionToken, targetEmail) {
  return deleteAccessRecord(sessionToken, targetEmail);
}

function getAdminNotifications(sessionToken) {
  requireAdminSession_(sessionToken);
  var notifications = listNotifications_(25, { audience: 'ADMIN' });
  return {
    notifications: notifications,
    unreadCount: notifications.filter(function (item) { return !item.isRead; }).length
  };
}

function getUserNotifications(sessionToken) {
  var session = requireApprovedSession_(sessionToken);
  var notifications = listNotifications_(25, {
    audience: 'USER',
    recipientEmail: session.email
  });
  return {
    notifications: notifications,
    unreadCount: notifications.filter(function (item) { return !item.isRead; }).length
  };
}

function markNotificationsRead(sessionToken, notificationIds) {
  requireAdminSession_(sessionToken);
  return markNotificationsRead_(notificationIds, { audience: 'ADMIN' });
}

function markUserNotificationsRead(sessionToken, notificationIds) {
  var session = requireApprovedSession_(sessionToken);
  return markNotificationsRead_(notificationIds, {
    audience: 'USER',
    recipientEmail: session.email
  });
}

function getStats(events, packages) {
  var ready = packages.filter(function (pkg) { return pkg.status === 'READY'; }).length;
  return {
    totalEvents: events.length,
    totalHps: packages.length,
    readyHps: ready,
    draftHps: packages.length - ready
  };
}

function getAuthorizationState() {
  var info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var required = info.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED;
  return {
    required: required,
    url: info.getAuthorizationUrl() || ''
  };
}

function getAdminBootstrapData(sessionToken) {
  var session = requireAdminSession_(sessionToken);
  var spreadsheetState = getSpreadsheetState();
  var driveState = getDriveFolderState();
  var authState = getAuthorizationState();

  var response = {
    profile: {
      email: session.email || 'unknown',
      name: session.name || ''
    },
    auth: authState,
    config: {
      staticSheetId: sanitizeText(CONFIG.STATIC_SHEET_ID),
      staticDriveFolderId: sanitizeText(CONFIG.STATIC_DRIVE_FOLDER_ID),
      effectiveSheetId: spreadsheetState.ok ? spreadsheetState.sheetId : '',
      effectiveDriveFolderId: driveState.ok ? driveState.driveFolderId : '',
      spreadsheetStatus: spreadsheetState.ok ? 'READY' : 'ERROR',
      driveStatus: driveState.ok ? 'READY' : 'ERROR',
      spreadsheetMessage: spreadsheetState.message || '',
      driveMessage: driveState.message || ''
    },
    stats: {
      totalEvents: 0,
      totalHps: 0,
      readyHps: 0,
      draftHps: 0
    },
    accessRecords: [],
    notifications: [],
    unreadNotificationCount: 0,
    events: [],
    packages: []
  };

  if (!spreadsheetState.ok) {
    return response;
  }

  setupSheets();
  response.accessRecords = listAccessRecords_();
  response.notifications = listNotifications_(25, { audience: 'ADMIN' });
  response.unreadNotificationCount = response.notifications.filter(function (item) {
    return !item.isRead;
  }).length;

  if (!driveState.ok) {
    return response;
  }

  var events = listEducationEvents();
  var packages = listHpsPackages({});
  response.events = events;
  response.packages = packages;
  response.stats = getStats(events, packages);
  return response;
}

function filterFilesByRole_(files, allowedKeys, ownerLabel) {
  var nextFiles = {};
  var allowedMap = {};

  (allowedKeys || []).forEach(function (key) {
    allowedMap[key] = true;
  });

  Object.keys(files || {}).forEach(function (key) {
    if (!files[key]) return;
    if (!allowedMap[key]) {
      var fileConfig = CONFIG.FILE_COLUMNS[key];
      var fileLabel = fileConfig ? fileConfig.label : key;
      throw new Error(fileLabel + ' hanya dapat diunggah oleh ' + ownerLabel + '.');
    }
    nextFiles[key] = files[key];
  });

  return nextFiles;
}
