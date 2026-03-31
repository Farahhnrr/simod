function doGet(e) {
  var page = sanitizeText(e && e.parameter && e.parameter.page).toLowerCase();
  if (page === 'admin') {
    assertAdminAccess_();
    return renderPage_('Admin', 'SIMOD HPS Admin');
  }
  return renderPage_('Index', 'SIMOD HPS Document Manager');
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

function assertAdminAccess_() {
  var allowedEmails = (CONFIG.ADMIN_ALLOWED_EMAILS || [])
    .map(function (email) { return sanitizeText(email).toLowerCase(); })
    .filter(function (email) { return !!email; });
  var activeEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();

  if (!allowedEmails.length) {
    throw new Error('ADMIN_ALLOWED_EMAILS belum dikonfigurasi di Config.gs.');
  }

  if (!activeEmail || allowedEmails.indexOf(activeEmail) === -1) {
    throw new Error('Anda tidak memiliki akses ke halaman admin.');
  }
}

function getBootstrapData() {
  var config = getProjectConfig();
  var response = {
    profile: {
      email: Session.getActiveUser().getEmail() || 'unknown'
    },
    configured: config.configured,
    stats: {
      totalEvents: 0,
      totalHps: 0,
      readyHps: 0,
      draftHps: 0
    },
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
  response.stats = getStats(events, packages);
  return response;
}

function addEducation(eventName) {
  return addEducationRecord(eventName);
}

function createHps(payload) {
  return createHpsRecord(payload);
}

function uploadHpsFiles(payload) {
  return uploadHpsFilesRecord(payload);
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

function getAdminBootstrapData() {
  assertAdminAccess_();
  var spreadsheetState = getSpreadsheetState();
  var driveState = getDriveFolderState();
  var authState = getAuthorizationState();

  var response = {
    profile: {
      email: Session.getActiveUser().getEmail() || 'unknown'
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
    events: [],
    packages: []
  };

  if (!spreadsheetState.ok || !driveState.ok) {
    return response;
  }

  setupSheets();

  var events = listEducationEvents();
  var packages = listHpsPackages({});
  response.events = events;
  response.packages = packages;
  response.stats = getStats(events, packages);
  return response;
}
