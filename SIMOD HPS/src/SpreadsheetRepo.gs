function getProjectConfig() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = sanitizeText(CONFIG.STATIC_SHEET_ID) || sanitizeText(props.getProperty('SHEET_ID'));
  var driveFolderId = sanitizeText(CONFIG.STATIC_DRIVE_FOLDER_ID) || sanitizeText(props.getProperty('DRIVE_FOLDER_ID'));

  return {
    sheetId: sheetId,
    driveFolderId: driveFolderId,
    configured: !!(sheetId && driveFolderId)
  };
}

function getRequiredProperty(key) {
  var config = getProjectConfig();
  var value = '';
  if (key === 'SHEET_ID') value = config.sheetId;
  if (key === 'DRIVE_FOLDER_ID') value = config.driveFolderId;
  if (!value) throw new Error('Missing configuration value: ' + key + '.');
  return value;
}

function getSpreadsheet() {
  var state = getSpreadsheetState();
  if (!state.ok) throw new Error(state.message);
  return state.spreadsheet;
}

function getSpreadsheetState() {
  var config = getProjectConfig();
  var sheetId = config.sheetId;

  if (!sheetId) {
    return {
      ok: false,
      code: 'missing',
      message: 'SHEET_ID belum diatur di Config.gs atau Script Properties.'
    };
  }

  try {
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    spreadsheet.getId();
    return {
      ok: true,
      spreadsheet: spreadsheet,
      sheetId: spreadsheet.getId(),
      recovered: false,
      message: ''
    };
  } catch (err) {
    if (sanitizeText(CONFIG.STATIC_SHEET_ID)) {
      return {
        ok: false,
        code: 'inaccessible-static',
        message: 'Spreadsheet pada STATIC_SHEET_ID tidak ditemukan atau Anda tidak punya akses. Perbarui ID tersebut di Config.gs.'
      };
    }

    try {
      var props = PropertiesService.getScriptProperties();
      var spreadsheetName = 'SIMOD HPS Data';
      var createdSpreadsheet = SpreadsheetApp.create(spreadsheetName);
      props.setProperty('SHEET_ID', createdSpreadsheet.getId());

      return {
        ok: true,
        spreadsheet: createdSpreadsheet,
        sheetId: createdSpreadsheet.getId(),
        recovered: true,
        message: 'Spreadsheet lama tidak bisa diakses. Sistem membuat spreadsheet baru secara otomatis.'
      };
    } catch (createErr) {
      return {
        ok: false,
        code: 'inaccessible-dynamic',
        message: 'Spreadsheet tidak bisa diakses dan spreadsheet pengganti gagal dibuat. Periksa izin Google Sheets akun yang menjalankan web app.'
      };
    }
  }
}

function getDriveFolderState() {
  var config = getProjectConfig();
  var driveFolderId = config.driveFolderId;

  if (!driveFolderId) {
    return {
      ok: false,
      code: 'missing',
      message: 'DRIVE_FOLDER_ID belum diatur di Config.gs atau Script Properties.'
    };
  }

  try {
    var folder = DriveApp.getFolderById(driveFolderId);
    folder.getId();
    return {
      ok: true,
      folder: folder,
      driveFolderId: folder.getId(),
      recovered: false,
      message: ''
    };
  } catch (err) {
    if (sanitizeText(CONFIG.STATIC_DRIVE_FOLDER_ID)) {
      return {
        ok: false,
        code: 'inaccessible-static',
        message: 'Folder Google Drive pada STATIC_DRIVE_FOLDER_ID tidak ditemukan atau Anda tidak punya akses. Perbarui ID tersebut di Config.gs.'
      };
    }

    try {
      var props = PropertiesService.getScriptProperties();
      var folderName = 'SIMOD HPS Documents';
      var createdFolder = DriveApp.createFolder(folderName);
      props.setProperty('DRIVE_FOLDER_ID', createdFolder.getId());

      return {
        ok: true,
        folder: createdFolder,
        driveFolderId: createdFolder.getId(),
        recovered: true,
        message: 'Folder Google Drive root lama tidak bisa diakses. Sistem membuat folder root baru secara otomatis.'
      };
    } catch (createErr) {
      return {
        ok: false,
        code: 'inaccessible-dynamic',
        message: 'Folder Google Drive root tidak bisa diakses dan folder pengganti gagal dibuat. Periksa izin Drive akun yang menjalankan web app.'
      };
    }
  }
}

function ensureDriveRootFolder() {
  var state = getDriveFolderState();
  if (!state.ok) throw new Error(state.message);
  return state.folder;
}

function getEventSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.EVENT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.EVENT_SHEET_NAME);
  return sheet;
}

function getHpsSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.HPS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.HPS_SHEET_NAME);
  return sheet;
}

function ensureSheetHeaders(sheet, headers) {
  var range = sheet.getRange(1, 1, 1, headers.length);
  var current = range.getValues()[0];
  var empty = current.join('').trim() === '';
  var mismatch = current.some(function (value, idx) {
    return (value || '').toString().trim() !== headers[idx];
  });

  if (empty || mismatch) {
    range.setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
}

function setupSheets() {
  ensureSheetHeaders(getEventSheet(), CONFIG.EVENT_HEADERS);
  ensureSheetHeaders(getHpsSheet(), CONFIG.HPS_HEADERS);
}

function listEducationEvents() {
  var sheet = getEventSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, CONFIG.EVENT_HEADERS.length).getValues();
  return values
    .map(mapEventRow)
    .filter(function (evt) {
      return evt.status !== 'ARCHIVED';
    })
    .sort(function (a, b) {
      return a.eventName.localeCompare(b.eventName);
    });
}

function getEventById(eventId) {
  var events = listEducationEvents();
  for (var i = 0; i < events.length; i++) {
    if (events[i].eventId === eventId) return events[i];
  }
  return null;
}

function listHpsPackages(filters) {
  setupSheets();
  filters = filters || {};

  var sheet = getHpsSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, CONFIG.HPS_HEADERS.length).getValues();
  var q = sanitizeText(filters.query).toLowerCase();
  var eventId = sanitizeText(filters.eventId);
  var status = sanitizeText(filters.status).toUpperCase();

  return values
    .map(mapHpsRow)
    .filter(function (pkg) {
      var eventMatch = eventId ? pkg.eventId === eventId : true;
      var statusMatch = status && status !== 'ALL' ? pkg.status === status : true;
      var haystack = [pkg.packageId, pkg.eventName, pkg.rupNumber, pkg.hpsName, pkg.noPesanan, pkg.createdBy].join(' ').toLowerCase();
      var queryMatch = q ? haystack.indexOf(q) !== -1 : true;
      return eventMatch && statusMatch && queryMatch;
    })
    .sort(function (a, b) {
      return new Date(b.createdAt) - new Date(a.createdAt);
    });
}

function addEducationRecord(eventName) {
  eventName = sanitizeText(eventName);
  if (!eventName) throw new Error('Event name is required');

  setupSheets();

  var existing = listEducationEvents();
  var dup = existing.some(function (evt) {
    return evt.eventName.toLowerCase() === eventName.toLowerCase();
  });
  if (dup) throw new Error('Event already exists');

  var now = new Date();
  var row = [
    buildId('EDU', now),
    eventName,
    Session.getActiveUser().getEmail() || 'unknown',
    now,
    now,
    'ACTIVE'
  ];

  getEventSheet().appendRow(row);

  return {
    ok: true,
    event: mapEventRow(row)
  };
}

function createHpsRecord(payload) {
  payload = payload || {};
  var eventId = sanitizeText(payload.eventId);
  var hpsName = sanitizeText(payload.hpsName);
  var rupNumber = sanitizeText(payload.rupNumber);
  var noPesanan = sanitizeText(payload.noPesanan);

  if (!eventId) throw new Error('eventId is required');
  if (!hpsName) throw new Error('HPS name is required');
  if (!noPesanan) throw new Error('No. Pesanan is required');

  setupSheets();

  var event = getEventById(eventId);
  if (!event) throw new Error('Event not found');

  var folderInfo = ensurePackageFolder(event.eventName, rupNumber, hpsName);
  var now = new Date();

  var row = [
    buildId('HPS', now),
    event.eventId,
    event.eventName,
    rupNumber,
    hpsName,
    '',
    '',
    noPesanan,
    '',
    '',
    '',
    '',
    '',
    folderInfo.folder.getId(),
    folderInfo.folder.getUrl(),
    Session.getActiveUser().getEmail() || 'unknown',
    now,
    'DRAFT',
    now
  ];

  getHpsSheet().appendRow(row);

  return {
    ok: true,
    hps: mapHpsRow(row)
  };
}

function uploadHpsFilesRecord(payload) {
  payload = payload || {};
  var packageId = sanitizeText(payload.packageId);
  var files = payload.files || {};

  if (!packageId) throw new Error('packageId is required');

  var hasAnyFile = Object.keys(CONFIG.FILE_COLUMNS).some(function (key) {
    return !!files[key];
  });
  if (!hasAnyFile) throw new Error('At least one file is required');

  setupSheets();

  var sheet = getHpsSheet();
  var data = sheet.getDataRange().getValues();
  var rowNumber = -1;
  var row = null;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === packageId) {
      rowNumber = i + 1;
      row = data[i];
      break;
    }
  }

  if (rowNumber === -1 || !row) throw new Error('HPS package not found');

  var packageFolder = getPackageFolder(row);

  Object.keys(CONFIG.FILE_COLUMNS).forEach(function (key) {
    if (!files[key]) return;

    validateFilePayload(files[key], CONFIG.FILE_COLUMNS[key].label);
    var uploaded = uploadFileToFolder(packageFolder, files[key], CONFIG.FILE_COLUMNS[key].prefix);
    row[CONFIG.FILE_COLUMNS[key].idIndex] = uploaded.fileId;
    row[CONFIG.FILE_COLUMNS[key].urlIndex] = uploaded.url;
  });

  row[17] = hasAllRequiredFiles(row) ? 'READY' : 'DRAFT';
  row[18] = new Date();

  sheet.getRange(rowNumber, 1, 1, CONFIG.HPS_HEADERS.length).setValues([row]);

  return {
    ok: true,
    hps: mapHpsRow(row)
  };
}

function getPackageFolder(row) {
  var existingFolderId = sanitizeText(row[13]);
  if (existingFolderId) {
    try {
      return DriveApp.getFolderById(existingFolderId);
    } catch (err) {
      // Ignore and recreate when folder id is invalid or access is gone.
    }
  }

  var created = ensurePackageFolder(row[2], row[3], row[4]).folder;
  row[13] = created.getId();
  row[14] = created.getUrl();
  return created;
}

function ensurePackageFolder(eventName, rupNumber, hpsName) {
  var root = ensureDriveRootFolder();
  var eventFolder = getOrCreateSubfolder(root, sanitizeFolderName(eventName));

  var packageName = rupNumber
    ? sanitizeFolderName(rupNumber + ' - ' + hpsName)
    : sanitizeFolderName(hpsName);

  var packageFolder = getOrCreateSubfolder(eventFolder, packageName);
  applyLinkSharingIfEnabled(root);
  applyLinkSharingIfEnabled(eventFolder);
  applyLinkSharingIfEnabled(packageFolder);

  return { folder: packageFolder };
}

function getOrCreateSubfolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function uploadFileToFolder(folder, filePayload, prefix) {
  var bytes = Utilities.base64Decode(filePayload.base64Data);
  var originalName = sanitizeFileName(filePayload.originalFileName);
  var fileName = sanitizeFileName(prefix + ' - ' + originalName);
  var blob = Utilities.newBlob(bytes, filePayload.mimeType, originalName);
  var file = folder.createFile(blob).setName(fileName);
  applyLinkSharingIfEnabled(file);

  return {
    fileId: file.getId(),
    url: file.getUrl()
  };
}

function applyLinkSharingIfEnabled(item) {
  if (!CONFIG.SHARE_WITH_LINK) return;
  try {
    item.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // Skip when policy blocks public sharing.
  }
}
