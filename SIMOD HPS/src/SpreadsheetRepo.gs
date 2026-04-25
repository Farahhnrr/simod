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
  if (!value) throw new Error('Nilai konfigurasi tidak ditemukan: ' + key + '.');
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

function getNotificationSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NOTIFICATION_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.NOTIFICATION_SHEET_NAME);
  return sheet;
}

function getEmailLogSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.EMAIL_LOG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.EMAIL_LOG_SHEET_NAME);
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
  ensureSheetHeaders(getAccessSheet(), CONFIG.ACCESS_HEADERS);
  ensureSheetHeaders(getNotificationSheet(), CONFIG.NOTIFICATION_HEADERS);
  ensureSheetHeaders(getEmailLogSheet(), CONFIG.EMAIL_LOG_HEADERS);
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
      var timeA = new Date((a && a.updatedAt) || (a && a.createdAt) || 0).getTime();
      var timeB = new Date((b && b.updatedAt) || (b && b.createdAt) || 0).getTime();
      return timeB - timeA;
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

function addEducationRecord(eventName, actorEmail) {
  eventName = sanitizeText(eventName);
  if (!eventName) throw new Error('Nama pendidikan wajib diisi.');

  setupSheets();

  var existing = listEducationEvents();
  var dup = existing.some(function (evt) {
    return evt.eventName.toLowerCase() === eventName.toLowerCase();
  });
  if (dup) throw new Error('Pendidikan sudah ada.');

  var now = new Date();
  var row = [
    buildId('EDU', now),
    eventName,
    sanitizeText(actorEmail) || 'unknown',
    now,
    now,
    'PROSES'
  ];

  getEventSheet().appendRow(row);
  sendNewEducationAvailableEmails_(row, actorEmail);

  return {
    ok: true,
    event: mapEventRow(row)
  };
}

function updateEducationStatusRecord(eventId, nextStatus, actorEmail) {
  eventId = sanitizeText(eventId);
  nextStatus = normalizeEventStatus(nextStatus);

  if (!eventId) throw new Error('Pendidikan wajib dipilih.');
  if (nextStatus !== 'PROSES' && nextStatus !== 'SELESAI') {
    throw new Error('Status pendidikan tidak valid.');
  }

  setupSheets();

  var found = findEventRowByEventId_(eventId);
  if (!found.row) throw new Error('Pendidikan tidak ditemukan.');

  var row = found.row;
  row[4] = new Date();
  row[5] = nextStatus;

  found.sheet.getRange(found.rowNumber, 1, 1, CONFIG.EVENT_HEADERS.length).setValues([row]);

  return {
    ok: true,
    event: mapEventRow(row),
    updatedBy: sanitizeText(actorEmail) || 'unknown'
  };
}

function deleteEducationRecord(eventId, actorEmail) {
  eventId = sanitizeText(eventId);
  if (!eventId) throw new Error('Pendidikan wajib dipilih.');

  setupSheets();

  var found = findEventRowByEventId_(eventId);
  if (!found.row) throw new Error('Pendidikan tidak ditemukan.');

  var linkedPackages = listHpsPackages({ eventId: eventId });
  if (linkedPackages.length) {
    throw new Error('Pendidikan tidak bisa dihapus karena masih memiliki ' + linkedPackages.length + ' HPS.');
  }

  var row = found.row;
  row[4] = new Date();
  row[5] = 'ARCHIVED';

  found.sheet.getRange(found.rowNumber, 1, 1, CONFIG.EVENT_HEADERS.length).setValues([row]);

  return {
    ok: true,
    eventId: eventId,
    archivedBy: sanitizeText(actorEmail) || 'unknown'
  };
}

function createHpsRecord(payload, actorEmail) {
  payload = payload || {};
  var eventId = sanitizeText(payload.eventId);
  var hpsName = sanitizeText(payload.hpsName);
  var rupNumber = sanitizeText(payload.rupNumber);

  if (!eventId) throw new Error('Pendidikan wajib dipilih.');
  if (!hpsName) throw new Error('Nama HPS wajib diisi.');
  if (!rupNumber) throw new Error('No. RUP wajib diisi.');

  setupSheets();

  var event = getEventById(eventId);
  if (!event) throw new Error('Pendidikan tidak ditemukan.');

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
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    folderInfo.folder.getId(),
    folderInfo.folder.getUrl(),
    sanitizeText(actorEmail) || 'unknown',
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

function updateHpsRecord(payload) {
  payload = payload || {};
  var packageId = sanitizeText(payload.packageId);
  var hpsName = sanitizeText(payload.hpsName);
  var rupNumber = sanitizeText(payload.rupNumber);

  if (!packageId) throw new Error('Paket HPS wajib dipilih.');
  if (!hpsName) throw new Error('Nama HPS wajib diisi.');
  if (!rupNumber) throw new Error('No. RUP wajib diisi.');

  setupSheets();

  var found = findHpsRowByPackageId_(packageId);
  if (!found.row) throw new Error('Paket HPS tidak ditemukan.');

  var row = found.row;
  row[3] = rupNumber;
  row[4] = hpsName;

  var packageFolder = getPackageFolder(row);
  packageFolder.setName(buildPackageFolderName_(rupNumber, hpsName));
  applyLinkSharingIfEnabled(packageFolder);
  row[17] = packageFolder.getId();
  row[18] = packageFolder.getUrl();
  row[22] = new Date();

  found.sheet.getRange(found.rowNumber, 1, 1, CONFIG.HPS_HEADERS.length).setValues([row]);

  return {
    ok: true,
    hps: mapHpsRow(row)
  };
}

function updateAdminRestrictedRecord(payload, actorEmail) {
  payload = payload || {};
  var packageId = sanitizeText(payload.packageId);
  var files = payload.files || {};
  var hasAnyFile = Object.keys(files).some(function (key) {
    return !!files[key];
  });
  var hasNoPesananField = Object.prototype.hasOwnProperty.call(payload, 'noPesanan');
  var noPesanan = sanitizeText(payload.noPesanan);

  if (!packageId) throw new Error('Paket HPS wajib dipilih.');
  if (!hasAnyFile && !hasNoPesananField) {
    throw new Error('Tidak ada perubahan admin yang dikirim.');
  }

  setupSheets();

  var found = findHpsRowByPackageId_(packageId);
  if (!found.row) throw new Error('Paket HPS tidak ditemukan.');

  var row = found.row;
  var previousStatus = sanitizeText(row[21]).toUpperCase();
  if (hasNoPesananField) {
    row[7] = noPesanan;
  }

  if (hasAnyFile) {
    var packageFolder = getPackageFolder(row);

    Object.keys(files).forEach(function (key) {
      if (!files[key]) return;

      var columnConfig = CONFIG.FILE_COLUMNS[key];
      var previousFileId = sanitizeText(row[columnConfig.idIndex]);

      validateFilePayload(files[key], columnConfig.label);
      var uploaded = uploadFileToFolder(packageFolder, files[key], columnConfig.prefix);
      row[columnConfig.idIndex] = uploaded.fileId;
      row[columnConfig.urlIndex] = uploaded.url;
      trashDriveFileById_(previousFileId);
    });
  }

  row[21] = hasAllRequiredFiles(row) ? 'READY' : 'DRAFT';
  row[22] = new Date();

  found.sheet.getRange(found.rowNumber, 1, 1, CONFIG.HPS_HEADERS.length).setValues([row]);
  createReadyNotificationIfNeeded_(row, previousStatus, actorEmail);

  return {
    ok: true,
    hps: mapHpsRow(row)
  };
}

function deleteHpsDocumentRecord(packageId, fileKey) {
  packageId = sanitizeText(packageId);
  fileKey = sanitizeText(fileKey);

  if (!packageId) throw new Error('Paket HPS wajib dipilih.');
  if (!fileKey || !CONFIG.FILE_COLUMNS[fileKey]) {
    throw new Error('Jenis dokumen tidak valid.');
  }

  setupSheets();

  var found = findHpsRowByPackageId_(packageId);
  if (!found.row) throw new Error('Paket HPS tidak ditemukan.');

  var row = found.row;
  var columnConfig = CONFIG.FILE_COLUMNS[fileKey];
  var previousFileId = sanitizeText(row[columnConfig.idIndex]);
  var previousFileUrl = sanitizeText(row[columnConfig.urlIndex]);

  if (!previousFileId && !previousFileUrl) {
    throw new Error('Dokumen ' + columnConfig.label + ' belum ada.');
  }

  trashDriveFileById_(previousFileId);
  row[columnConfig.idIndex] = '';
  row[columnConfig.urlIndex] = '';
  row[21] = hasAllRequiredFiles(row) ? 'READY' : 'DRAFT';
  row[22] = new Date();

  found.sheet.getRange(found.rowNumber, 1, 1, CONFIG.HPS_HEADERS.length).setValues([row]);

  return {
    ok: true,
    hps: mapHpsRow(row),
    deletedFileKey: fileKey
  };
}

function uploadHpsFilesRecord(payload, actorEmail) {
  payload = payload || {};
  var packageId = sanitizeText(payload.packageId);
  var files = payload.files || {};

  if (!packageId) throw new Error('Paket HPS wajib dipilih.');

  var hasAnyFile = Object.keys(CONFIG.FILE_COLUMNS).some(function (key) {
    return !!files[key];
  });
  if (!hasAnyFile) throw new Error('Pilih minimal satu dokumen untuk diunggah.');

  setupSheets();

  var found = findHpsRowByPackageId_(packageId);
  if (!found.row) throw new Error('Paket HPS tidak ditemukan.');

  var row = found.row;
  var packageFolder = getPackageFolder(row);
  var previousStatus = sanitizeText(row[21]).toUpperCase();
  if (!sanitizeText(row[19]) && sanitizeText(actorEmail) && !isAdminEmail_(actorEmail)) {
    row[19] = sanitizeText(actorEmail).toLowerCase();
  }

  Object.keys(CONFIG.FILE_COLUMNS).forEach(function (key) {
    if (!files[key]) return;

    var columnConfig = CONFIG.FILE_COLUMNS[key];
    var previousFileId = sanitizeText(row[columnConfig.idIndex]);

    validateFilePayload(files[key], columnConfig.label);
    var uploaded = uploadFileToFolder(packageFolder, files[key], columnConfig.prefix);
    row[columnConfig.idIndex] = uploaded.fileId;
    row[columnConfig.urlIndex] = uploaded.url;
    trashDriveFileById_(previousFileId);
  });

  row[21] = hasAllRequiredFiles(row) ? 'READY' : 'DRAFT';
  row[22] = new Date();

  found.sheet.getRange(found.rowNumber, 1, 1, CONFIG.HPS_HEADERS.length).setValues([row]);

  if (files.hps) {
    createNotification_('USER_HPS_UPLOAD', {
      audience: 'ADMIN',
      packageId: row[0],
      eventId: row[1],
      eventName: row[2],
      hpsName: row[4],
      actorEmail: sanitizeText(actorEmail) || sanitizeText(row[19]) || 'unknown',
      message: 'Pengguna mengunggah Dokumen HPS.'
    });
  }
  createReadyNotificationIfNeeded_(row, previousStatus, sanitizeText(actorEmail) || sanitizeText(row[19]) || 'unknown');

  return {
    ok: true,
    hps: mapHpsRow(row)
  };
}

function createReadyNotificationIfNeeded_(row, previousStatus, actorEmail) {
  var nextStatus = sanitizeText(row[21]).toUpperCase();
  var normalizedPreviousStatus = sanitizeText(previousStatus).toUpperCase();
  if (nextStatus !== 'READY' || normalizedPreviousStatus === 'READY') return;

  var recipientEmail = resolveNotificationRecipientEmail_(row, actorEmail);
  if (!recipientEmail) return;

  createNotification_('USER_HPS_READY', {
    audience: 'USER',
    recipientEmail: recipientEmail,
    packageId: row[0],
    eventId: row[1],
    eventName: row[2],
    hpsName: row[4],
    actorEmail: sanitizeText(actorEmail) || 'system',
    message: 'HPS "' + sanitizeText(row[4]) + '" sudah berstatus Siap.'
  });
  sendReadyStatusEmail_(row, recipientEmail, actorEmail);
}

function resolveNotificationRecipientEmail_(row, actorEmail) {
  var ownerEmail = sanitizeText(row[19]).toLowerCase();
  if (ownerEmail) return ownerEmail;

  var normalizedActorEmail = sanitizeText(actorEmail).toLowerCase();
  if (normalizedActorEmail && !isAdminEmail_(normalizedActorEmail)) {
    return normalizedActorEmail;
  }

  var packageId = sanitizeText(row[0]);
  if (!packageId) return '';

  return findLatestUserUploadActorByPackageId_(packageId);
}

function findLatestUserUploadActorByPackageId_(packageId) {
  setupSheets();
  var sheet = getNotificationSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return '';

  var rows = sheet.getRange(2, 1, lastRow - 1, CONFIG.NOTIFICATION_HEADERS.length).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    var row = rows[i];
    var type = sanitizeText(row[1]).toUpperCase();
    var rowPackageId = sanitizeText(row[4]);
    var actorEmail = sanitizeText(row[8]).toLowerCase();
    if (type !== 'USER_HPS_UPLOAD') continue;
    if (rowPackageId !== packageId) continue;
    if (!actorEmail || isAdminEmail_(actorEmail)) continue;
    return actorEmail;
  }
  return '';
}

function isAdminEmail_(email) {
  var normalizedEmail = sanitizeText(email).toLowerCase();
  return (CONFIG.ADMIN_ALLOWED_EMAILS || []).some(function (candidate) {
    return sanitizeText(candidate).toLowerCase() === normalizedEmail;
  });
}

function findHpsRowByPackageId_(packageId) {
  var sheet = getHpsSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === packageId) {
      return {
        sheet: sheet,
        rowNumber: i + 1,
        row: data[i]
      };
    }
  }

  return {
    sheet: sheet,
    rowNumber: -1,
    row: null
  };
}

function findEventRowByEventId_(eventId) {
  var sheet = getEventSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === eventId) {
      return {
        sheet: sheet,
        rowNumber: i + 1,
        row: data[i]
      };
    }
  }

  return {
    sheet: sheet,
    rowNumber: -1,
    row: null
  };
}

function getPackageFolder(row) {
  var existingFolderId = sanitizeText(row[17]);
  if (existingFolderId) {
    try {
      return DriveApp.getFolderById(existingFolderId);
    } catch (err) {
      // Ignore and recreate when folder id is invalid or access is gone.
    }
  }

  var created = ensurePackageFolder(row[2], row[3], row[4]).folder;
  row[17] = created.getId();
  row[18] = created.getUrl();
  return created;
}

function ensurePackageFolder(eventName, rupNumber, hpsName) {
  var root = ensureDriveRootFolder();
  var eventFolder = getOrCreateSubfolder(root, sanitizeFolderName(eventName));
  var packageName = buildPackageFolderName_(rupNumber, hpsName);
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

function buildPackageFolderName_(rupNumber, hpsName) {
  return rupNumber
    ? sanitizeFolderName(rupNumber + ' - ' + hpsName)
    : sanitizeFolderName(hpsName);
}

function trashDriveFileById_(fileId) {
  if (!fileId) return;
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (err) {
    // Abaikan jika file lama sudah tidak ada atau tidak bisa dihapus.
  }
}

function applyLinkSharingIfEnabled(item) {
  if (!CONFIG.SHARE_WITH_LINK) return;
  try {
    item.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // Skip when policy blocks public sharing.
  }
}
