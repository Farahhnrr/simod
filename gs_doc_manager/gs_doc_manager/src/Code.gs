var EVENT_SHEET_NAME = 'Pendidikan_Events';
var HPS_SHEET_NAME = 'HPS_Packages';
var STATIC_SHEET_ID = '1d4nxj-G9N2pR_4hkmhNwTdsBeO5O_fE7GpYQQOekqdI';
var STATIC_DRIVE_FOLDER_ID = '1dykGZ6P5FbllcMoFutHWtphdVBIl7BY1';

var EVENT_HEADERS = [
  'EventId',
  'EventName',
  'CreatedBy',
  'CreatedAt',
  'UpdatedAt',
  'Status'
];

var HPS_HEADERS = [
  'PackageId',
  'EventId',
  'EventName',
  'RupNumber',
  'HpsName',
  'HpsFileId',
  'HpsFileUrl',
  'NoPesanan',
  'LegacySuratPesananFileUrl',
  'LegacySuratBastFileId',
  'LegacySuratBastFileUrl',
  'EFakturFileId',
  'EFakturFileUrl',
  'PackageFolderId',
  'PackageFolderUrl',
  'CreatedBy',
  'CreatedAt',
  'Status',
  'UpdatedAt'
];

var FILE_COLUMNS = {
  hps: { idIndex: 5, urlIndex: 6, label: 'HPS', prefix: 'HPS' },
  eFaktur: { idIndex: 11, urlIndex: 12, label: 'E-Faktur', prefix: 'E-Faktur' }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Pendidikan Militer - HPS Document Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getBootstrapData() {
  var config = getProjectConfig_();
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
    events: [],
    packages: []
  };

  if (!config.configured) return response;

  SpreadsheetApp.openById(config.sheetId).getId();
  DriveApp.getFolderById(config.driveFolderId).getId();

  setupSheets_();

  var events = listEducationEvents_();
  var packages = listHpsPackages({});

  response.events = events;
  response.packages = packages;
  response.stats = getStats_(events, packages);
  return response;
}

function addEducation(eventName) {
  eventName = sanitizeText_(eventName);
  if (!eventName) throw new Error('Event name is required');

  setupSheets_();

  var existing = listEducationEvents_();
  var dup = existing.some(function (evt) {
    return evt.eventName.toLowerCase() === eventName.toLowerCase();
  });
  if (dup) throw new Error('Event already exists');

  var now = new Date();
  var row = [
    buildId_('EDU', now),
    eventName,
    Session.getActiveUser().getEmail() || 'unknown',
    now,
    now,
    'ACTIVE'
  ];

  getEventSheet_().appendRow(row);

  return {
    ok: true,
    event: mapEventRow_(row)
  };
}

function createHps(payload) {
  payload = payload || {};
  var eventId = sanitizeText_(payload.eventId);
  var hpsName = sanitizeText_(payload.hpsName);
  var rupNumber = sanitizeText_(payload.rupNumber);
  var noPesanan = sanitizeText_(payload.noPesanan);

  if (!eventId) throw new Error('eventId is required');
  if (!hpsName) throw new Error('HPS name is required');
  if (!noPesanan) throw new Error('No. Pesanan is required');

  setupSheets_();

  var event = getEventById_(eventId);
  if (!event) throw new Error('Event not found');

  var folderInfo = ensurePackageFolder_(event.eventName, rupNumber, hpsName);
  var now = new Date();

  var row = [
    buildId_('HPS', now),
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

  getHpsSheet_().appendRow(row);

  return {
    ok: true,
    hps: mapHpsRow_(row)
  };
}

function uploadHpsFiles(payload) {
  payload = payload || {};
  var packageId = sanitizeText_(payload.packageId);
  var files = payload.files || {};

  if (!packageId) throw new Error('packageId is required');

  var hasAnyFile = Object.keys(FILE_COLUMNS).some(function (key) {
    return !!files[key];
  });
  if (!hasAnyFile) throw new Error('At least one file is required');

  setupSheets_();

  var sheet = getHpsSheet_();
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

  var packageFolder = getPackageFolder_(row);

  Object.keys(FILE_COLUMNS).forEach(function (key) {
    if (!files[key]) return;

    validateFilePayload_(files[key], FILE_COLUMNS[key].label);
    var uploaded = uploadFileToFolder_(packageFolder, files[key], FILE_COLUMNS[key].prefix);
    row[FILE_COLUMNS[key].idIndex] = uploaded.fileId;
    row[FILE_COLUMNS[key].urlIndex] = uploaded.url;
  });

  row[17] = hasAllRequiredFiles_(row) ? 'READY' : 'DRAFT';
  row[18] = new Date();

  sheet.getRange(rowNumber, 1, 1, HPS_HEADERS.length).setValues([row]);

  return {
    ok: true,
    hps: mapHpsRow_(row)
  };
}

function listHpsPackages(filters) {
  setupSheets_();
  filters = filters || {};

  var sheet = getHpsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, HPS_HEADERS.length).getValues();
  var q = sanitizeText_(filters.query).toLowerCase();
  var eventId = sanitizeText_(filters.eventId);
  var status = sanitizeText_(filters.status).toUpperCase();

  return values
    .map(mapHpsRow_)
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

function setupSheets_() {
  ensureSheetHeaders_(getEventSheet_(), EVENT_HEADERS);
  ensureSheetHeaders_(getHpsSheet_(), HPS_HEADERS);
}

function ensureSheetHeaders_(sheet, headers) {
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

function listEducationEvents_() {
  var sheet = getEventSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, EVENT_HEADERS.length).getValues();
  return values
    .map(mapEventRow_)
    .filter(function (evt) {
      return evt.status !== 'ARCHIVED';
    })
    .sort(function (a, b) {
      return a.eventName.localeCompare(b.eventName);
    });
}

function getEventById_(eventId) {
  var events = listEducationEvents_();
  for (var i = 0; i < events.length; i++) {
    if (events[i].eventId === eventId) return events[i];
  }
  return null;
}

function getStats_(events, packages) {
  var ready = packages.filter(function (pkg) { return pkg.status === 'READY'; }).length;
  return {
    totalEvents: events.length,
    totalHps: packages.length,
    readyHps: ready,
    draftHps: packages.length - ready
  };
}

function getPackageFolder_(row) {
  var existingFolderId = sanitizeText_(row[13]);
  if (existingFolderId) {
    try {
      return DriveApp.getFolderById(existingFolderId);
    } catch (err) {
      // Fall through and recreate when folder id is invalid or missing access.
    }
  }

  var created = ensurePackageFolder_(row[2], row[3], row[4]).folder;
  row[13] = created.getId();
  row[14] = created.getUrl();
  return created;
}

function ensurePackageFolder_(eventName, rupNumber, hpsName) {
  var root = DriveApp.getFolderById(getRequiredProperty_('DRIVE_FOLDER_ID'));
  var eventFolder = getOrCreateSubfolder_(root, sanitizeFolderName_(eventName));

  var packageName = rupNumber
    ? sanitizeFolderName_(rupNumber + ' - ' + hpsName)
    : sanitizeFolderName_(hpsName);

  var packageFolder = getOrCreateSubfolder_(eventFolder, packageName);

  return { folder: packageFolder };
}

function getOrCreateSubfolder_(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function validateFilePayload_(filePayload, label) {
  if (!filePayload) throw new Error(label + ' file is required');
  if (!filePayload.base64Data) throw new Error(label + ' base64Data is required');
  if (!filePayload.mimeType) throw new Error(label + ' mimeType is required');
  if (!filePayload.originalFileName) throw new Error(label + ' originalFileName is required');
}

function uploadFileToFolder_(folder, filePayload, prefix) {
  var bytes = Utilities.base64Decode(filePayload.base64Data);
  var originalName = sanitizeFileName_(filePayload.originalFileName);
  var fileName = sanitizeFileName_(prefix + ' - ' + originalName);
  var blob = Utilities.newBlob(bytes, filePayload.mimeType, originalName);
  var file = folder.createFile(blob).setName(fileName);

  return {
    fileId: file.getId(),
    url: file.getUrl()
  };
}

function hasAllRequiredFiles_(row) {
  return !!(
    sanitizeText_(row[5]) &&
    sanitizeText_(row[7]) &&
    sanitizeText_(row[11])
  );
}

function getEventSheet_() {
  var ss = SpreadsheetApp.openById(getRequiredProperty_('SHEET_ID'));
  var sheet = ss.getSheetByName(EVENT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(EVENT_SHEET_NAME);
  return sheet;
}

function getHpsSheet_() {
  var ss = SpreadsheetApp.openById(getRequiredProperty_('SHEET_ID'));
  var sheet = ss.getSheetByName(HPS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(HPS_SHEET_NAME);
  return sheet;
}

function getRequiredProperty_(key) {
  var config = getProjectConfig_();
  var value = '';
  if (key === 'SHEET_ID') value = config.sheetId;
  if (key === 'DRIVE_FOLDER_ID') value = config.driveFolderId;
  if (!value) throw new Error('Missing configuration value: ' + key + '.');
  return value;
}

function getProjectConfig_() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = (STATIC_SHEET_ID || '').trim() || (props.getProperty('SHEET_ID') || '').trim();
  var driveFolderId = (STATIC_DRIVE_FOLDER_ID || '').trim() || (props.getProperty('DRIVE_FOLDER_ID') || '').trim();

  return {
    sheetId: sheetId,
    driveFolderId: driveFolderId,
    configured: !!(sheetId && driveFolderId)
  };
}

function mapEventRow_(row) {
  return {
    eventId: row[0],
    eventName: row[1],
    createdBy: row[2],
    createdAt: toIsoString_(row[3]),
    updatedAt: toIsoString_(row[4]),
    status: row[5]
  };
}

function mapHpsRow_(row) {
  return {
    packageId: row[0],
    eventId: row[1],
    eventName: row[2],
    rupNumber: row[3],
    hpsName: row[4],
    hpsFileId: row[5],
    hpsFileUrl: row[6],
    noPesanan: row[7],
    eFakturFileId: row[11],
    eFakturFileUrl: row[12],
    packageFolderId: row[13],
    packageFolderUrl: row[14],
    createdBy: row[15],
    createdAt: toIsoString_(row[16]),
    status: row[17],
    updatedAt: toIsoString_(row[18]),
    filesReady: hasAllRequiredFiles_(row)
  };
}

function buildId_(prefix, now) {
  return prefix + '-' +
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') +
    '-' + Math.floor(Math.random() * 1000);
}

function sanitizeText_(value) {
  return (value || '').toString().trim();
}

function sanitizeFolderName_(name) {
  return sanitizeFileName_(name || 'untitled-folder').substring(0, 120);
}

function sanitizeFileName_(name) {
  return (name || 'untitled')
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function toIsoString_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value.toISOString();
  }
  var maybeDate = new Date(value);
  return isNaN(maybeDate.getTime()) ? String(value) : maybeDate.toISOString();
}
