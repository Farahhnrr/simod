function createSessionToken_() {
  return Utilities.getUuid() + '-' + new Date().getTime();
}

function getAccessSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ACCESS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.ACCESS_SHEET_NAME);
  return sheet;
}

function authenticateEmail(email, displayName) {
  var normalizedEmail = sanitizeText(email).toLowerCase();
  var normalizedName = sanitizeText(displayName);

  if (!normalizedEmail) throw new Error('Email wajib diisi.');
  if (!isValidEmail_(normalizedEmail)) throw new Error('Format email tidak valid.');

  setupSheets();

  var record = findAccessRecordByEmail_(normalizedEmail);
  var status = record ? sanitizeText(record.record.status).toUpperCase() : 'NOT_REQUESTED';
  var sessionToken = createSessionToken_();
  var sessionPayload = {
    email: normalizedEmail,
    name: normalizedName,
    status: status,
    approved: status === 'APPROVED',
    isAdmin: false
  };

  CacheService.getScriptCache().put(
    sessionToken,
    JSON.stringify(sessionPayload),
    CONFIG.SESSION_TTL_SECONDS || 21600
  );

  if (record) {
    touchAccessRecord_(record.rowNumber, normalizedEmail, normalizedName);
  }

  return {
    ok: true,
    sessionToken: sessionToken,
    profile: {
      email: normalizedEmail,
      name: normalizedName
    },
    access: buildAccessState_(status, false)
  };
}

function authenticateAdmin(adminCode) {
  var code = sanitizeText(adminCode);
  var expected = sanitizeText(CONFIG.ADMIN_ACCESS_CODE);

  if (!expected) throw new Error('ADMIN_ACCESS_CODE belum diisi di Config.gs.');
  if (!code) throw new Error('Kode admin wajib diisi.');
  if (code !== expected) throw new Error('Kode admin tidak valid.');

  var sessionToken = createSessionToken_();
  var sessionPayload = {
    email: 'admin',
    name: 'Administrator',
    status: 'APPROVED',
    approved: true,
    isAdmin: true
  };

  CacheService.getScriptCache().put(
    sessionToken,
    JSON.stringify(sessionPayload),
    CONFIG.SESSION_TTL_SECONDS || 21600
  );

  return {
    ok: true,
    sessionToken: sessionToken,
    profile: {
      email: 'admin',
      name: 'Administrator'
    },
    access: buildAccessState_('APPROVED', true)
  };
}

function getSessionState(sessionToken) {
  var session = requireSession_(sessionToken);
  return {
    ok: true,
    profile: {
      email: session.email || '',
      name: session.name || ''
    },
    access: buildAccessState_(session.status || 'NOT_REQUESTED', !!session.isAdmin)
  };
}

function submitAccessRequest(sessionToken) {
  var session = requireSession_(sessionToken);
  if (!session.email) throw new Error('Sesi email tidak valid. Masukkan email terlebih dahulu.');

  setupSheets();
  var record = findAccessRecordByEmail_(session.email);
  var now = new Date();
  var status = 'PENDING';

  if (!record) {
    getAccessSheet().appendRow([
      buildId('REQ', now),
      session.email,
      session.name,
      status,
      now,
      '',
      '',
      now
    ]);
  } else {
    var row = record.row;
    row[1] = session.email;
    row[2] = session.name;
    row[3] = status;
    row[4] = now;
    row[5] = '';
    row[6] = '';
    row[7] = now;
    getAccessSheet().getRange(record.rowNumber, 1, 1, CONFIG.ACCESS_HEADERS.length).setValues([row]);
  }

  updateSessionStatus_(sessionToken, {
    status: status,
    approved: false
  });

  return {
    ok: true,
    status: status,
    message: 'Permintaan akses berhasil dikirim.'
  };
}

function requireSession_(sessionToken) {
  var token = sanitizeText(sessionToken);
  if (!token) throw new Error('Sesi tidak ditemukan. Masuk ulang.');

  var cached = CacheService.getScriptCache().get(token);
  if (!cached) throw new Error('Sesi sudah berakhir. Masuk ulang.');

  return JSON.parse(cached);
}

function requireApprovedSession_(sessionToken) {
  var session = requireSession_(sessionToken);
  if (!session.approved) {
    throw new Error('Akses belum disetujui. Kirim permintaan akses terlebih dahulu.');
  }
  return session;
}

function requireAdminSession_(sessionToken) {
  var session = requireSession_(sessionToken);
  if (!session.isAdmin) {
    throw new Error('Anda tidak memiliki akses admin.');
  }
  return session;
}

function updateSessionStatus_(sessionToken, nextFields) {
  var session = requireSession_(sessionToken);
  Object.keys(nextFields || {}).forEach(function (key) {
    session[key] = nextFields[key];
  });
  CacheService.getScriptCache().put(
    sessionToken,
    JSON.stringify(session),
    CONFIG.SESSION_TTL_SECONDS || 21600
  );
  return session;
}

function listAccessRecords_() {
  var sheet = getAccessSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var rows = sheet.getRange(2, 1, lastRow - 1, CONFIG.ACCESS_HEADERS.length).getValues();
  return rows
    .map(mapAccessRow)
    .sort(function (a, b) {
      return new Date(b.requestedAt || 0) - new Date(a.requestedAt || 0);
    });
}

function reviewAccessRequest(sessionToken, targetEmail, decision) {
  requireAdminSession_(sessionToken);

  var normalizedEmail = sanitizeText(targetEmail).toLowerCase();
  var nextStatus = sanitizeText(decision).toUpperCase();

  if (!normalizedEmail) throw new Error('Email target wajib diisi.');
  if (!isValidEmail_(normalizedEmail)) throw new Error('Format email target tidak valid.');
  if (['APPROVED', 'DENIED'].indexOf(nextStatus) === -1) {
    throw new Error('Keputusan akses tidak valid.');
  }

  setupSheets();
  var record = findAccessRecordByEmail_(normalizedEmail);
  if (!record) throw new Error('Permintaan akses tidak ditemukan.');

  var row = record.row;
  row[3] = nextStatus;
  row[5] = new Date();
  row[6] = 'Administrator';
  row[7] = new Date();
  getAccessSheet().getRange(record.rowNumber, 1, 1, CONFIG.ACCESS_HEADERS.length).setValues([row]);

  return {
    ok: true,
    accessRecord: mapAccessRow(row)
  };
}

function deleteAccessRecord(sessionToken, targetEmail) {
  requireAdminSession_(sessionToken);

  var normalizedEmail = sanitizeText(targetEmail).toLowerCase();
  if (!normalizedEmail) throw new Error('Email target wajib diisi.');
  if (!isValidEmail_(normalizedEmail)) throw new Error('Format email target tidak valid.');

  setupSheets();
  var record = findAccessRecordByEmail_(normalizedEmail);
  if (!record) throw new Error('Data akses tidak ditemukan.');

  getAccessSheet().deleteRow(record.rowNumber);

  return {
    ok: true,
    email: normalizedEmail
  };
}

function findAccessRecordByEmail_(email) {
  var sheet = getAccessSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  var rows = sheet.getRange(2, 1, lastRow - 1, CONFIG.ACCESS_HEADERS.length).getValues();
  var normalizedEmail = sanitizeText(email).toLowerCase();

  for (var i = 0; i < rows.length; i++) {
    var rowEmail = sanitizeText(rows[i][1]).toLowerCase();
    if (normalizedEmail && rowEmail === normalizedEmail) {
      return {
        rowNumber: i + 2,
        row: rows[i],
        record: mapAccessRow(rows[i])
      };
    }
  }

  return null;
}

function touchAccessRecord_(rowNumber, email, displayName) {
  var row = getAccessSheet().getRange(rowNumber, 1, 1, CONFIG.ACCESS_HEADERS.length).getValues()[0];
  row[1] = email;
  if (displayName) row[2] = displayName;
  row[7] = new Date();
  getAccessSheet().getRange(rowNumber, 1, 1, CONFIG.ACCESS_HEADERS.length).setValues([row]);
}

function buildAccessState_(status, isAdmin) {
  var normalizedStatus = sanitizeText(status).toUpperCase() || 'NOT_REQUESTED';
  return {
    status: normalizedStatus,
    approved: normalizedStatus === 'APPROVED' || !!isAdmin,
    isAdmin: !!isAdmin,
    canRequest: !isAdmin && normalizedStatus !== 'APPROVED' && normalizedStatus !== 'PENDING'
  };
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(sanitizeText(email));
}

function mapAccessRow(row) {
  return {
    requestId: row[0],
    email: row[1],
    displayName: row[2],
    status: sanitizeText(row[3]).toUpperCase(),
    requestedAt: toIsoString(row[4]),
    reviewedAt: toIsoString(row[5]),
    reviewedBy: row[6],
    lastAuthenticatedAt: toIsoString(row[7])
  };
}
