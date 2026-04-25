function createNotification_(type, payload) {
  payload = payload || {};
  var now = new Date();
  var row = [
    buildId('NOTIF', now),
    sanitizeText(type),
    sanitizeText(payload.audience).toUpperCase() || 'ADMIN',
    sanitizeText(payload.recipientEmail).toLowerCase(),
    sanitizeText(payload.packageId),
    sanitizeText(payload.eventId),
    sanitizeText(payload.eventName),
    sanitizeText(payload.hpsName),
    sanitizeText(payload.actorEmail) || 'unknown',
    sanitizeText(payload.message),
    'FALSE',
    now,
    ''
  ];

  getNotificationSheet().appendRow(row);
  return mapNotificationRow(row);
}

function listNotifications_(limit, options) {
  setupSheets();
  options = options || {};
  var sheet = getNotificationSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var rows = sheet.getRange(2, 1, lastRow - 1, CONFIG.NOTIFICATION_HEADERS.length).getValues();
  var normalizedLimit = Math.max(1, Math.min(Number(limit) || 20, 100));
  var audience = sanitizeText(options.audience).toUpperCase();
  var recipientEmail = sanitizeText(options.recipientEmail).toLowerCase();

  return rows
    .map(mapNotificationRow)
    .filter(function (item) {
      if (audience && item.audience !== audience) return false;
      if (recipientEmail && sanitizeText(item.recipientEmail).toLowerCase() !== recipientEmail) return false;
      return true;
    })
    .sort(function (a, b) {
      return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
    })
    .slice(0, normalizedLimit);
}

function markNotificationsRead_(notificationIds, options) {
  setupSheets();
  options = options || {};
  var sheet = getNotificationSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return {
      ok: true,
      updatedCount: 0
    };
  }

  var idMap = {};
  (notificationIds || []).forEach(function (id) {
    var normalized = sanitizeText(id);
    if (normalized) idMap[normalized] = true;
  });

  var markAll = !Object.keys(idMap).length;
  var now = new Date();
  var updatedCount = 0;
  var audience = sanitizeText(options.audience).toUpperCase();
  var recipientEmail = sanitizeText(options.recipientEmail).toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var notificationId = sanitizeText(row[0]);
    var isRead = sanitizeText(row[10]).toUpperCase() === 'TRUE';
    var rowAudience = sanitizeText(row[2]).toUpperCase() || 'ADMIN';
    var rowRecipientEmail = sanitizeText(row[3]).toLowerCase();
    if (isRead) continue;
    if (audience && rowAudience !== audience) continue;
    if (recipientEmail && rowRecipientEmail !== recipientEmail) continue;
    if (!markAll && !idMap[notificationId]) continue;

    row[10] = 'TRUE';
    row[12] = now;
    sheet.getRange(i + 1, 1, 1, CONFIG.NOTIFICATION_HEADERS.length).setValues([row]);
    updatedCount += 1;
  }

  return {
    ok: true,
    updatedCount: updatedCount
  };
}

function mapNotificationRow(row) {
  return {
    notificationId: row[0],
    type: row[1],
    audience: sanitizeText(row[2]).toUpperCase() || 'ADMIN',
    recipientEmail: row[3],
    packageId: row[4],
    eventId: row[5],
    eventName: row[6],
    hpsName: row[7],
    actorEmail: row[8],
    message: row[9],
    isRead: sanitizeText(row[10]).toUpperCase() === 'TRUE',
    createdAt: toIsoString(row[11]),
    readAt: toIsoString(row[12])
  };
}
