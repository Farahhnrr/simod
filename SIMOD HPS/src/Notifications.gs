function createNotification_(type, payload) {
  payload = payload || {};
  var now = new Date();
  var row = [
    buildId('NOTIF', now),
    sanitizeText(type),
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

function listNotifications_(limit) {
  setupSheets();
  var sheet = getNotificationSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var rows = sheet.getRange(2, 1, lastRow - 1, CONFIG.NOTIFICATION_HEADERS.length).getValues();
  var normalizedLimit = Math.max(1, Math.min(Number(limit) || 20, 100));

  return rows
    .map(mapNotificationRow)
    .sort(function (a, b) {
      return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
    })
    .slice(0, normalizedLimit);
}

function markNotificationsRead_(notificationIds) {
  setupSheets();
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

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var notificationId = sanitizeText(row[0]);
    var isRead = sanitizeText(row[8]).toUpperCase() === 'TRUE';
    if (isRead) continue;
    if (!markAll && !idMap[notificationId]) continue;

    row[8] = 'TRUE';
    row[10] = now;
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
    packageId: row[2],
    eventId: row[3],
    eventName: row[4],
    hpsName: row[5],
    actorEmail: row[6],
    message: row[7],
    isRead: sanitizeText(row[8]).toUpperCase() === 'TRUE',
    createdAt: toIsoString(row[9]),
    readAt: toIsoString(row[10])
  };
}
