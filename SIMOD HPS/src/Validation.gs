function sanitizeText(value) {
  return (value || '').toString().trim();
}

function sanitizeFileName(name) {
  return (name || 'untitled')
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function sanitizeFolderName(name) {
  return sanitizeFileName(name || 'untitled-folder').substring(0, 120);
}

function toIsoString(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value.toISOString();
  }
  var maybeDate = new Date(value);
  return isNaN(maybeDate.getTime()) ? String(value) : maybeDate.toISOString();
}

function buildId(prefix, now) {
  return prefix + '-' +
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') +
    '-' + Math.floor(Math.random() * 1000);
}

function validateFilePayload(filePayload, label) {
  if (!filePayload) throw new Error(label + ' file is required');
  if (!filePayload.base64Data) throw new Error(label + ' base64Data is required');
  if (!filePayload.mimeType) throw new Error(label + ' mimeType is required');
  if (!filePayload.originalFileName) throw new Error(label + ' originalFileName is required');
}

function hasAllRequiredFiles(row) {
  return !!(
    sanitizeText(row[5]) &&
    sanitizeText(row[7]) &&
    sanitizeText(row[11])
  );
}

function mapEventRow(row) {
  return {
    eventId: row[0],
    eventName: row[1],
    createdBy: row[2],
    createdAt: toIsoString(row[3]),
    updatedAt: toIsoString(row[4]),
    status: row[5]
  };
}

function mapHpsRow(row) {
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
    createdAt: toIsoString(row[16]),
    status: row[17],
    updatedAt: toIsoString(row[18]),
    filesReady: hasAllRequiredFiles(row)
  };
}