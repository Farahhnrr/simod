function sanitizeText(value) {
  return (value || '').toString().trim();
}

function sanitizeFileName(name) {
  return (name || 'tanpa-nama')
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function sanitizeFolderName(name) {
  return sanitizeFileName(name || 'folder-tanpa-nama').substring(0, 120);
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
  if (!filePayload) throw new Error('File ' + label + ' wajib diisi.');
  if (!filePayload.base64Data) throw new Error('Data base64 untuk file ' + label + ' wajib ada.');
  if (!filePayload.mimeType) throw new Error('Tipe MIME untuk file ' + label + ' wajib ada.');
  if (!filePayload.originalFileName) throw new Error('Nama file asli untuk file ' + label + ' wajib ada.');
}

function hasAllRequiredFiles(row) {
  return !!(
    sanitizeText(row[5]) &&
    sanitizeText(row[9]) &&
    sanitizeText(row[11]) &&
    sanitizeText(row[13]) &&
    sanitizeText(row[15])
  );
}

function normalizeEventStatus(value) {
  var status = sanitizeText(value).toUpperCase();
  if (!status || status === 'ACTIVE') return 'PROSES';
  if (status === 'PROSES' || status === 'SELESAI' || status === 'ARCHIVED') return status;
  return status;
}

function mapEventRow(row) {
  return {
    eventId: row[0],
    eventName: row[1],
    createdBy: row[2],
    createdAt: toIsoString(row[3]),
    updatedAt: toIsoString(row[4]),
    status: normalizeEventStatus(row[5])
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
    hpsLinkInaprocFileId: row[9],
    hpsLinkInaprocFileUrl: row[10],
    eFakturFileId: row[11],
    eFakturFileUrl: row[12],
    suratPesananFileId: row[13],
    suratPesananFileUrl: row[14],
    bastFileId: row[15],
    bastFileUrl: row[16],
    packageFolderId: row[17],
    packageFolderUrl: row[18],
    createdBy: row[19],
    createdAt: toIsoString(row[20]),
    status: row[21],
    updatedAt: toIsoString(row[22]),
    filesReady: hasAllRequiredFiles(row)
  };
}
