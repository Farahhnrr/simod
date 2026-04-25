function getAppBaseUrl_() {
  return ScriptApp.getService().getUrl() || '';
}

function listApprovedUserEmails_() {
  return listAccessRecords_()
    .filter(function (record) {
      var email = sanitizeText(record && record.email).toLowerCase();
      return sanitizeText(record && record.status).toUpperCase() === 'APPROVED'
        && !!email
        && !isAdminEmail_(email);
    })
    .map(function (record) {
      return sanitizeText(record.email).toLowerCase();
    })
    .filter(function (email, index, arr) {
      return arr.indexOf(email) === index;
    });
}

function logEmailAttempt_(type, recipientEmail, subject, status, errorMessage, payload) {
  payload = payload || {};
  setupSheets();

  getEmailLogSheet().appendRow([
    buildId('MAIL', new Date()),
    sanitizeText(type),
    sanitizeText(recipientEmail).toLowerCase(),
    sanitizeText(subject),
    sanitizeText(status).toUpperCase(),
    sanitizeText(errorMessage),
    sanitizeText(payload.eventId),
    sanitizeText(payload.eventName),
    sanitizeText(payload.packageId),
    sanitizeText(payload.hpsName),
    sanitizeText(payload.triggeredBy) || 'system',
    new Date()
  ]);
}

function sendEmailNotification_(type, recipientEmail, subject, body, htmlBody, payload) {
  payload = payload || {};
  var normalizedRecipient = sanitizeText(recipientEmail).toLowerCase();
  var normalizedSubject = sanitizeText(subject);

  if (!CONFIG.EMAIL_NOTIFICATIONS_ENABLED) {
    logEmailAttempt_(type, normalizedRecipient, normalizedSubject, 'SKIPPED_DISABLED', 'Email notifications disabled.', payload);
    return false;
  }

  if (!normalizedRecipient || !isValidEmail_(normalizedRecipient)) {
    logEmailAttempt_(type, normalizedRecipient, normalizedSubject, 'SKIPPED_NO_RECIPIENT', 'Recipient email missing or invalid.', payload);
    return false;
  }

  if (isAdminEmail_(normalizedRecipient)) {
    logEmailAttempt_(type, normalizedRecipient, normalizedSubject, 'SKIPPED_ADMIN', 'Recipient is admin email.', payload);
    return false;
  }

  var remainingQuota = 0;
  try {
    remainingQuota = MailApp.getRemainingDailyQuota();
  } catch (err) {
    logEmailAttempt_(type, normalizedRecipient, normalizedSubject, 'ERROR_QUOTA_CHECK', err && err.message ? err.message : String(err), payload);
    return false;
  }

  if (remainingQuota <= 0) {
    logEmailAttempt_(type, normalizedRecipient, normalizedSubject, 'SKIPPED_QUOTA', 'Remaining daily MailApp quota is zero.', payload);
    return false;
  }

  try {
    MailApp.sendEmail({
      to: normalizedRecipient,
      subject: normalizedSubject,
      body: sanitizeText(body),
      htmlBody: htmlBody,
      name: sanitizeText(CONFIG.APP_DISPLAY_NAME) || 'SIMOD HPS'
    });
    logEmailAttempt_(type, normalizedRecipient, normalizedSubject, 'SENT', '', payload);
    return true;
  } catch (err) {
    logEmailAttempt_(type, normalizedRecipient, normalizedSubject, 'ERROR_SEND', err && err.message ? err.message : String(err), payload);
    return false;
  }
}

function buildEmailFooterHtml_() {
  var appUrl = sanitizeText(getAppBaseUrl_());
  var linkHtml = appUrl
    ? '<p style="margin:16px 0 0;"><a href="' + htmlEscape_(appUrl) + '" style="display:inline-block;padding:10px 14px;border-radius:10px;background:#314537;color:#f8f7f2;text-decoration:none;font-weight:700;">Buka SIMOD HPS</a></p>'
    : '';

  return ''
    + '<p style="margin:16px 0 0;color:#778070;font-size:12px;">'
    + 'Email ini dikirim otomatis oleh ' + htmlEscape_(sanitizeText(CONFIG.APP_DISPLAY_NAME) || 'SIMOD HPS') + '.'
    + '</p>'
    + linkHtml;
}

function sendNewEducationAvailableEmails_(eventRow, actorEmail) {
  if (!CONFIG.EMAIL_NOTIFY_NEW_EDUCATION) return;

  var recipients = listApprovedUserEmails_();
  if (!recipients.length) {
    logEmailAttempt_(
      'NEW_EDUCATION_AVAILABLE',
      '',
      'Pendidikan baru tersedia di SIMOD HPS',
      'SKIPPED_NO_RECIPIENT',
      'No approved user emails available.',
      {
        eventId: sanitizeText(eventRow && eventRow[0]),
        eventName: sanitizeText(eventRow && eventRow[1]),
        triggeredBy: sanitizeText(actorEmail)
      }
    );
    return;
  }

  var eventId = sanitizeText(eventRow && eventRow[0]);
  var eventName = sanitizeText(eventRow && eventRow[1]);
  var appUrl = sanitizeText(getAppBaseUrl_());
  var subject = 'Pendidikan baru tersedia di SIMOD HPS';
  var textBody = [
    'Pendidikan baru tersedia di SIMOD HPS.',
    '',
    'Nama Pendidikan: ' + eventName,
    eventId ? 'ID Pendidikan: ' + eventId : '',
    appUrl ? 'Buka aplikasi: ' + appUrl : ''
  ].filter(Boolean).join('\n');
  var htmlBody = ''
    + '<div style="font-family:Arial,sans-serif;color:#171d17;line-height:1.6;">'
    + '<p style="margin:0 0 12px;">Pendidikan baru tersedia di <strong>SIMOD HPS</strong>.</p>'
    + '<div style="padding:14px 16px;border:1px solid #d8ddd2;border-radius:14px;background:#fbfaf6;">'
    + '<div style="font-size:12px;color:#778070;text-transform:uppercase;letter-spacing:.08em;">Nama Pendidikan</div>'
    + '<div style="margin-top:6px;font-size:16px;font-weight:700;">' + htmlEscape_(eventName) + '</div>'
    + (eventId ? '<div style="margin-top:8px;font-size:12px;color:#778070;">ID: ' + htmlEscape_(eventId) + '</div>' : '')
    + '</div>'
    + buildEmailFooterHtml_()
    + '</div>';

  recipients.forEach(function (recipientEmail) {
    sendEmailNotification_(
      'NEW_EDUCATION_AVAILABLE',
      recipientEmail,
      subject,
      textBody,
      htmlBody,
      {
        eventId: eventId,
        eventName: eventName,
        triggeredBy: sanitizeText(actorEmail)
      }
    );
  });
}

function sendReadyStatusEmail_(row, recipientEmail, actorEmail) {
  if (!CONFIG.EMAIL_NOTIFY_HPS_READY) return;

  var packageId = sanitizeText(row && row[0]);
  var eventId = sanitizeText(row && row[1]);
  var eventName = sanitizeText(row && row[2]);
  var hpsName = sanitizeText(row && row[4]);
  var noPesanan = sanitizeText(row && row[7]);
  var subject = 'HPS Anda sudah Siap';
  var appUrl = sanitizeText(getAppBaseUrl_());
  var textBody = [
    'HPS Anda sudah berstatus Siap.',
    '',
    'Nama Pendidikan: ' + eventName,
    'Nama HPS: ' + hpsName,
    noPesanan ? 'No. Surat Pesanan: ' + noPesanan : '',
    packageId ? 'ID HPS: ' + packageId : '',
    appUrl ? 'Buka aplikasi: ' + appUrl : ''
  ].filter(Boolean).join('\n');
  var htmlBody = ''
    + '<div style="font-family:Arial,sans-serif;color:#171d17;line-height:1.6;">'
    + '<p style="margin:0 0 12px;">HPS Anda sudah berstatus <strong>Siap</strong>.</p>'
    + '<div style="padding:14px 16px;border:1px solid #d8ddd2;border-radius:14px;background:#fbfaf6;">'
    + '<div style="font-size:12px;color:#778070;text-transform:uppercase;letter-spacing:.08em;">Nama Pendidikan</div>'
    + '<div style="margin-top:6px;font-size:15px;font-weight:700;">' + htmlEscape_(eventName) + '</div>'
    + '<div style="margin-top:12px;font-size:12px;color:#778070;text-transform:uppercase;letter-spacing:.08em;">Nama HPS</div>'
    + '<div style="margin-top:6px;font-size:15px;font-weight:700;">' + htmlEscape_(hpsName) + '</div>'
    + (noPesanan ? '<div style="margin-top:10px;font-size:13px;color:#55604f;">No. Surat Pesanan: ' + htmlEscape_(noPesanan) + '</div>' : '')
    + (packageId ? '<div style="margin-top:8px;font-size:12px;color:#778070;">ID HPS: ' + htmlEscape_(packageId) + '</div>' : '')
    + '</div>'
    + buildEmailFooterHtml_()
    + '</div>';

  sendEmailNotification_(
    'USER_HPS_READY',
    recipientEmail,
    subject,
    textBody,
    htmlBody,
    {
      eventId: eventId,
      eventName: eventName,
      packageId: packageId,
      hpsName: hpsName,
      triggeredBy: sanitizeText(actorEmail)
    }
  );
}

function htmlEscape_(value) {
  return sanitizeText(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
