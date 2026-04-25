const CONFIG = {
  EVENT_SHEET_NAME: 'Pendidikan_Events',
  HPS_SHEET_NAME: 'HPS_Packages',
  ACCESS_SHEET_NAME: 'Access_Control',
  NOTIFICATION_SHEET_NAME: 'Notification_Log',
  EMAIL_LOG_SHEET_NAME: 'Email_Log',

  // Static IDs (recommended). Leave empty to use Script Properties.
  STATIC_SHEET_ID: '1AO0-M0LdIudc6_CKLXTdVL9XfclB3IOqfGqg3w0pbWc',
  STATIC_DRIVE_FOLDER_ID: '1ZwoMKvIhl1rJJH2vNHXPjMRDmBPmPSFW',
  // Used by the simplified admin page login. Change this before deploying.
  ADMIN_ACCESS_CODE: '202020',
  EMAIL_NOTIFICATIONS_ENABLED: true,
  EMAIL_NOTIFY_NEW_EDUCATION: true,
  EMAIL_NOTIFY_HPS_READY: true,
  APP_DISPLAY_NAME: 'SIMOD HPS',

  EVENT_HEADERS: [
    'EventId',
    'EventName',
    'CreatedBy',
    'CreatedAt',
    'UpdatedAt',
    'Status'
  ],

  HPS_HEADERS: [
    'PackageId',
    'EventId',
    'EventName',
    'RupNumber',
    'HpsName',
    'HpsFileId',
    'HpsFileUrl',
    'NoPesanan',
    'LegacyInaprocLink',
    'HpsLinkInaprocFileId',
    'HpsLinkInaprocFileUrl',
    'EFakturFileId',
    'EFakturFileUrl',
    'SuratPesananFileId',
    'SuratPesananFileUrl',
    'BastFileId',
    'BastFileUrl',
    'PackageFolderId',
    'PackageFolderUrl',
    'CreatedBy',
    'CreatedAt',
    'Status',
    'UpdatedAt'
  ],

  ACCESS_HEADERS: [
    'RequestId',
    'Email',
    'DisplayName',
    'Status',
    'RequestedAt',
    'ReviewedAt',
    'ReviewedBy',
    'LastAuthenticatedAt'
  ],

  NOTIFICATION_HEADERS: [
    'NotificationId',
    'Type',
    'Audience',
    'RecipientEmail',
    'PackageId',
    'EventId',
    'EventName',
    'HpsName',
    'ActorEmail',
    'Message',
    'IsRead',
    'CreatedAt',
    'ReadAt'
  ],

  EMAIL_LOG_HEADERS: [
    'EmailLogId',
    'Type',
    'RecipientEmail',
    'Subject',
    'Status',
    'ErrorMessage',
    'EventId',
    'EventName',
    'PackageId',
    'HpsName',
    'TriggeredBy',
    'CreatedAt'
  ],

  FILE_COLUMNS: {
    hps: { idIndex: 5, urlIndex: 6, label: 'HPS', prefix: 'HPS' },
    hpsLinkInaproc: { idIndex: 9, urlIndex: 10, label: 'HPS + LINK INAPROC', prefix: 'HPS + LINK INAPROC' },
    eFaktur: { idIndex: 11, urlIndex: 12, label: 'E-Faktur', prefix: 'E-Faktur' },
    suratPesanan: { idIndex: 13, urlIndex: 14, label: 'Surat Pesanan', prefix: 'Surat Pesanan' },
    bast: { idIndex: 15, urlIndex: 16, label: 'BAST', prefix: 'BAST' }
  },

  USER_UPLOAD_KEYS: ['hps'],
  ADMIN_UPLOAD_KEYS: ['hpsLinkInaproc', 'eFaktur', 'suratPesanan', 'bast'],

  ADMIN_ALLOWED_EMAILS: [
    'bahrobah@gmail.com',
    'monetapli22@gmail.com'
  ],

  // Optional behavior beyond gs_doc_manager base pattern.
  SHARE_WITH_LINK: true,
  SESSION_TTL_SECONDS: 21600
};
