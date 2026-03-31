const CONFIG = {
  EVENT_SHEET_NAME: 'Pendidikan_Events',
  HPS_SHEET_NAME: 'HPS_Packages',

  // Static IDs (recommended). Leave empty to use Script Properties.
  STATIC_SHEET_ID: '1xxpVUWpqk4f1tHMixml6fGhEqNUjFbKaKVDNp5z8cGo',
  STATIC_DRIVE_FOLDER_ID: '1NbfNtue54-Rxq3tZzFWm-wym4Z2j0Wk3',

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
  ],

  FILE_COLUMNS: {
    hps: { idIndex: 5, urlIndex: 6, label: 'HPS', prefix: 'HPS' },
    eFaktur: { idIndex: 11, urlIndex: 12, label: 'E-Faktur', prefix: 'E-Faktur' }
  },

  ADMIN_ALLOWED_EMAILS: [
    'bahrobah@gmail.com',
    'monetapli22@gmail.com'
  ],

  // Optional behavior beyond gs_doc_manager base pattern.
  SHARE_WITH_LINK: true
};
