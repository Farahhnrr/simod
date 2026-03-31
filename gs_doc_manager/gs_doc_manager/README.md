# Pendidikan Militer HPS Document Manager (Google Free Tier)

This starter uses only free Google services:
- Google Sheets (metadata database)
- Google Drive (file storage)
- Google Apps Script (backend + web app)

You can edit code in VS Code and deploy with `clasp`.

## 1) Prerequisites

- Google account
- Node.js 18+ and npm
- VS Code
- `clasp` installed globally:

```bash
npm install -g @google/clasp
```

## 2) Create Google resources

1. Create a Google Spreadsheet and copy its ID from the URL.
2. Create a Google Drive folder for documents and copy its folder ID from the URL.
3. In the spreadsheet, open `Extensions > Apps Script`.
4. In Apps Script editor, open `Project Settings` and copy the `Script ID`.

## 3) Local setup in VS Code

```bash
npm install
cp .clasp.example.json .clasp.json
```

Edit `.clasp.json`:
- `scriptId`: your Apps Script project Script ID
- Keep `rootDir` as `src`

## 4) Configure IDs (static mode)

Set IDs directly in backend code:
- `src/Code.gs`
  - `STATIC_SHEET_ID`
  - `STATIC_DRIVE_FOLDER_ID`

This mode does not require entering IDs in the web page.

## 5) Push code

```bash
npm run push
```

## 6) Deploy as web app

1. Open Apps Script project.
2. Click `Deploy > New deployment`.
3. Type: `Web app`
4. Execute as: `User accessing the web app`
5. Who has access: `Anyone with Google account` (or your org only)
6. Deploy and open URL.

## 7) How it works

- Step 1: Add Pendidikan (Event) from UI.
- Step 2: Select Pendidikan, then create HPS:
  - `HPS Name` is required
  - `RUP Number` is optional
  - `No. Pesanan` is required
- Step 3: Upload files to selected HPS (can be gradual):
  - HPS
  - E-Faktur
- Files are saved to Drive in nested folders:
  - Root folder / Event / `RUP - HPS Name` (or `HPS Name` if no RUP)
- Metadata is written into:
  - `Pendidikan_Events`
  - `HPS_Packages`
- HPS status:
  - `DRAFT` when `No. Pesanan`, HPS, or E-Faktur is incomplete
  - `READY` when `No. Pesanan` is filled and both files exist

## Notes and limits

- Apps Script request limits apply, so very large files are not ideal.
- For production scale, chunked uploads or Drive Picker can be added later.
- This starter is intentionally simple and free-tier friendly.
