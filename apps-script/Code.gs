/**
 * דיירי תמורה — מעוז אביב / אביב אורבן
 * Apps Script: מקבל הגשת טופס, יוצר תיקייה לדייר ב-Drive ושומר שורה בגיליון.
 */

const SHEET_ID = '1U9VKI_PlGwp0UUXMYUx35r2lLnYI_30fCX47lCWEcDQ';
const SHEET_TAB_NAME = 'דיירי תמורה';
const MAOZ_AVIV_FOLDER_ID = '1WeWpt3iRAWrSEiLs8bS596BSET6-8yNJ';

const FILE_FIELDS = {
  C29_rental_contract: 'חוזה שכירות',
  C35_arnona: 'חשבון ארנונה',
  C36_water: 'חשבון מים',
  C37_electricity: 'חשבון חשמל',
  C38_gas: 'חשבון גז',
  C39_id_copy: 'צילום ת.ז',
  C43_bank_confirmation: 'אישור בנק',
  deed_file: 'נסח טאבו',
};

const COLUMN_ORDER = [
  '__timestamp__',
  'C7_first_name', 'C8_last_name', 'C10_id', 'C11_mobile', 'C12_email',
  'C3_address', 'C4_entrance', 'C5_floor', 'C6_apt_num', 'C13_current_address',
  'has_contact', 'C15_contact_name', 'C16_contact_phone', 'C17_contact_email',
  'C18_senior', 'C19_senior_details', 'C20_age',
  'C21_status', 'C23_tenant_names', 'C24_tenant_phone', 'C25_eviction_notice',
  'C40_bank_name', 'C41_branch', 'C42_account',
  'C2_tat_helka', 'deed_address', 'deed_floor', 'deed_apt_num',
  'C44_signed', 'final_confirm',
  '__folder_link__',
  'C29_rental_contract', 'C35_arnona', 'C36_water', 'C37_electricity',
  'C38_gas', 'C39_id_copy', 'C43_bank_confirmation', 'deed_file',
  '__submission_status__',
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (!data.C10_id) return jsonResponse({ success: false, error: 'ת.ז. חסרה' });
    if (!data.C3_address) return jsonResponse({ success: false, error: 'כתובת חסרה' });
    if (!data.C6_apt_num) return jsonResponse({ success: false, error: 'מספר דירה חסר' });

    const buildingFolder = findFolderByName(
      DriveApp.getFolderById(MAOZ_AVIV_FOLDER_ID),
      data.C3_address.trim()
    );
    if (!buildingFolder) {
      return jsonResponse({ success: false, error: 'תיקיית הכתובת לא נמצאה: ' + data.C3_address });
    }

    const residentName = `${data.C7_first_name || ''} ${data.C8_last_name || ''}`.trim();
    const residentFolderName = `דירה ${String(data.C6_apt_num).trim()} — ${residentName}`;

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB_NAME);
    if (!sheet) return jsonResponse({ success: false, error: 'גיליון לא נמצא: ' + SHEET_TAB_NAME });

    const isResubmission = idAlreadySubmitted(sheet, data.C10_id);

    let residentFolder = findFolderByName(buildingFolder, residentFolderName);
    let uploadTarget;

    if (residentFolder) {
      let completions = findFolderByName(residentFolder, 'השלמות');
      if (!completions) completions = residentFolder.createFolder('השלמות');
      uploadTarget = completions;
    } else {
      residentFolder = buildingFolder.createFolder(residentFolderName);
      uploadTarget = residentFolder;
    }

    const isCompletion = uploadTarget.getName() === 'השלמות';
    const tsPrefix = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd_HHmm');

    const fileLinks = {};
    for (const field of Object.keys(FILE_FIELDS)) {
      const f = data[field];
      if (!f || !f.base64) continue;
      const ext = (f.name && f.name.indexOf('.') >= 0) ? f.name.split('.').pop() : 'bin';
      const baseName = FILE_FIELDS[field];
      const filename = isCompletion
        ? `${tsPrefix}_${baseName}.${ext}`
        : `${baseName}.${ext}`;
      const blob = Utilities.newBlob(Utilities.base64Decode(f.base64), f.mimeType || 'application/octet-stream', filename);
      const driveFile = uploadTarget.createFile(blob);
      fileLinks[field] = driveFile.getUrl();
    }

    const row = COLUMN_ORDER.map(col => {
      if (col === '__timestamp__') return new Date();
      if (col === '__folder_link__') return residentFolder.getUrl();
      if (col === '__submission_status__') {
        return isResubmission ? (isCompletion ? 'השלמה' : 'הגשה ראשונה') : 'הגשה ראשונה';
      }
      if (FILE_FIELDS[col]) return fileLinks[col] || '';
      const v = data[col];
      return (v === undefined || v === null) ? '' : v;
    });

    sheet.appendRow(row);

    return jsonResponse({ success: true, folder: residentFolder.getUrl() });

  } catch (err) {
    return jsonResponse({ success: false, error: String(err && err.message || err) });
  }
}

function doGet() {
  return ContentService
    .createTextOutput('דיירי תמורה — Web App פעיל')
    .setMimeType(ContentService.MimeType.TEXT);
}

function findFolderByName(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

function idAlreadySubmitted(sheet, id) {
  if (!id) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const values = sheet.getRange(2, 4, lastRow - 1, 1).getValues(); // column D = ת.ז.
  const idStr = String(id).trim();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === idStr) return true;
  }
  return false;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * הרצה ידנית פעם אחת — מגדיר RTL, מקפיא שורה ראשונה, מעצב כותרת.
 */
function setupSheet() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB_NAME);
  if (!sheet) throw new Error('גיליון לא נמצא: ' + SHEET_TAB_NAME);
  sheet.setRightToLeft(true);
  sheet.setFrozenRows(1);
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol);
  header.setFontWeight('bold');
  header.setBackground('#1a1a1a');
  header.setFontColor('#ffffff');
  header.setHorizontalAlignment('center');
  sheet.autoResizeColumns(1, lastCol);
}
