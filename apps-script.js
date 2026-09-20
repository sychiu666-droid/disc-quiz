/**
 * DISC × Belbin 測驗紀錄 — Google Apps Script
 *
 * 使用方式：
 * 1. 建立 Google Sheets，記下網址中的 ID（/d/xxxxx/edit 中間那段）
 * 2. 在 Sheets 裡點「擴充功能 → Apps Script」
 * 3. 把這整份程式碼貼進去，把下方的 SHEET_ID 換成你的 ID
 * 4. 點「部署 → 新增部署」→ 類型選「網路應用程式」
 *    - 執行身分：我（自己的帳號）
 *    - 誰可以存取：任何人（即使是匿名使用者）
 * 5. 複製部署後的網址，貼回 disc_belbin_quiz.html 的 APPS_SCRIPT_URL 變數
 */

const SHEET_ID   = 'YOUR_GOOGLE_SHEET_ID';
const SHEET_NAME = '測驗紀錄';

// 儲存一筆測驗結果（quiz 頁面 POST 過來）
function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['時間', '代碼', '課程', '班級', '學號', '姓名',
                       'DISC類型', 'DISC動物', 'Belbin角色代碼', 'Belbin角色名稱', '複合身份']);
    }

    sheet.appendRow([
      new Date().toLocaleString('zh-TW'),
      data.code,
      data.course,
      data.className,
      data.studentId,
      data.name,
      data.discType,
      data.discAnimal,
      data.belbinRole,
      data.belbinRoleName,
      data.comboName || ''
    ]);

    return ok('saved');
  } catch (err) {
    return ok('error: ' + err);
  }
}

// 依代碼查詢（lookup.html 的 JSONP 查詢，或直接 GET ?code=XXXXXXXX）
function doGet(e) {
  const code     = (e.parameter.code     || '').trim().toUpperCase();
  const callback = (e.parameter.callback || '').trim();

  let result = { found: false };

  if (code) {
    const sheet  = getSheet();
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][1]).toUpperCase() === code) {
        result = {
          found:         true,
          code:          values[i][1],
          course:        values[i][2],
          className:     values[i][3],
          studentId:     values[i][4],
          name:          values[i][5],
          discType:      values[i][6],
          discAnimal:    values[i][7],
          belbinRole:    values[i][8],
          belbinRoleName: values[i][9]
        };
        break;
      }
    }
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function ok(msg) {
  return ContentService.createTextOutput(msg);
}
