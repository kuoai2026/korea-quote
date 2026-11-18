/**
 * 韓國代購報價 · Google 試算表後端
 * 貼到「擴充功能 → Apps Script」，改好 KEY 後部署成「網頁應用程式」。
 */

// ① 改成你們自己的密碼（網站設定裡要填一樣的）
const KEY = '請改成你們的密碼';

const SHEET_NAME = '紀錄';
const HEAD = ['id', '日期', '分類', '品名', '成本', '報價', '數量', '填寫人', '建立時間'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.key !== KEY) return out_({ ok: false, error: '密碼錯誤' });
  const sh = sheet_();
  return out_({ ok: true, items: readAll_(sh), sheetUrl: SpreadsheetApp.getActive().getUrl() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: '資料格式錯誤' }); }
  if (body.key !== KEY) return out_({ ok: false, error: '密碼錯誤' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const it = body.item || {};
    if (body.action === 'add') {
      const id = it.id || Utilities.getUuid();
      if (findRow_(sh, id) < 0) sh.appendRow(toRow_(Object.assign({}, it, { id: id })));
      return out_({ ok: true, id: id });
    }
    const row = findRow_(sh, String(it.id || ''));
    if (row < 0) return out_({ ok: false, error: '找不到這筆資料，可能已被刪除' });
    if (body.action === 'update') {
      sh.getRange(row, 1, 1, HEAD.length).setValues([toRow_(it)]);
      return out_({ ok: true });
    }
    if (body.action === 'delete') {
      sh.deleteRow(row);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: '未知的動作' });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
    sh.getRange('I:I').setNumberFormat('@');
  }
  return sh;
}

function toRow_(it) {
  return [String(it.id), String(it.date || ''), it.cat || '', it.name || '',
    Number(it.cost) || 0, Number(it.quote) || 0, Number(it.qty) || 0,
    it.by || '', String(it.createdAt || new Date().toISOString())];
}

function text_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Taipei', 'M/d');
  return String(v == null ? '' : v);
}

// 讀全部資料；在試算表裡手動新增、沒有 id 的列會自動補上 id
function readAll_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const range = sh.getRange(2, 1, n, HEAD.length);
  const v = range.getValues();
  let fixed = false;
  const items = [];
  v.forEach(function (r) {
    if (!r[3]) return;
    if (!r[0]) { r[0] = Utilities.getUuid(); fixed = true; }
    items.push({
      id: String(r[0]), date: text_(r[1]), cat: String(r[2] || '其他小物'), name: String(r[3]),
      cost: Number(r[4]) || 0, quote: Number(r[5]) || 0, qty: Number(r[6]) || 0,
      by: String(r[7] || ''), createdAt: text_(r[8])
    });
  });
  if (fixed) range.setValues(v);
  return items;
}

function findRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1 || !id) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
