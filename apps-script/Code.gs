/**
 * 韓國代購報價 · Google 試算表後端
 * 貼到「擴充功能 → Apps Script」，改好 KEY 後部署成「網頁應用程式」。
 */

// ① 改成你們自己的密碼（網站設定裡要填一樣的）
const KEY = '請改成你們的密碼';

const SHEET_NAME = '紀錄';
const HEAD = ['id', '日期', '分類', '品名', '成本', '報價', '數量', '填寫人', '建立時間'];  // A–I：網站寫入
// J–N：自動計算（不要手動改）。用 INDEX(…,ROW()) 寫法，刪列、插列都不會錯位
const CALC_HEAD = ['營收', '毛利', '毛利率', '底價', '提醒'];
const R_ = c => 'INDEX(' + c + ':' + c + ',ROW())';
const MIN_ = "VLOOKUP(" + R_('C') + ",'規則'!A:C,2,FALSE)";
const PCT_ = "VLOOKUP(" + R_('C') + ",'規則'!A:C,3,FALSE)";
// O–P：韓元成本、匯率（每筆保留當時的匯率）
const EXTRA_HEAD = ['韓元成本', '匯率'];
const ALL_COLS = HEAD.length + 5 + EXTRA_HEAD.length;  // 16
const CALC = [
  '=' + R_('F') + '*' + R_('G'),
  '=(' + R_('F') + '-' + R_('E') + ')*' + R_('G'),
  '=IFERROR((' + R_('F') + '-' + R_('E') + ')/' + R_('F') + ',0)',
  '=IFERROR(CEILING(MAX(' + R_('E') + '+' + MIN_ + ',IF(' + PCT_ + '>0,' + R_('E') + '/(1-' + PCT_ + '),0)),5),"")',
  '=IF(' + R_('F') + '=0,"",IF(' + R_('L') + "<'規則'!$F$2,\"建議不接\",IF(AND(ISNUMBER(" + R_('M') + '),' + R_('F') + '<' + R_('M') + '),"低於底線","")))'
];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.key !== KEY) return out_({ ok: false, error: '密碼錯誤' });
  const sh = sheet_();
  return out_({ ok: true, items: readAll_(sh), rate: getRate_(), sheetUrl: SpreadsheetApp.getActive().getUrl() });
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
    if (body.action === 'setRate') {
      const r = Number(it.rate);
      if (!(r >= 1)) return out_({ ok: false, error: '匯率格式錯誤' });
      setRate_(r, it.by || '');
      return out_({ ok: true, rate: getRate_() });
    }
    if (body.action === 'add') {
      const id = it.id || Utilities.getUuid();
      if (findRow_(sh, id) < 0) sh.appendRow(toRow_(Object.assign({}, it, { id: id })).concat(CALC).concat(extra_(it)));
      return out_({ ok: true, id: id });
    }
    const row = findRow_(sh, String(it.id || ''));
    if (row < 0) return out_({ ok: false, error: '找不到這筆資料，可能已被刪除' });
    if (body.action === 'update') {
      sh.getRange(row, 1, 1, HEAD.length).setValues([toRow_(it)]);
      sh.getRange(row, HEAD.length + 1, 1, CALC.length).setFormulas([CALC]);
      sh.getRange(row, HEAD.length + CALC.length + 1, 1, EXTRA_HEAD.length).setValues([extra_(it)]);
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
    sh.getRange(1, 1, 1, HEAD.length + CALC_HEAD.length).setValues([HEAD.concat(CALC_HEAD)]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
    sh.getRange('I:I').setNumberFormat('@');
  }
  if (!sh.getRange(1, HEAD.length + CALC.length + 1).getValue()) {
    sh.getRange(1, HEAD.length + CALC.length + 1, 1, EXTRA_HEAD.length).setValues([EXTRA_HEAD]).setFontWeight('bold');
  }
  return sh;
}

function extra_(it) {
  return [Number(it.krw) || '', Number(it.rate) || ''];
}

// 匯率存在「總覽」E4:F7
function rateCell_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('總覽') || ss.insertSheet('總覽');
  if (sh.getRange('E4').getValue() !== '匯率設定') {
    sh.getRange('E4:F7').setValues([['匯率設定', ''], ['1 台幣 = ? 韓元', 41.5], ['更新時間', ''], ['更新者', '']]);
    sh.getRange('E4').setFontWeight('bold');
    sh.getRange('F5').setFontColor('#0000FF');
    sh.getRange('F6').setNumberFormat('@');
  }
  return sh;
}
function getRate_() {
  const v = rateCell_().getRange('F5:F7').getValues();
  return { rate: Number(v[0][0]) || 41.5, at: text_(v[1][0]), by: String(v[2][0] || '') };
}
function setRate_(r, by) {
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'M/d HH:mm');
  rateCell_().getRange('F5:F7').setValues([[r], [now], [by]]);
}

function toRow_(it) {
  return [String(it.id), String(it.date || ''), it.cat || '', it.name || '',
    Number(it.cost) || 0, Number(it.quote) || 0, Number(it.qty) || 0,
    it.by || '', String(it.createdAt || new Date().toISOString())];
}

function text_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'Asia/Taipei', 'M/d');
  return String(v == null ? '' : v);
}

// 讀全部資料；在試算表裡手動新增、沒有 id 的列會自動補上 id
function readAll_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const v = sh.getRange(2, 1, n, ALL_COLS).getValues();
  const items = [];
  v.forEach(function (r, i) {
    if (!r[3]) return;
    if (!r[0]) {  // 在試算表手動新增的列：補 id 和計算欄
      r[0] = Utilities.getUuid();
      sh.getRange(i + 2, 1).setValue(r[0]);
      sh.getRange(i + 2, HEAD.length + 1, 1, CALC.length).setFormulas([CALC]);
    }
    items.push({
      id: String(r[0]), date: text_(r[1]), cat: String(r[2] || '其他小物'), name: String(r[3]),
      cost: Number(r[4]) || 0, quote: Number(r[5]) || 0, qty: Number(r[6]) || 0,
      by: String(r[7] || ''), createdAt: text_(r[8]),
      krw: Number(r[14]) || 0, rate: Number(r[15]) || 0
    });
  });
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
