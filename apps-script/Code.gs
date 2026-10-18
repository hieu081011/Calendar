/**
 * Lịch Nhớ — API đọc/ghi Google Sheet
 *
 * Cài đặt:
 *  1. Tạo Google Sheet mới → Extensions → Apps Script → dán toàn bộ file này vào Code.gs.
 *  2. Chọn hàm `setup` → Run (cấp quyền lần đầu). Xem API_KEY trong Execution log.
 *  3. Deploy → New deployment → loại "Web app":
 *       Execute as: Me   ·   Who has access: Anyone
 *     Copy URL kết thúc bằng /exec.
 *  4. Mở app → tab "Ngày quan trọng" → "Kết nối Google Sheet" → dán URL + API key.
 *  Sửa code sau này: Deploy → Manage deployments → Edit → Version: New version (giữ nguyên URL).
 *
 * Sheet "events" (dòng 1 là tiêu đề):
 *  id | name | type (solar|lunar) | day | month | cat | remind | note | updatedAt
 */
const SHEET_NAME = 'events';
const HEADERS = ['id', 'name', 'type', 'day', 'month', 'cat', 'remind', 'note', 'updatedAt'];
const CATS = ['tet', 'gio', 'birthday', 'anniv', 'other'];

/* ---------- Chạy tay 1 lần ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  // Cột chữ để dạng text, tránh Sheets tự đổi "1/2" thành ngày tháng
  ['A', 'B', 'C', 'F', 'H'].forEach(c => sh.getRange(c + '2:' + c).setNumberFormat('@'));
  sh.getRange('I2:I').setNumberFormat('yyyy-mm-dd hh:mm');
  const list = values => SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
  sh.getRange('C2:C').setDataValidation(list(['solar', 'lunar']));
  sh.getRange('F2:F').setDataValidation(list(CATS));

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('API_KEY')) props.setProperty('API_KEY', Utilities.getUuid());
  Logger.log('API_KEY = ' + props.getProperty('API_KEY'));
}

/* ---------- Web app ---------- */
// GET ?key=...  → danh sách (tiện để thử trên trình duyệt)
function doGet(e) {
  return respond_(() => {
    auth_(e.parameter.key);
    return { events: readAll_() };
  });
}

// POST body (text/plain, JSON): { key, action: 'list'|'upsert'|'delete'|'replaceAll', event?, id?, events? }
function doPost(e) {
  return respond_(() => {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    auth_(body.key);
    if (body.action === 'list') return { events: readAll_() };
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      switch (body.action) {
        case 'upsert': return { event: upsert_(body.event) };
        case 'delete': return { deleted: delete_(body.id) };
        case 'replaceAll': return { count: replaceAll_(body.events) };
        default: throw new Error('Action không hợp lệ: ' + body.action);
      }
    } finally {
      lock.releaseLock();
    }
  });
}

/* ---------- Nội bộ ---------- */
function respond_(fn) {
  let out;
  try { out = Object.assign({ ok: true }, fn()); }
  catch (err) { out = { ok: false, error: String((err && err.message) || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function auth_(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!expected || key !== expected) throw new Error('Sai API key');
}

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Chưa có sheet "' + SHEET_NAME + '" — hãy chạy hàm setup()');
  return sh;
}

function clean_(e) {
  if (!e || !String(e.name || '').trim()) throw new Error('Thiếu tên sự kiện');
  const type = String(e.type).trim() === 'solar' ? 'solar' : 'lunar';
  const num = (v, min, max, def) => { const n = parseInt(v, 10); return isNaN(n) ? def : Math.min(max, Math.max(min, n)); };
  return {
    id: String(e.id || '').trim() || 'e' + Date.now() + Math.random().toString(36).slice(2, 6),
    name: String(e.name).trim().slice(0, 200),
    type,
    day: num(e.day, 1, type === 'solar' ? 31 : 30, 1),
    month: num(e.month, 1, 12, 1),
    cat: CATS.indexOf(String(e.cat).trim()) >= 0 ? String(e.cat).trim() : 'other',
    remind: num(e.remind, 0, 365, 0),
    note: String(e.note || '').slice(0, 2000)
  };
}

const toRow_ = e => [e.id, e.name, e.type, e.day, e.month, e.cat, e.remind, e.note, new Date()];

function readAll_() {
  const sh = sheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEADERS.length).getValues()
    .filter(r => String(r[0]).trim() && String(r[1]).trim())
    .map(r => clean_({ id: r[0], name: r[1], type: r[2], day: r[3], month: r[4], cat: r[5], remind: r[6], note: r[7] }));
}

function rowOf_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function upsert_(raw) {
  const sh = sheet_(), e = clean_(raw), row = rowOf_(sh, e.id);
  sh.getRange(row > 0 ? row : sh.getLastRow() + 1, 1, 1, HEADERS.length).setValues([toRow_(e)]);
  return e;
}

function delete_(id) {
  const sh = sheet_(), row = rowOf_(sh, String(id));
  if (row > 0) sh.deleteRow(row);
  return row > 0;
}

function replaceAll_(list) {
  if (!Array.isArray(list)) throw new Error('events phải là mảng');
  const sh = sheet_(), rows = list.map(clean_).map(toRow_), n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, HEADERS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  return rows.length;
}
