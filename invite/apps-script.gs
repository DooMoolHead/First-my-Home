/**
 * 밍짜네 집들이 — 참석/날짜투표 저장용 Google Apps Script
 * 사용법은 invite/README.md 참고. 시트 1행에 머리글이 자동으로 만들어집니다.
 */
const HEAD = ['ts', 'name', 'attend', 'dates', 'count', 'car', 'message'];

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(HEAD);
  return sh;
}
function rows_() {
  const v = sheet_().getDataRange().getValues().slice(1);
  return v.filter(r => r[1]).map(r => ({
    ts: r[0], name: String(r[1]), attend: String(r[2] || 'yes'),
    dates: String(r[3] || '').split(',').filter(Boolean),
    count: Number(r[4]) || 1, car: String(r[5] || ''), message: String(r[6] || ''),
  }));
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return out_({ ok: true, rows: rows_() }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const x = body.entry || {};
    const name = String(x.name || '').trim().slice(0, 30);
    if (!name) return out_({ ok: false, error: 'name' });
    const row = [new Date(), name, x.attend === 'no' ? 'no' : 'yes',
      (x.dates || []).slice(0, 8).join(','), Math.min(Number(x.count) || 1, 10),
      String(x.car || '').slice(0, 15), String(x.message || '').slice(0, 200)];
    const sh = sheet_();
    const names = sh.getRange(1, 2, sh.getLastRow(), 1).getValues().map(r => String(r[0]).replace(/\s+/g, ''));
    const i = names.indexOf(name.replace(/\s+/g, ''));
    if (i > 0) sh.getRange(i + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
    return out_({ ok: true, rows: rows_() });
  } finally { lock.releaseLock(); }
}
