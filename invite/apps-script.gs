/**
 * 밍짜네 집들이 — 참석/날짜투표 저장용 Google Apps Script
 * 사용법은 invite/README.md 참고. 시트 1행에 머리글이 자동으로 만들어집니다.
 */
/** 호스트 비밀번호: 초대장에서 '호스트'를 누르고 이 값을 넣으면 응답을 삭제할 수 있어요. 꼭 바꾸세요. */
const ADMIN_KEY = '여기에-비밀번호';
const HEAD = ['ts', 'name', 'attend', 'dates', 'count', 'car', 'message', 'transport', 'ride'];

/* 시트에 넣는 글자는 앞에 ' 를 붙여 '그냥 글자'로 저장한다.
   - '=…', '+…', '-…', '@…' 로 시작하는 값을 시트가 수식으로 실행하지 않게 (수식 주입 방지)
   - '2027-03-13' 같은 값을 시트가 날짜로 바꿔버리지 않게 (NaN 원인)
   ' 는 시트 화면에도, 읽어올 때도 보이지 않는다. */
function txt_(v) { v = String(v == null ? '' : v); return v ? "'" + v : ''; }
/* 예전에 날짜로 바뀌어 저장된 칸도 'yyyy-MM-dd' 글자로 되돌려 읽는다. */
function str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v == null ? '' : v);
}
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

/* ── 도배·과부하 방어 ──
   등록: 1분에 MAX_WRITES_PER_MIN 건이 넘으면 잠깐 거절 (친구들끼리면 충분한 양)
   인원: MAX_PEOPLE 명이 넘는 새 이름은 거절 (가짜 이름 도배 방지)
   비밀번호: 1분에 MAX_BAD_KEYS_PER_MIN 번 넘게 틀리면 잠깐 막음 (4자리 비밀번호 찍기 방지)
   목록: 같은 응답을 LIST_CACHE_SEC 초 동안 저장해 두고 돌려줌 (조회 폭주 시 시트 부담 감소) */
const MAX_WRITES_PER_MIN = 30;
const MAX_PEOPLE = 150;
const MAX_BAD_KEYS_PER_MIN = 20;
const LIST_CACHE_SEC = 5;
function cache_() { return CacheService.getScriptCache(); }
function bump_(name, limit) {
  const k = name + Math.floor(Date.now() / 60000), c = cache_();
  const n = Number(c.get(k) || 0);
  if (n >= limit) return false;
  c.put(k, String(n + 1), 120);
  return true;
}
function closed_() { return PropertiesService.getScriptProperties().getProperty('closed') === '1'; }
function keyOk_(key) {
  if (!bump_('kchk', 1000)) return 'rate';
  const c = cache_(), k = 'kbad' + Math.floor(Date.now() / 60000);
  if (Number(c.get(k) || 0) >= MAX_BAD_KEYS_PER_MIN) return 'rate';
  if (String(key || '') === ADMIN_KEY) return '';
  c.put(k, String(Number(c.get(k) || 0) + 1), 120);
  return 'key';
}

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(HEAD);
  return sh;
}
function rows_() {
  const v = sheet_().getDataRange().getValues().slice(1);
  return v.filter(r => r[1]).map(r => ({
    ts: r[0], name: str_(r[1]), attend: str_(r[2]) || 'yes',
    dates: str_(r[3]).split(',').map(s => s.trim()).filter(d => DATE_RE.test(d)),
    count: Number(r[4]) || 1, car: str_(r[5]), message: str_(r[6]),
    transport: str_(r[7]), ride: str_(r[8]),
  }));
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function settings_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('settings') || '{}'); } catch (e) { return {}; }
}
function payload_() { return { ok: true, rows: rows_(), settings: Object.assign(settings_(), { closed: closed_() }) }; }
function doGet() {
  const hit = cache_().get('list');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  const json = JSON.stringify(payload_());
  if (json.length < 90000) cache_().put('list', json, LIST_CACHE_SEC);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function done_() { cache_().remove('list'); return out_(payload_()); }

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (er) { return out_({ ok: false, error: 'bad' }); }
  const host = ['settings', 'delete', 'close'].indexOf(body.action) >= 0;
  if (host) { const bad = keyOk_(body.key); if (bad) return out_({ ok: false, error: bad }); }
  else {
    if (closed_()) return out_({ ok: false, error: 'closed' });
    if (!bump_('w', MAX_WRITES_PER_MIN)) return out_({ ok: false, error: 'rate' });
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return out_({ ok: false, error: 'busy' });
  try {
    if (body.action === 'close') {
      PropertiesService.getScriptProperties().setProperty('closed', body.closed ? '1' : '');
      return done_();
    }
    if (body.action === 'settings') {
      const s = body.settings || {};
      const dates = (s.dates || []).map(String).filter(d => DATE_RE.test(d)).slice(0, 20);
      PropertiesService.getScriptProperties().setProperty('settings', JSON.stringify({ dates: dates }));
      return done_();
    }
    if (body.action === 'delete') {
      const sh0 = sheet_(), target = String(body.name || '').replace(/\s+/g, '');
      const ns = sh0.getRange(1, 2, sh0.getLastRow(), 1).getValues().map(r => str_(r[0]).replace(/\s+/g, ''));
      const k = ns.indexOf(target);
      if (k > 0) sh0.deleteRow(k + 1);
      return done_();
    }
    const x = body.entry || {};
    const name = String(x.name || '').trim().slice(0, 30);
    if (!name) return out_({ ok: false, error: 'name' });
    const dates = (x.dates || []).map(String).filter(d => DATE_RE.test(d)).slice(0, 20);
    const row = [new Date(), txt_(name), x.attend === 'no' ? 'no' : 'yes',
      txt_(dates.join(',')), Math.min(Number(x.count) || 1, 10),
      txt_(String(x.car || '').slice(0, 15)), txt_(String(x.message || '').slice(0, 200)),
      txt_(String(x.transport || '').slice(0, 10)), txt_(String(x.ride || '').slice(0, 20))];
    const sh = sheet_();
    const names = sh.getRange(1, 2, sh.getLastRow(), 1).getValues().map(r => str_(r[0]).replace(/\s+/g, ''));
    const i = names.indexOf(name.replace(/\s+/g, ''));
    if (i <= 0 && names.length - 1 >= MAX_PEOPLE) return out_({ ok: false, error: 'full' });
    if (i > 0) sh.getRange(i + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
    return done_();
  } finally { lock.releaseLock(); }
}
