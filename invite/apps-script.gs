/**
 * 밍짜네 집들이 — 참석/날짜투표 저장용 Google Apps Script
 * 사용법은 invite/README.md 참고. 시트 1행에 머리글이 자동으로 만들어집니다.
 */
/** 호스트 비밀번호: 초대장에서 '호스트'를 누르고 이 값을 넣으면 응답을 삭제할 수 있어요. 꼭 바꾸세요. */
const ADMIN_KEY = '여기에-비밀번호';
const HEAD = ['ts', 'name', 'attend', 'dates', 'count', 'car', 'message', 'transport', 'ride', 'tok'];
/* 호스트가 아직 '날짜 저장'을 한 번도 안 했을 때 쓰는 후보 날짜 (초대장 config.js 와 같게) */
const DEFAULT_DATES = ['2027-03-06', '2027-03-13', '2027-03-20', '2027-03-27'];

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
/* ── 입력 검사: 초대장 화면의 검사는 우회할 수 있으니 여기서 한 번 더 ── */
const NAME_RE = /^[가-힣a-zA-Z0-9 ]{1,10}$/;            // 한글 완성형·영문·숫자·공백, 10자
const CAR_RE = /^\d{2,3}\s?[가-힣]\s?\d{4}$/;          // 12가 3456, 123가4567
const TRANSPORTS = ['transit', 'car', 'ride'];
function candidates_() { const s = settings_(); return (s.dates && s.dates.length) ? s.dates : DEFAULT_DATES; }
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
    const name = String(x.name || '').trim().replace(/\s+/g, ' ');
    if (!NAME_RE.test(name)) return out_({ ok: false, error: 'name-format' });
    const attend = x.attend === 'no' ? 'no' : 'yes';
    const go = attend === 'yes';
    const cand = candidates_();
    // 후보에 있는 날짜만, 중복 없이. 불참이면 날짜 없음
    const dates = go ? cand.filter(d => (x.dates || []).map(String).indexOf(d) >= 0) : [];
    if (go && !dates.length) return out_({ ok: false, error: 'dates' });
    const transport = go ? (TRANSPORTS.indexOf(x.transport) >= 0 ? x.transport : 'transit') : '';
    const car = transport === 'car' ? String(x.car || '').trim() : '';
    if (transport === 'car' && !CAR_RE.test(car)) return out_({ ok: false, error: 'car' });
    const ride = transport === 'ride' ? String(x.ride || '').trim().slice(0, 10) : '';
    const message = String(x.message || '').trim().slice(0, 60);

    const sh = sheet_();
    const all = sh.getDataRange().getValues();
    const i = all.findIndex((r, k) => k > 0 && str_(r[1]).replace(/\s+/g, '') === name.replace(/\s+/g, ''));
    if (i < 0 && all.length - 1 >= MAX_PEOPLE) return out_({ ok: false, error: 'full' });
    // 이미 있는 이름은 처음 등록한 폰(같은 tok)이나 호스트만 고칠 수 있음
    const oldTok = i > 0 ? str_(all[i][9]) : '';
    const isHost = body.key ? keyOk_(body.key) === '' : false;
    if (oldTok && oldTok !== String(x.token || '') && !isHost) return out_({ ok: false, error: 'owner' });
    const tok = oldTok || Utilities.getUuid();
    const row = [new Date(), txt_(name), attend, txt_(dates.join(',')), go ? 1 : 0,
      txt_(car), txt_(message), txt_(transport), txt_(ride), txt_(tok)];
    if (i > 0) sh.getRange(i + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
    cache_().remove('list');
    return out_(Object.assign(payload_(), { token: tok }));
  } finally { lock.releaseLock(); }
}
