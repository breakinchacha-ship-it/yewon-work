// ============================================================
//  예원 작업 등록 — 구글 시트 저장 스크립트
//  이 파일 전체를 Apps Script 편집기에 붙여넣고, "배포 > 새 배포 > 웹 앱"으로 배포합니다.
//  (실행 계정: 나 / 액세스 권한: 모든 사용자)
// ============================================================

const TZ = 'Asia/Seoul';
const DAY_START = '08:30';
const DAY_END = '18:00';
const BREAKS = [['11:00', '11:15'], ['13:00', '14:00'], ['16:00', '16:15']];

const LOG_HEADERS = ['날짜', '이름', '제품', '공정', '시작', '종료', '분(휴게제외)', '등록시각', '기기'];
const CNT_HEADERS = ['날짜', '제품', '개수', '입력자', '입력시각'];

// ---------- 시트 준비
function sheetOf(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSpreadsheetTimeZone() !== TZ) ss.setSpreadsheetTimeZone(TZ);   // 시트 시간대를 한국으로 고정
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  // 날짜·시작·종료 열은 글자로 저장 (시트가 시각 값으로 바꾸지 않게)
  if (name === '기록') sh.getRangeList(['A:A', 'E:F']).setNumberFormat('@');
  else sh.getRange('A:A').setNumberFormat('@');
  return sh;
}
const logSheet = () => sheetOf('기록', LOG_HEADERS);
const cntSheet = () => sheetOf('수량', CNT_HEADERS);

// ---------- 시간 유틸 (모두 'HH:MM' 문자열)
const toMin = t => { const [h, m] = String(t).split(':').map(Number); return h * 60 + m; };
const toHM = m => Utilities.formatString('%02d:%02d', Math.floor(m / 60), m % 60);
function today() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowHM() { return Utilities.formatDate(new Date(), TZ, 'HH:mm'); }
function nowISO() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }

// 휴게시간을 뺀 작업 분
function workMinutes(start, end) {
  let s = toMin(start), e = toMin(end); if (e <= s) return 0;
  let total = e - s;
  BREAKS.forEach(([a, b]) => { const o = Math.min(e, toMin(b)) - Math.max(s, toMin(a)); if (o > 0) total -= o; });
  return total;
}

// ---------- 읽기
function readLogs(date) {
  const sh = logSheet(); const last = sh.getLastRow(); if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues();
  const out = [];
  vals.forEach((r, i) => {
    if (r[1] === '' || r[4] === '') return;                         // 이름·시작 없는 줄은 건너뜀
    const d = r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') : String(r[0]);
    if (d !== date) return;
    out.push({ row: i + 2, date: d, who: r[1], product: r[2], proc: r[3], start: fmt(r[4]), end: fmt(r[5]), mins: r[6], at: String(r[7]) });
  });
  return out;
}
function readCounts(date) {
  const sh = cntSheet(); const last = sh.getLastRow(); if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, CNT_HEADERS.length).getValues();
  return vals.map((r, i) => ({ row: i + 2, date: r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') : String(r[0]), product: r[1], qty: Number(r[2]) || 0, who: r[3] }))
             .filter(c => c.product !== '' && c.date === date);
}
// 시트가 시간을 Date로 바꿔 저장한 경우 대비
function fmt(v) { if (v === '' || v == null) return ''; if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm'); return String(v); }

// ---------- 열린 기록 닫기 (다음 등록 시각 또는 18:00)
function closeOpen(who, date, atHM) {
  const sh = logSheet();
  readLogs(date).filter(l => l.who === who && !l.end).forEach(l => {
    let end = atHM; if (toMin(end) > toMin(DAY_END)) end = DAY_END; if (toMin(end) < toMin(l.start)) end = l.start;
    sh.getRange(l.row, 6).setValue(end);
    sh.getRange(l.row, 7).setValue(workMinutes(l.start, end));
  });
}
// 지난 날짜의 열린 기록은 전부 18:00으로 마감 (요청이 올 때마다 정리)
function closeStale() {
  const sh = logSheet(); const last = sh.getLastRow(); if (last < 2) return;
  const t = today(); const nowM = toMin(nowHM());
  const vals = sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues();
  vals.forEach((r, i) => {
    if (r[1] === '' || r[4] === '') return;                         // 빈 줄 건너뜀
    if (r[5] !== '' && r[5] != null) return;
    const d = r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd') : String(r[0]);
    if (d < t || (d === t && nowM > toMin(DAY_END))) {
      const s = fmt(r[4]); sh.getRange(i + 2, 6).setValue(DAY_END); sh.getRange(i + 2, 7).setValue(workMinutes(s, DAY_END));
    }
  });
}

// ---------- 응답
function reply(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function payload(who, date, all) {
  const logs = readLogs(date);
  return { ok: true, date, logs: all ? logs : logs.filter(l => l.who === who), counts: readCounts(date) };
}

// ---------- GET: 오늘 기록 읽기  ?who=이름&date=yyyy-mm-dd&all=1
function doGet(e) {
  const p = (e && e.parameter) || {};
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try { closeStale(); return reply(payload(p.who || '', p.date || today(), p.all === '1')); }
  finally { lock.releaseLock(); }
}

// ---------- POST: 등록 / 취소 / 수량
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try {
    closeStale();
    const date = d.date || today();
    if (d.type === 'start') {
      let start = d.start || nowHM();
      if (toMin(start) < toMin(DAY_START)) start = DAY_START;
      closeOpen(d.who, date, start);
      logSheet().appendRow([date, d.who, d.product, d.proc || '', start, '', '', nowISO(), d.device || '']);
    } else if (d.type === 'undo') {
      const mine = readLogs(date).filter(l => l.who === d.who);
      const last = mine[mine.length - 1];
      if (last) {
        logSheet().deleteRow(last.row);
        const prev = mine[mine.length - 2];
        if (prev) { logSheet().getRange(prev.row, 6).setValue(''); logSheet().getRange(prev.row, 7).setValue(''); }
      }
    } else if (d.type === 'edit' || d.type === 'delete') {
      const mine = readLogs(date).filter(l => l.who === d.who);
      let idx = mine.findIndex(l => d.row && l.row === Number(d.row));
      if (idx < 0) idx = mine.findIndex(l => l.product === d.product && l.start === d.oldStart);
      if (idx >= 0) {
        const cur = mine[idx], prev = mine[idx - 1], sh = logSheet();
        if (d.type === 'edit') {
          let ns = d.newStart;
          if (toMin(ns) < toMin(DAY_START)) ns = DAY_START;
          if (prev && toMin(ns) <= toMin(prev.start)) ns = toHM(toMin(prev.start) + 1);
          if (cur.end && toMin(ns) > toMin(cur.end)) ns = cur.end;
          sh.getRange(cur.row, 5).setValue(ns);
          if (cur.end) sh.getRange(cur.row, 7).setValue(workMinutes(ns, cur.end));
          if (prev && prev.end === cur.start) { sh.getRange(prev.row, 6).setValue(ns); sh.getRange(prev.row, 7).setValue(workMinutes(prev.start, ns)); }
        } else {
          if (prev && prev.end === cur.start) {
            if (cur.end) { sh.getRange(prev.row, 6).setValue(cur.end); sh.getRange(prev.row, 7).setValue(workMinutes(prev.start, cur.end)); }
            else { sh.getRange(prev.row, 6).setValue(''); sh.getRange(prev.row, 7).setValue(''); }
          }
          sh.deleteRow(cur.row);
        }
      }
    } else if (d.type === 'count') {
      const ex = readCounts(date).find(c => c.product === d.product);
      if (ex) { cntSheet().getRange(ex.row, 3).setValue(Number(d.qty) || 0); cntSheet().getRange(ex.row, 4).setValue(d.who); cntSheet().getRange(ex.row, 5).setValue(nowISO()); }
      else cntSheet().appendRow([date, d.product, Number(d.qty) || 0, d.who, nowISO()]);
    }
    return reply(payload(d.who, date, !!d.all));
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}
