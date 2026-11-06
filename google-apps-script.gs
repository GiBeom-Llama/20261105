/**
 * 위험물 Quiz 제출 기록 저장용 Apps Script
 * - 이 스크립트는 교수님 구글 스프레드시트에 붙어 있고, 웹앱으로 배포한다.
 * - 학생 퀴즈(GitHub Pages)가 결과를 보내면 '제출기록' 시트에 한 줄씩 쌓는다.
 * - 관리자 화면은 ADMIN_KEY가 있어야 목록 보기, 코드 추가, 삭제를 할 수 있다.
 */
const ADMIN_KEY = '여기에-관리자-키를-넣으세요';
const SHEET_NAME = '제출기록';
const HEADER = ['제출 시각', '학번', '이름', '유형', '점수', '정답 수', '문제 수', '제출 방식', 't', '문제 번호', '답안'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME, 0);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* 셀 값이 수식으로 해석되지 않게 막는다 */
function text_(v, max) {
  const s = String(v == null ? '' : v).slice(0, max || 40);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function clean_(r) {
  if (!r || typeof r !== 'object') return null;
  const t = Number(r.t);
  if (!isFinite(t) || t <= 0) return null;
  if (!Array.isArray(r.order) || !Array.isArray(r.answers)) return null;
  if (!r.order.length || r.order.length > 60 || r.order.length !== r.answers.length) return null;
  const order = r.order.map(function (x) { return String(x).slice(0, 12); });
  if (!order.every(function (x) { return /^(mc|ox|sa)\d{1,3}$/.test(x); })) return null;
  const answers = r.answers.map(function (a) {
    if (a == null) return null;
    return typeof a === 'number' ? a : String(a).slice(0, 60);
  });
  const sid = String(r.sid == null ? '' : r.sid).trim().slice(0, 20);
  const name = String(r.name == null ? '' : r.name).trim().slice(0, 20);
  if (!sid || !name) return null;
  const num = function (v, lo, hi) { v = Number(v); return isFinite(v) ? Math.max(lo, Math.min(hi, Math.round(v))) : ''; };
  return {
    sid: sid, name: name, t: t, mode: String(r.mode || '').slice(0, 20), order: order, answers: answers,
    score: num(r.score, 0, 100), right: num(r.right, 0, 60), total: order.length
  };
}

function row_(r, src) {
  return [
    Utilities.formatDate(new Date(r.t), 'Asia/Seoul', 'yyyy.MM.dd HH:mm'),
    "'" + r.sid, text_(r.name, 20), text_(r.mode, 20), r.score, r.right, r.total,
    src === 'code' ? '제출 코드' : '자동 제출', r.t, JSON.stringify(r.order), JSON.stringify(r.answers)
  ];
}

function list_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, HEADER.length).getValues().map(function (v) {
    let order = [], answers = [];
    try { order = JSON.parse(v[9]); answers = JSON.parse(v[10]); } catch (e) {}
    return {
      sid: String(v[1]), name: String(v[2]), mode: String(v[3]), t: Number(v[8]),
      src: v[7] === '제출 코드' ? 'code' : 'auto', order: order, answers: answers
    };
  }).filter(function (r) { return r.t && r.order.length; });
}

function doGet() {
  return out_({ ok: true, service: '위험물 Quiz 제출 기록' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();

    if (body.action === 'submit') {
      const r = clean_(body.record);
      if (!r) return out_({ ok: false, error: 'bad_record' });
      const dup = list_().some(function (x) { return x.sid === r.sid && x.t === r.t; });
      if (!dup) sh.appendRow(row_(r, 'auto'));
      return out_({ ok: true });
    }

    if (body.key !== ADMIN_KEY) return out_({ ok: false, error: 'key' });

    if (body.action === 'list') return out_({ ok: true, rows: list_() });

    if (body.action === 'import') {
      const seen = {};
      list_().forEach(function (x) { seen[x.sid + '|' + x.t] = true; });
      const rows = [];
      let dup = 0, bad = 0;
      (Array.isArray(body.records) ? body.records : []).slice(0, 200).forEach(function (raw) {
        const r = clean_(raw);
        if (!r) { bad++; return; }
        const k = r.sid + '|' + r.t;
        if (seen[k]) { dup++; return; }
        seen[k] = true;
        rows.push(row_(r, 'code'));
      });
      if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADER.length).setValues(rows);
      return out_({ ok: true, added: rows.length, dup: dup, bad: bad });
    }

    if (body.action === 'delete') {
      const n = sh.getLastRow() - 1;
      if (n > 0) {
        const vals = sh.getRange(2, 1, n, HEADER.length).getValues();
        for (let i = vals.length - 1; i >= 0; i--) {
          if (String(vals[i][1]) === String(body.sid) && Number(vals[i][8]) === Number(body.t)) { sh.deleteRow(i + 2); break; }
        }
      }
      return out_({ ok: true });
    }

    return out_({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}
