/**
 * 7분임 상황판 — 백엔드 (Google Apps Script)
 *
 * 하는 일: 구글 스프레드시트 한 장을 데이터베이스로 써서
 *          현장학습 응답을 읽고(GET) 쓰는(POST) 아주 작은 API입니다.
 *
 * 설치: 스프레드시트 → 확장 프로그램 → Apps Script → 이 내용 전부 붙여넣기 →
 *       배포 → 배포 관리 → 편집(연필) → 버전: 새 버전 → 배포
 *       (액세스 권한이 있는 사용자 = '모든 사용자' 여야 합니다)
 *
 * ※ 휴대전화·계좌번호는 수집하지 않습니다.
 */

var FIELD_SHEET = 'field';
var FIELD_HEADERS = [
  'gb',        // 교번
  'name',      // 성명
  'stay',      // 10.7 숙박 여부
  'hotel',     // 숙소명
  'room',      // 객실 형태
  'mate',      // 동실자
  'receipt',   // 숙박 영수증 PDF 제출
  'home7',     // 10.7 실습 후 바로 귀가
  'dinner',    // 10.7 저녁식사 참가
  'home8',     // 10.8 봉사 후 바로 귀가
  'walk',      // 범계역→복지관 도보 이동
  'walkwhy',   // 도보 곤란 사유
  'note',      // 비고
  'ts'         // 마지막 저장 시각
];

/* ───────── 시트 준비 ───────── */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FIELD_SHEET);
  if (!sh) {
    sh = ss.insertSheet(FIELD_SHEET);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, FIELD_HEADERS.length).setValues([FIELD_HEADERS]);
    sh.setFrozenRows(1);
    // 교번의 앞자리 0이 사라지지 않도록 A열 전체를 텍스트 서식으로
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}

function readRows_() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, FIELD_HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var v = values[i];
    if (!v[0]) continue;
    var o = {};
    for (var j = 0; j < FIELD_HEADERS.length; j++) {
      o[FIELD_HEADERS[j]] = (v[j] === null || v[j] === undefined) ? '' : String(v[j]);
    }
    o.gb = normGb_(o.gb);
    out.push(o);
  }
  return out;
}

// 109 / "109" / 109.0 을 모두 "109"로 맞춥니다.
function normGb_(v) {
  var s = String(v == null ? '' : v).trim();
  if (s.indexOf('.') > -1) s = s.split('.')[0];
  return s;
}

function findRow_(sh, gb) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (normGb_(ids[i][0]) === normGb_(gb)) return i + 2;
  }
  return -1;
}

/* ───────── 응답 ───────── */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ───────── 읽기 ───────── */
function doGet() {
  try {
    return json_({ ok: true, rows: readRows_(), at: new Date().getTime() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ───────── 쓰기 ───────── */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    if (body.action !== 'save') throw new Error('unknown action');

    var p = body.payload || {};
    var gb = normGb_(p.gb).slice(0, 6);
    if (!gb) throw new Error('bad input');

    var row = [];
    for (var i = 0; i < FIELD_HEADERS.length; i++) {
      var k = FIELD_HEADERS[i];
      if (k === 'gb') { row.push(gb); continue; }
      if (k === 'ts') { row.push(new Date()); continue; }
      row.push(String(p[k] == null ? '' : p[k]).slice(0, 120));
    }

    var sh = getSheet_();
    var r = findRow_(sh, gb);
    if (r < 0) r = sh.getLastRow() + 1;          // 새 줄
    sh.getRange(r, 1).setNumberFormat('@');       // 교번은 텍스트로
    sh.getRange(r, 1, 1, FIELD_HEADERS.length).setValues([row]);

    SpreadsheetApp.flush();
    return json_({ ok: true, rows: readRows_(), at: new Date().getTime() });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
