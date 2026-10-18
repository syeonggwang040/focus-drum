// 몰입 드럼 ↔ 구글 시트 연결용 Apps Script
// 1) 구글 시트 → 확장 프로그램 → Apps Script → 이 코드 전체를 붙여 넣기
// 2) 아래 PIN을 원하는 숫자 3자리로 바꾸기
// 3) 배포 → 새 배포 → 유형: 웹 앱, 실행: 나, 액세스: 모든 사용자 → 배포 → 웹 앱 URL 복사
// 4) 몰입 드럼 → 설정 → 구글 시트 연결에 URL과 PIN 입력
// 코드를 고친 뒤에는 배포 → 배포 관리 → 수정 → 새 버전으로 다시 배포해야 반영됩니다.

const PIN = '000';          // 숫자 3자리로 바꾸세요
const SHEET_NAME = '기록';
const MAX_FAILS = 10;       // 1시간에 틀린 PIN이 이만큼 쌓이면 1시간 잠금
const COLS = ['시작 시각', '몰입(분)', '세트', '몰입 횟수', '탭 이탈', 'id', '태도', '충만', '몰입', '돌아봄'];

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ error: 'bad_request' }); }
  const check = checkPin_(body.pin);
  if (check !== 'ok') return out_({ error: check });

  const sheet = sheet_();
  if (body.action === 'add') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const last = sheet.getLastRow();
      const ids = new Set(last > 1 ? sheet.getRange(2, 6, last - 1, 1).getValues().map(r => String(r[0])) : []);
      const rows = (body.sessions || [])
        .filter(s => !ids.has(String(s.id)))
        .map(s => [new Date(s.start), Math.round(s.ms / 6000) / 10, s.set || '', s.rounds ?? '', s.leaves ?? 0, String(s.id),
          s.attitude || '', s.full || '', s.flow || '', s.note || '']);
      if (rows.length) {
        const at = sheet.getLastRow() + 1;
        sheet.getRange(at, 6, rows.length, 5).setNumberFormat('@'); // id·태도·충만·몰입·돌아봄은 글자 그대로 ("4,5"가 숫자 45로 바뀌지 않게)
        sheet.getRange(at, 1, rows.length, COLS.length).setValues(rows);
      }
      return out_({ ok: true, added: rows.length });
    } finally { lock.releaseLock(); }
  }
  if (body.action === 'list') {
    const last = sheet.getLastRow();
    const rows = last > 1 ? sheet.getRange(2, 1, last - 1, COLS.length).getValues() : [];
    return out_({
      sessions: rows.filter(r => r[0] && r[1] !== '').map(r => ({
        start: new Date(r[0]).getTime(), ms: Math.round(Number(r[1]) * 60000), set: String(r[2]),
        rounds: r[3] === '' ? null : Number(r[3]), leaves: Number(r[4]) || 0, id: String(r[5] || new Date(r[0]).getTime()),
        attitude: String(r[6]), full: String(r[7]), flow: String(r[8]), note: String(r[9]),
      })),
    });
  }
  return out_({ error: 'bad_request' });
}

function checkPin_(pin) {
  const props = PropertiesService.getScriptProperties();
  const fails = JSON.parse(props.getProperty('fails') || '[]').filter(t => Date.now() - t < 3600000);
  if (fails.length >= MAX_FAILS) return 'locked';
  if (String(pin) === PIN) return 'ok';
  fails.push(Date.now());
  props.setProperty('fails', JSON.stringify(fails));
  return 'wrong';
}

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.setFrozenRows(1);
  }
  // 머리글이 없거나 예전 6열짜리면 새 열 이름으로 채움 (기존 기록은 그대로)
  sheet.getRange(1, 1, 1, COLS.length).setValues([COLS]);
  return sheet;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
