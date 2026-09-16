/**
 * "나를 발견하기" 응답 수집 백엔드
 * ─────────────────────────────────────────────────────────────
 * Google 스프레드시트에 붙여 쓰는 Apps Script 웹앱입니다.
 * 설치 방법은 같은 폴더의 SETUP.md 를 보세요.
 */

/** 응답이 쌓일 시트 이름 */
var SHEET_NAME = '응답';

/** 응답이 올 때마다 알림 메일을 받을 주소. 비워두면 메일을 보내지 않습니다. */
var NOTIFY_EMAIL = '';

/** 장난 제출을 막고 싶을 때 쓰는 공용 토큰. 비워두면 검사하지 않습니다.
 *  값을 넣으면 index.html 의 ENDPOINT 뒤에 ?token=값 을 붙여야 합니다. */
var SHARED_TOKEN = '';

/** 시트 열 구성: [보낼 때 쓰는 키, 시트 머리글] */
var FIELDS = [
  ['ts',         '제출 시각'],
  ['respondent', '응답자'],
  ['relation',   '관계'],
  ['strength',   '강점'],
  ['energyOn',   '에너지가 켜질 때'],
  ['energyOff',  '에너지가 꺼질 때'],
  ['redButton',  '레드버튼'],
  ['shadow',     '그림자'],
  ['extra',      '한마디'],
  ['ua',         '브라우저']
];

/** 웹앱이 살아 있는지 확인용. 브라우저로 /exec 를 열면 이게 보입니다. */
function doGet(e) {
  return json({ ok: true, service: 'inspire', sheet: SHEET_NAME });
}

/** 설문 페이지가 응답을 보낼 때 호출됩니다. */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var data = {};
    if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); }
      catch (err) { data = (e && e.parameter) || {}; }
    } else {
      data = (e && e.parameter) || {};
    }

    if (SHARED_TOKEN) {
      var token = data.token || (e && e.parameter && e.parameter.token) || '';
      if (token !== SHARED_TOKEN) return json({ ok: false, error: 'unauthorized' });
    }

    var sheet = getSheet_();
    var row = FIELDS.map(function (f) {
      if (f[0] === 'ts') return new Date();
      var v = data[f[0]];
      return (v === null || v === undefined) ? '' : String(v);
    });
    sheet.appendRow(row);
    SpreadsheetApp.flush();

    notify_(data);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (err2) {}
  }
}

/** 메뉴에서 한 번 실행해 두면 머리글과 서식이 만들어집니다. (선택) */
function setup() {
  getSheet_();
  SpreadsheetApp.getActive().toast('시트 준비 완료: ' + SHEET_NAME);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    var header = FIELDS.map(function (f) { return f[1]; });
    sheet.appendRow(header);
    var head = sheet.getRange(1, 1, 1, header.length);
    head.setFontWeight('bold').setBackground('#f1f1f5');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);   // 제출 시각
    sheet.setColumnWidth(2, 110);   // 응답자
    sheet.setColumnWidth(3, 100);   // 관계
    for (var c = 4; c <= 9; c++) sheet.setColumnWidth(c, 320);
    sheet.setColumnWidth(10, 120);  // 브라우저
    sheet.getRange(1, 1, sheet.getMaxRows(), FIELDS.length)
         .setVerticalAlignment('top').setWrap(true);
  }
  return sheet;
}

function notify_(data) {
  if (!NOTIFY_EMAIL) return;
  try {
    var who = (data.respondent || '익명') + (data.relation ? ' · ' + data.relation : '');
    var body = FIELDS.filter(function (f) { return f[0] !== 'ts' && f[0] !== 'ua'; })
      .map(function (f) {
        var v = data[f[0]];
        return v ? ('[' + f[1] + ']\n' + v + '\n') : '';
      })
      .filter(String)
      .join('\n');
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: '[나를 발견하기] ' + who + ' 님의 응답이 도착했습니다',
      body: body + '\n\n— 시트에서 전체 보기: ' + SpreadsheetApp.getActive().getUrl()
    });
  } catch (err) {
    // 알림 실패가 응답 저장을 막지 않게 한다
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
