/**
 * 5급 승진예정자 연수 서버 (Google Apps Script)
 *  1) '내 정보 확인'  : 휴대폰 번호 -> 본인 연수 정보
 *  2) '음료 선택'     : 본인 확인 후 음료 1건 저장(다시 선택하면 덮어쓰기), 담당자는 관리자 번호로 집계 확인
 *
 * - 이 스크립트는 '명단 시트'에 연결(컨테이너 바인딩)해서 사용한다.
 * - 개인정보(이름·연락처)는 시트에만 있고, 이 코드와 GitHub에는 들어 있지 않다.
 * - 명단 시트(첫 번째 시트) 머리글: 연수번호 | 이름 | 연락처 | 1일차분임 | 2일차AI분임 | 호차
 * - 음료 선택 결과는 '음료선택' 시트가 자동으로 만들어져 쌓인다.
 * - 관리자 번호는 코드에 쓰지 않고 [프로젝트 설정 > 스크립트 속성]의 ADMIN_PIN 에 넣는다.
 *   (선택) 선택 마감: 스크립트 속성 DRINK_CLOSE = 2026-10-22T10:00  (한국시간, 이후에는 변경 불가)
 */

var MAX_FAILS = 60;        // 아래 시간 안에 이만큼 틀리면 잠시 조회 중단(번호 무작위 대입 방지)
var ADMIN_MAX_FAILS = 10;  // 관리자 번호 오입력 허용 횟수
var WINDOW_SEC = 600;      // 10분
var DRINK_SHEET = '음료선택';

// 음료 메뉴(여기만 고치면 안내 페이지 선택지도 함께 바뀐다)
var MENU = [
  { id: 'americano',  name: '아메리카노', temps: ['HOT', 'ICE'], decaf: true },
  { id: 'latte',      name: '카페라떼',   temps: ['HOT', 'ICE'], decaf: true },
  { id: 'vanilla',    name: '바닐라라떼', temps: ['HOT', 'ICE'], decaf: true },
  { id: 'yuja',       name: '유자차',     temps: ['HOT', 'ICE'], decaf: false },
  { id: 'chamomile',  name: '카모마일',   temps: ['HOT', 'ICE'], decaf: false },
  { id: 'grapefruit', name: '자몽에이드', temps: ['ICE'],        decaf: false }
];

function digits_(v) {
  var d = String(v == null ? '' : v).replace(/\D/g, '');
  if (d.length === 10 && d.charAt(0) === '1') d = '0' + d;   // 시트가 앞의 0을 지운 경우 복원
  return d;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return out_({ ok: true, service: 'promo5-me' });   // 상태 확인용(개인정보 없음)
}

// 명단에서 번호가 일치하는 한 사람. 시트 형식이 틀리면 'sheet', 없으면 null
function findPerson_(phone) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var values = sh.getDataRange().getValues();
  var head = values[0].map(function (h) { return String(h).replace(/\s/g, ''); });
  var col = function (name) { return head.indexOf(name); };
  var iNo = col('연수번호'), iName = col('이름'), iTel = col('연락처'),
      iG1 = col('1일차분임'), iAi = col('2일차AI분임'), iCar = col('호차');
  if ([iNo, iName, iTel, iG1, iAi, iCar].indexOf(-1) >= 0) return 'sheet';
  for (var r = 1; r < values.length; r++) {
    if (digits_(values[r][iTel]) === phone) {
      return { no: values[r][iNo], name: values[r][iName], group1: values[r][iG1], ai: values[r][iAi], car: values[r][iCar] };
    }
  }
  return null;
}

function drinkSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DRINK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DRINK_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([['연수번호', '이름', '메뉴', '온도', '디카페인', '선택시각']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function menuByName_(name) {
  for (var i = 0; i < MENU.length; i++) if (MENU[i].name === name) return MENU[i];
  return null;
}

function menuById_(id) {
  for (var i = 0; i < MENU.length; i++) if (MENU[i].id === id) return MENU[i];
  return null;
}

function closed_() {
  var v = PropertiesService.getScriptProperties().getProperty('DRINK_CLOSE');
  if (!v) return false;
  var t = Date.parse(String(v).trim() + ':00+09:00');
  return !isNaN(t) && Date.now() >= t;
}

function getChoice_(no) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DRINK_SHEET);
  if (!sh) return null;
  var values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(no)) {
      var m = menuByName_(values[r][2]);
      return { item: m ? m.id : '', temp: values[r][3], decaf: values[r][4] === 'Y' };
    }
  }
  return null;
}

function setDrink_(p, body) {
  if (closed_()) return { ok: false, reason: 'closed' };
  var m = menuById_(String(body.item || ''));
  var temp = String(body.temp || '');
  if (!m || m.temps.indexOf(temp) < 0) return { ok: false, reason: 'invalid' };
  var decaf = !!body.decaf && m.decaf;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = drinkSheet_();
    var values = sh.getDataRange().getValues();
    var row = -1;
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][0]) === String(p.no)) { row = r + 1; break; }
    }
    if (row < 0) row = sh.getLastRow() + 1;
    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'MM-dd HH:mm');
    sh.getRange(row, 1, 1, 6).setValues([[p.no, p.name, m.name, temp, decaf ? 'Y' : '', now]]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, choice: { item: m.id, temp: temp, decaf: decaf } };
}

// 관리자: 집계 + 선택 목록 + 미응답자
function admin_(body) {
  var pin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!pin) return { ok: false, reason: 'noadmin' };
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('adminFails') || 0);
  if (fails >= ADMIN_MAX_FAILS) return { ok: false, reason: 'locked' };
  if (String(body.pin == null ? '' : body.pin).trim() !== String(pin).trim()) {
    cache.put('adminFails', String(fails + 1), WINDOW_SEC);
    return { ok: false, reason: 'badpin' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roster = ss.getSheets()[0].getDataRange().getValues();
  var head = roster[0].map(function (h) { return String(h).replace(/\s/g, ''); });
  var iNo = head.indexOf('연수번호'), iName = head.indexOf('이름');
  if (iNo < 0 || iName < 0) return { ok: false, reason: 'sheet' };

  var rows = [], byKey = {}, answered = {};
  var dsh = ss.getSheetByName(DRINK_SHEET);
  if (dsh) {
    var dv = dsh.getDataRange().getValues();
    for (var r = 1; r < dv.length; r++) {
      if (dv[r][0] === '' || dv[r][0] == null) continue;
      var rec = { no: dv[r][0], name: dv[r][1], item: dv[r][2], temp: dv[r][3], decaf: dv[r][4] === 'Y', time: String(dv[r][5]) };
      rows.push(rec);
      answered[String(rec.no)] = true;
      var key = rec.item + '|' + rec.temp + '|' + (rec.decaf ? 'Y' : '');
      byKey[key] = (byKey[key] || 0) + 1;
    }
  }
  rows.sort(function (a, b) { return Number(a.no) - Number(b.no); });

  var summary = [];
  MENU.forEach(function (m) {
    m.temps.forEach(function (t) {
      [false, true].forEach(function (d) {
        if (d && !m.decaf) return;
        var n = byKey[m.name + '|' + t + '|' + (d ? 'Y' : '')] || 0;
        if (n) summary.push({ item: m.name, temp: t, decaf: d, count: n });
      });
    });
  });

  var missing = [];
  for (var i = 1; i < roster.length; i++) {
    if (roster[i][iNo] === '' || roster[i][iNo] == null) continue;
    if (!answered[String(roster[i][iNo])]) missing.push({ no: roster[i][iNo], name: roster[i][iName] });
  }
  return { ok: true, total: roster.length - 1, answered: rows.length, summary: summary, rows: rows, missing: missing };
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var action = String(body.action || 'me');
    if (action === 'admin') return out_(admin_(body));

    var phone = digits_(body.phone);
    if (phone.length !== 11 || phone.indexOf('010') !== 0) return out_({ ok: false, reason: 'format' });

    var cache = CacheService.getScriptCache();
    var fails = Number(cache.get('fails') || 0);
    if (fails >= MAX_FAILS) return out_({ ok: false, reason: 'locked' });

    var p = findPerson_(phone);
    if (p === 'sheet') return out_({ ok: false, reason: 'sheet' });
    if (!p) {
      cache.put('fails', String(fails + 1), WINDOW_SEC);
      return out_({ ok: false, reason: 'notfound' });
    }

    if (action === 'me') {
      return out_({ ok: true, no: p.no, name: p.name, group1: p.group1, ai: p.ai, car: p.car });
    }
    if (action === 'drink_get') {
      return out_({ ok: true, no: p.no, name: p.name, menu: MENU, choice: getChoice_(p.no), closed: closed_() });
    }
    if (action === 'drink_set') {
      return out_(setDrink_(p, body));
    }
    return out_({ ok: false, reason: 'action' });
  } catch (err) {
    return out_({ ok: false, reason: 'error' });
  }
}
