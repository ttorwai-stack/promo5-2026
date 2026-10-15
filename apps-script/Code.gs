/**
 * 5급 승진예정자 연수 '내 정보 확인' 서버 (Google Apps Script)
 *
 * - 이 스크립트는 '명단 시트'에 연결(컨테이너 바인딩)해서 사용한다.
 * - 개인정보(이름·연락처)는 시트에만 있고, 이 코드와 GitHub에는 들어 있지 않다.
 * - 입력한 휴대폰 번호가 시트와 정확히 일치할 때만 '그 한 사람'의 정보를 돌려준다.
 * - 시트 첫 줄(머리글): 연수번호 | 이름 | 연락처 | 1일차분임 | 2일차AI분임 | 호차
 * - 강사·멘토 이름은 시트에 넣지 않는다(응답에도 포함되지 않음).
 */

var MAX_FAILS = 60;      // 아래 시간 안에 이만큼 틀리면 잠시 조회 중단(번호 무작위 대입 방지)
var WINDOW_SEC = 600;    // 10분

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

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var phone = digits_(body.phone);
    if (phone.length !== 11 || phone.indexOf('010') !== 0) return out_({ ok: false, reason: 'format' });

    var cache = CacheService.getScriptCache();
    var fails = Number(cache.get('fails') || 0);
    if (fails >= MAX_FAILS) return out_({ ok: false, reason: 'locked' });

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var values = sh.getDataRange().getValues();
    var head = values[0].map(function (h) { return String(h).replace(/\s/g, ''); });
    var col = function (name) { return head.indexOf(name); };
    var iNo = col('연수번호'), iName = col('이름'), iTel = col('연락처'),
        iG1 = col('1일차분임'), iAi = col('2일차AI분임'), iCar = col('호차');
    if ([iNo, iName, iTel, iG1, iAi, iCar].indexOf(-1) >= 0) return out_({ ok: false, reason: 'sheet' });

    for (var r = 1; r < values.length; r++) {
      if (digits_(values[r][iTel]) === phone) {
        return out_({
          ok: true,
          no: values[r][iNo],
          name: values[r][iName],
          group1: values[r][iG1],
          ai: values[r][iAi],
          car: values[r][iCar]
        });
      }
    }
    cache.put('fails', String(fails + 1), WINDOW_SEC);
    return out_({ ok: false, reason: 'notfound' });
  } catch (err) {
    return out_({ ok: false, reason: 'error' });
  }
}
