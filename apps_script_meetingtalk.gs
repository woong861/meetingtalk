// ================================================================
// 전국대학 미팅단톡 — 승인·문자발송 백엔드 (Apps Script)
// ================================================================
// 【설치 방법】
// 1) 미팅단톡 전용 구글폼의 "응답 시트"를 연다
// 2) 확장 프로그램 → Apps Script → 파일 추가 → 이 내용 전체 붙여넣기
// 3) ⚠️ 붙여넣은 뒤 전체를 다시 복사해서 원본과 같은지 확인할 것
//    (예전에 cmd+V가 무시된 채 배포된 사고가 있었음)
// 4) 프로젝트 설정 → 스크립트 속성에 아래 4개 추가
//      SOLAPI_API_KEY    = 솔라피 API 키
//      SOLAPI_API_SECRET = 솔라피 API 시크릿
//      OPENCHAT_URL_M    = 남자방 오픈채팅 링크
//      OPENCHAT_URL_F    = 여자방 오픈채팅 링크
//      ADMIN_KEY         = 관리 키 (admin.html 에 입력하는 값과 동일)
//    ⚠️ 오픈채팅 링크는 코드에 직접 적지 말 것.
//       (레포가 public이라 링크가 새면 승인제가 무의미해짐)
// 6) 배포 → 새 배포 → 웹 앱 → 실행: 나 / 액세스: 모든 사용자 → 배포
// 7) 나온 /exec URL 을 admin.html 의 API_URL 에 넣는다
//
// ※ '승인상태' '발송시각' '메모' 컬럼은 시트 맨 끝에 자동으로 만들어진다.
//   기존 폼 응답 컬럼은 절대 건드리지 않는다.
// ================================================================

// 응답 시트 ID. 비워두면 "이 스크립트가 붙어 있는 시트"를 쓴다.
// (독립 프로젝트로 만든 경우엔 반드시 채워야 동작함)
const MT_SHEET_ID  = '1S39xHyfjHMwNch28dNotveEQ8NGAai5OTxESGStApkM';
// 관리 키는 스크립트 속성 ADMIN_KEY 에 넣는다 (코드에 적으면 GitHub에 공개됨).
// 속성이 없을 때만 아래 기본값이 쓰인다.
const MT_ADMIN_KEY_FALLBACK = 'meetingtalk-admin';
function mtAdminKey_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || MT_ADMIN_KEY_FALLBACK;
}
const MT_JOIN_CODE = '';                     // 참여코드 (안 쓰면 빈칸)
const MT_SENDER    = '01057182024';          // 솔라피에 등록된 발신번호
const MT_BRAND     = '전국대학 미팅단톡';
const MT_PAY_FORM  = 'https://forms.gle/14sji6FuT4gU9WWS7';  // 입장료 입금 확인 폼 (남학우)
// 단톡방 입장료. 0이면 자동 안내 문자를 보내지 않는다 (2026-09-02 무료 전환)
const MT_ENTRY_FEE = 0;

// 오픈채팅 링크는 스크립트 속성에서 읽는다 (코드에 하드코딩 금지)
function mtOpenChatUrl_(gender) {
  var props = PropertiesService.getScriptProperties();
  var isF = String(gender).indexOf('여') !== -1;
  var isM = String(gender).indexOf('남') !== -1;
  if (!isF && !isM) throw new Error('성별을 알 수 없어 어느 방으로 보낼지 정할 수 없어요: "' + gender + '"');
  var url = props.getProperty(isF ? 'OPENCHAT_URL_F' : 'OPENCHAT_URL_M');
  if (!url) throw new Error('스크립트 속성에 ' + (isF ? 'OPENCHAT_URL_F' : 'OPENCHAT_URL_M') + ' 를 설정해주세요.');
  return { url: url, room: isF ? '여자방' : '남자방' };
}

const MT_COL_STATUS = '승인상태';
const MT_COL_SENT   = '발송시각';
const MT_COL_MEMO   = '메모';
const MT_COL_PAYSENT = '입금안내발송';

// ---------------- 라우터 ----------------
function doGet(e) {
  var out;
  try {
    var p = (e && e.parameter) || {};
    var isAdmin = p.key === mtAdminKey_();
    if (!isAdmin) throw new Error('auth');

    switch (p.action) {
      case 'list':    out = mtList_(); break;
      case 'approve': out = mtDecide_(Number(p.row), true,  p.memo || ''); break;
      case 'reject':  out = mtDecide_(Number(p.row), false, p.memo || ''); break;
      case 'resend':  out = mtDecide_(Number(p.row), true,  p.memo || '', true); break;
      case 'payinfo': out = mtSendPayInfo_(Number(p.row)); break;
      case 'ping':    out = { ok: true, sheet: mtSheet_().getName() }; break;
      default:        out = { error: 'unknown action' };
    }
  } catch (err) {
    out = { error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------- 시트 ----------------
function mtSheet_() {
  var ss = MT_SHEET_ID
    ? SpreadsheetApp.openById(MT_SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('응답 시트를 열지 못했어요. MT_SHEET_ID 를 확인해주세요.');
  return ss.getSheets()[0];
}

// 헤더 탐색: 완전일치 우선 → 부분일치. 오탐 후보는 반드시 제외.
// (연애학개론에서 '추천인 이름 / 인스타 ID' 컬럼이 '이름'·'인스타'로 오탐된 적 있음)
function mtFindCol_(headers, exacts, parts, excludes) {
  excludes = excludes || [];
  var bad = function (h) {
    return excludes.some(function (x) { return h.indexOf(x) !== -1; });
  };
  var i;
  for (i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    if (bad(h)) continue;
    if (exacts.indexOf(h) !== -1) return i;
  }
  for (i = 0; i < headers.length; i++) {
    var h2 = String(headers[i]).trim();
    if (bad(h2)) continue;
    for (var j = 0; j < parts.length; j++) {
      if (h2.indexOf(parts[j]) !== -1) return i;
    }
  }
  return -1;
}

// 관리용 컬럼이 없으면 시트 맨 끝에 새로 만든다 (기존 컬럼 덮어쓰기 금지)
function mtEnsureCol_(sh, headers, name) {
  var idx = headers.indexOf(name);
  if (idx !== -1) return idx;
  var col = headers.length + 1;
  sh.getRange(1, col).setValue(name);
  headers.push(name);
  return col - 1;
}

function mtCols_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim();
  });
  var c = {
    name   : mtFindCol_(headers, ['이름'], ['성함'], ['추천', '인스타', '학교', '학과']),
    gender : mtFindCol_(headers, ['성별'], ['성별'], []),
    school : mtFindCol_(headers, ['재학중인 학교', '학교'], ['학교', '대학'], ['학과']),
    major  : mtFindCol_(headers, ['학과/학년', '학과'], ['학과', '전공', '학번'], []),
    age    : mtFindCol_(headers, ['나이'], ['나이'], []),
    mbti   : mtFindCol_(headers, ['MBTI'], ['MBTI', 'mbti'], []),
    phone  : mtFindCol_(headers, ['연락처', '휴대폰 번호'], ['연락처', '휴대폰', '전화'], []),
    kakao  : mtFindCol_(headers, ['카카오톡', '카카오톡 ID'], ['카카오', '카톡'], ['추천']),
    insta  : mtFindCol_(headers, ['인스타 ID', '인스타그램 아이디'], ['인스타'], ['추천']),
    type   : mtFindCol_(headers, ['소개팅/미팅'], ['소개팅', '미팅', '만남'], []),
    intro  : mtFindCol_(headers, ['한 줄 자기소개'], ['자기소개', '소개'], ['추천', '인스타'])
  };
  c.status = mtEnsureCol_(sh, headers, MT_COL_STATUS);
  c.sent   = mtEnsureCol_(sh, headers, MT_COL_SENT);
  c.memo   = mtEnsureCol_(sh, headers, MT_COL_MEMO);
  c.paysent = mtEnsureCol_(sh, headers, MT_COL_PAYSENT);
  c._headers = headers;
  return c;
}

// 시트가 숫자로 저장해 앞 0이 사라진 번호 복원
function mtFixPhone_(v) {
  var s = String(v == null ? '' : v).replace(/[^0-9]/g, '');
  if (!s) return '';
  if (s.length === 9 || s.length === 10) {
    if (s.charAt(0) !== '0') s = '0' + s;
  }
  return s;
}

function mtPick_(row, idx) {
  return idx >= 0 && idx < row.length ? String(row[idx] == null ? '' : row[idx]).trim() : '';
}

// 시트가 '2026-08-28 13:39' 를 날짜로 자동 변환하는 경우가 있어 다시 포맷해준다
function mtPickTime_(row, idx) {
  if (idx < 0 || idx >= row.length) return '';
  var v = row[idx];
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'MM/dd HH:mm');
  return String(v == null ? '' : v).trim();
}

// ---------------- 목록 ----------------
function mtList_() {
  var sh = mtSheet_();
  var c = mtCols_(sh);
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, items: [], stats: { pending: 0, approvedM: 0, approvedF: 0 } };

  var values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var items = [];
  var stats = { pending: 0, approvedM: 0, approvedF: 0 };

  values.forEach(function (row, i) {
    // 타임스탬프 빈 행 = 운영자가 적어둔 메모 등 → 신청자 아님
    if (!row[0]) return;

    var status = mtPick_(row, c.status) || 'pending';
    var gender = mtPick_(row, c.gender);
    var insta  = mtPick_(row, c.insta).replace(/^@/, '').trim();

    if (status === 'pending') stats.pending++;
    if (status === 'approved') {
      if (gender.indexOf('여') !== -1) stats.approvedF++;
      else if (gender.indexOf('남') !== -1) stats.approvedM++;
    }

    items.push({
      row     : i + 2,
      ts      : row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Seoul', 'MM/dd HH:mm') : String(row[0]),
      tsRaw   : row[0] instanceof Date ? row[0].getTime() : null,
      name    : mtPick_(row, c.name),
      gender  : gender,
      school  : mtPick_(row, c.school),
      major   : mtPick_(row, c.major),
      phone   : mtFixPhone_(mtPick_(row, c.phone)),
      kakao   : mtPick_(row, c.kakao),
      insta   : insta,
      age     : mtPick_(row, c.age),
      mbti    : mtPick_(row, c.mbti),
      type    : mtPick_(row, c.type),
      intro   : mtPick_(row, c.intro),
      status  : status,
      sent    : mtPickTime_(row, c.sent),
      memo    : mtPick_(row, c.memo),
      paysent : mtPickTime_(row, c.paysent)
    });
  });

  items.reverse();  // 최신 신청이 위로
  return { ok: true, items: items, stats: stats };
}

// ---------------- 승인 / 거절 ----------------
function mtDecide_(row, approve, memo, force) {
  if (!row || row < 2) throw new Error('행 번호가 잘못됐어요.');
  var sh = mtSheet_();
  var c = mtCols_(sh);
  var data = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  if (!data[0]) throw new Error('빈 행이에요.');

  var prev = mtPick_(data, c.status);
  if (!approve) {
    sh.getRange(row, c.status + 1).setValue('rejected');
    if (memo) sh.getRange(row, c.memo + 1).setValue(memo);
    return { ok: true, status: 'rejected' };
  }

  // 이미 발송된 건 중복 발송 방지 (resend 로만 강제)
  if (prev === 'approved' && !force) {
    return { ok: true, status: 'approved', skipped: true, message: '이미 발송된 신청이에요.' };
  }

  var target = mtOpenChatUrl_(mtPick_(data, c.gender));

  var phone = mtFixPhone_(mtPick_(data, c.phone));
  if (!/^01[0-9]{8,9}$/.test(phone)) throw new Error('휴대폰 번호가 올바르지 않아요: ' + phone);

  var name = mtPick_(data, c.name) || '신청자';
  var text = mtBuildText_(name, target);
  mtSendSms_(phone, text);

  var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  sh.getRange(row, c.status + 1).setValue('approved');
  sh.getRange(row, c.sent + 1).setValue(now);
  if (memo) sh.getRange(row, c.memo + 1).setValue(memo);

  return { ok: true, status: 'approved', sent: now, to: phone, room: target.room };
}

// 남학우에게 입장료 입금 안내 SMS (90바이트 이내 단문)
function mtSendPayInfo_(row) {
  if (!row || row < 2) throw new Error('행 번호가 잘못됐어요.');
  var sh = mtSheet_();
  var c = mtCols_(sh);
  var data = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  if (!data[0]) throw new Error('빈 행이에요.');

  var gender = mtPick_(data, c.gender);
  if (gender.indexOf('남') === -1) throw new Error('입장료 안내는 남학우에게만 보냅니다.');

  var phone = mtFixPhone_(mtPick_(data, c.phone));
  if (!/^01[0-9]{8,9}$/.test(phone)) throw new Error('휴대폰 번호가 올바르지 않아요: ' + phone);

  var text = '[미팅단톡] 신청 확인! 입장료 5천원 입금 후 폼 제출 ' + MT_PAY_FORM;
  mtSendSms_(phone, text, 'SMS');

  var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  sh.getRange(row, c.paysent + 1).setValue(now);
  return { ok: true, sent: now, to: phone, kind: 'payinfo' };
}

function mtBuildText_(name, target) {
  var lines = [];
  lines.push('[' + MT_BRAND + ']');
  lines.push(name + '님, 신청이 확인됐어요!');
  lines.push('아래 ' + target.room + ' 링크로 입장해주세요.');
  lines.push('');
  lines.push(target.url);
  if (MT_JOIN_CODE) lines.push('참여코드: ' + MT_JOIN_CODE);
  lines.push('');
  lines.push('※ 정치·종교 대화, 부적절한 언행 시 즉시 강퇴됩니다.');
  return lines.join('\n');
}

// ---------------- 솔라피 발송 ----------------
function mtSolapiAuth_() {
  var props  = PropertiesService.getScriptProperties();
  var key    = props.getProperty('SOLAPI_API_KEY');
  var secret = props.getProperty('SOLAPI_API_SECRET');
  if (!key || !secret) {
    throw new Error('스크립트 속성에 SOLAPI_API_KEY / SOLAPI_API_SECRET 을 설정해주세요.');
  }
  var date = new Date().toISOString();
  var salt = Utilities.getUuid().replace(/-/g, '');
  var raw  = Utilities.computeHmacSha256Signature(date + salt, secret);
  var sig  = raw.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
  return 'HMAC-SHA256 apiKey=' + key + ', date=' + date + ', salt=' + salt + ', signature=' + sig;
}

function mtSendSms_(to, text, type) {
  var msg = { to: to, from: MT_SENDER, text: text };
  if (type === 'SMS') {
    msg.type = 'SMS';          // 90바이트 이내 단문 (약 11원)
  } else {
    msg.type = 'LMS';          // 장문 (약 33원)
    msg.subject = MT_BRAND;
  }
  var payload = { message: msg };
  var res = UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send', {
    method            : 'post',
    contentType       : 'application/json',
    headers           : { Authorization: mtSolapiAuth_() },
    payload           : JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('솔라피 발송 실패 (' + code + '): ' + body);
  }
  var json = {};
  try { json = JSON.parse(body); } catch (e) {}
  if (json.statusCode && String(json.statusCode) !== '2000') {
    throw new Error('솔라피 응답 오류: ' + body);
  }
  return json;
}

// ---------------- 신청 즉시 자동 발송 (남학우 입장료 안내) ----------------
// 폼 제출이 시트에 들어오면 실행된다. 남학우면 입장료 안내 SMS를 자동 발송.
// ※ mtInstallTrigger() 를 에디터에서 한 번 실행해야 작동한다.
function mtOnFormSubmit(e) {
  if (MT_ENTRY_FEE <= 0) return;   // 무료 운영 중에는 아무것도 보내지 않는다
  try {
    var sh = mtSheet_();
    var row = sh.getLastRow();
    // 트리거 이벤트에 행 정보가 있으면 그걸 우선 사용
    if (e && e.range && e.range.getRow) row = e.range.getRow();
    if (row < 2) return;

    var c = mtCols_(sh);
    var data = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    if (!data[0]) return;
    if (mtPick_(data, c.gender).indexOf('남') === -1) return;   // 여학우는 발송 안 함
    if (mtPick_(data, c.paysent)) return;                       // 이미 보냈으면 재발송 안 함

    mtSendPayInfo_(row);
  } catch (err) {
    Logger.log('자동 입장료 안내 실패: ' + err);
  }
}

// 트리거 해제 (입장료를 없앨 때 에디터에서 실행)
function mtRemoveTrigger() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'mtOnFormSubmit') { ScriptApp.deleteTrigger(t); n++; }
  });
  var msg = '트리거 ' + n + '개 삭제됨: 자동 입장료 안내가 중지됩니다.';
  Logger.log(msg);
  return msg;
}

// 트리거 설치 (에디터에서 한 번만 실행)
function mtInstallTrigger() {
  var ss = SpreadsheetApp.openById(MT_SHEET_ID);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'mtOnFormSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('mtOnFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  var msg = '설치 완료: 남학우 신청 시 입장료 안내가 자동 발송됩니다.';
  Logger.log(msg);
  return msg;
}

// ---------------- 설치 확인용 ----------------
// 에디터에서 이 함수를 직접 실행하면 컬럼 인식 결과를 로그로 보여준다.
function mtCheckSetup() {
  var sh = mtSheet_();
  var c = mtCols_(sh);
  var headers = c._headers;
  var label = {
    name: '이름', gender: '성별', school: '학교', major: '학과', age: '나이',
    phone: '연락처', kakao: '카톡', insta: '인스타', mbti: 'MBTI',
    type: '희망유형(단톡방 폼엔 없어도 됨)', intro: '자기소개(단톡방 폼엔 없어도 됨)'
  };
  var lines = ['시트: ' + sh.getName(), '헤더: ' + headers.join(' | '), ''];
  Object.keys(label).forEach(function (k) {
    var i = c[k];
    lines.push(label[k] + ' → ' + (i >= 0 ? ('[' + (i + 1) + '] ' + headers[i]) : '❌ 못 찾음'));
  });
  lines.push('');
  var props = PropertiesService.getScriptProperties();
  lines.push('남자방 링크: ' + (props.getProperty('OPENCHAT_URL_M') ? 'OK' : '❌ 없음'));
  lines.push('여자방 링크: ' + (props.getProperty('OPENCHAT_URL_F') ? 'OK' : '❌ 없음'));
  lines.push('솔라피 키: ' + (props.getProperty('SOLAPI_API_KEY') ? 'OK' : '❌ 없음'));
  lines.push('솔라피 시크릿: ' + (props.getProperty('SOLAPI_API_SECRET') ? 'OK' : '❌ 없음'));
  Logger.log(lines.join('\n'));
  return lines.join('\n');
}
