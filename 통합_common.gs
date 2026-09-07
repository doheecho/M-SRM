/* =====================================================================
 * 통합_common.gs  ·  메-SRM 전 모듈 공용 백엔드 헬퍼 (Google Apps Script V8)
 * ---------------------------------------------------------------------
 * 목적 : 모든 모듈 Code.gs 에 중복 선언돼 있던 사용자/권한/접속로그/메뉴/
 *        include 로직을 단일 소스로 통합하고, 스프레드시트 접근을
 *        "1회 배치 읽기 + CacheService" 로 최소화한다. (개발계획서 §4-1)
 *
 * 배포 : 이 파일은 모든 모듈 GAS 프로젝트에 "동일 파일"로 복사한다.
 *        모듈 Code.gs 에서는 아래 전역 함수를 그대로 호출하면 되고,
 *        중복 정의(getUserDetails / fetchUserDeptAndName / checkUserFullAccess /
 *        logAccess / getMenuLinks / include)는 각 Code.gs 에서 삭제한다.
 * ===================================================================== */

/** 사용자정보 · 바로가기 · 접속로그가 들어있는 공용 Admin 스프레드시트 */
var SRM_ADMIN_SS_ID = '1-qBeuf94mvboL0zZkfVgEAtgKsZiAAD6A6d9XjPlwb8';

/** 세션 이메일 추출 실패 시 최종 폴백 이메일 */
var SRM_FALLBACK_EMAIL = 'dohee.cho@ai.samsunghealthcare.com';

/** 네트워크 단절 / 멀티 구글 로그인 충돌 대비 로컬 사용자 폴백 맵 */
var SRM_USER_FALLBACK = {
  'dohee.cho': { name: '조도희', dept: '구매그룹' },
  'hm2521.kwon': { name: '권혁민', dept: '구매그룹' },
  'ky3247.kim': { name: '김경율', dept: '구매그룹' },
  'nt.kim': { name: '김나단', dept: '구매그룹' },
  'mh501.kim': { name: '김미현', dept: '구매그룹' },
  'kmjun.kim': { name: '김민준', dept: '구매그룹' },
  'yumi91.kim': { name: '김유미', dept: '구매그룹' },
  'yubeom.kim': { name: '김유범', dept: '구매그룹' },
  'es0901.kim': { name: '김은선', dept: '구매그룹' },
  'jung.bae.kim': { name: '김정배', dept: '구매그룹' },
  'jinchul2.kim': { name: '김진철', dept: '구매그룹' },
  'hr1206.kim': { name: '김효령', dept: '구매그룹' },
  'seongoh.noh': { name: '노성오', dept: '구매그룹' },
  'sehyun4.park': { name: '박세현', dept: '구매그룹' },
  'jy3124.park': { name: '박재용', dept: '구매그룹' },
  'jps.baek': { name: '백정필', dept: '구매그룹' },
  'jungjin.seo': { name: '서정진', dept: '구매그룹' },
  'juwon.seo': { name: '서주원', dept: '구매그룹' },
  'mjnew.wang': { name: '왕민정', dept: '구매그룹' },
  'sangduek.lee': { name: '이상득', dept: '구매그룹' },
  'eunho3.lee': { name: '이은호', dept: '구매그룹' },
  'je0408.lee': { name: '이정은', dept: '구매그룹' },
  'sujin.jeong': { name: '정수진', dept: '구매그룹' },
  'jinsang.jung': { name: '정진상', dept: '구매그룹' },
  'hyoseok.cho': { name: '조효석', dept: '구매그룹' },
  'jisoon8.park': { name: '박지순', dept: '구매그룹' },
  'tu.kang': { name: '강태욱', dept: '구매그룹' },
  'syn.joo': { name: '주승연', dept: '구매그룹' }
};

/* 실행(요청) 단위 메모리 캐시 — 동일 doGet 안에서 시트 재조회 방지 */
var SRM__memo = {};

/* ---------------------------------------------------------------------
 * 0. HTML 인클루드
 * ------------------------------------------------------------------- */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return '';
  }
}

/* ---------------------------------------------------------------------
 * 1. 안전 실행 가드 — 특정 시트/셀 손상 시 전체 웹앱이 죽지 않도록
 * ------------------------------------------------------------------- */
function SRM_safe_(fn, fallback) {
  try {
    return fn();
  } catch (e) {
    try { Logger.log('[SRM_safe_] ' + (e && e.stack ? e.stack : e)); } catch (_) {}
    return fallback;
  }
}

/* ---------------------------------------------------------------------
 * 2. 배치 읽기 + CacheService  (개발계획서 §4-1 Batch Read/Cache 극대화)
 *    - openById / getSheetByName / getValues 호출을 요청당 1회로 최소화
 *    - 결과 2D 배열을 CacheService 에 청크 저장(키당 100KB 한계 회피)
 *    - opts: { id, useDisplay(bool, 기본 true), ttl(초, 기본 21600),
 *              cache(bool, 기본 true), cols(정수: 읽을 열 수 상한 — 버퍼열 제외 최적화) }
 * ------------------------------------------------------------------- */
function SRM_readSheet_(sheetName, opts) {
  opts = opts || {};
  var ssId = opts.id || SRM_ADMIN_SS_ID;
  var useDisplay = opts.useDisplay !== false;
  var ttl = opts.ttl || 21600;
  var useCache = opts.cache !== false;
  var colCap = (opts.cols && opts.cols > 0) ? Math.floor(opts.cols) : 0;
  var memoKey = ssId + '||' + sheetName + '||' + (useDisplay ? 'd' : 'v') + (colCap ? ('||c' + colCap) : '');

  if (SRM__memo[memoKey]) return SRM__memo[memoKey];

  var cache = null, cacheKey = 'SRMSHEET::' + memoKey;
  if (useCache) {
    cache = SRM_safe_(function () { return CacheService.getScriptCache(); }, null);
    if (cache) {
      var hit = SRM__cacheGetChunked_(cache, cacheKey);
      if (hit) { SRM__memo[memoKey] = hit; return hit; }
    }
  }

  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss ? ss.getSheetByName(sheetName) : null;
  if (!sheet) throw new Error("'" + sheetName + "' 시트를 찾을 수 없습니다.");

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var nCol = colCap ? Math.min(lastCol, colCap) : lastCol;
  var values = (lastRow < 1 || nCol < 1)
    ? []
    : (useDisplay
        ? sheet.getRange(1, 1, lastRow, nCol).getDisplayValues()
        : sheet.getRange(1, 1, lastRow, nCol).getValues());

  SRM__memo[memoKey] = values;
  if (cache) SRM_safe_(function () { SRM__cachePutChunked_(cache, cacheKey, values, ttl); }, null);
  return values;
}

/** 시트 캐시 무효화 (데이터 수정/삭제 후 호출) */
function SRM_clearSheetCache_(sheetName, id) {
  var ssId = id || SRM_ADMIN_SS_ID;
  ['d', 'v'].forEach(function (mode) {
    var key = 'SRMSHEET::' + ssId + '||' + sheetName + '||' + mode;
    delete SRM__memo[ssId + '||' + sheetName + '||' + mode];
    SRM_safe_(function () {
      var cache = CacheService.getScriptCache();
      var meta = cache.get(key + '::n');
      var n = meta ? parseInt(meta, 10) : 0;
      var ks = [key + '::n'];
      for (var i = 0; i < n; i++) ks.push(key + '::' + i);
      cache.removeAll(ks);
    }, null);
  });
}

function SRM__cachePutChunked_(cache, key, obj, ttl) {
  var json = JSON.stringify(obj);
  if (json.length > 900000) return;           // 너무 크면 캐시 생략(라이브 읽기 유지)
  var SIZE = 90000, n = Math.ceil(json.length / SIZE), batch = {};
  for (var i = 0; i < n; i++) batch[key + '::' + i] = json.substr(i * SIZE, SIZE);
  batch[key + '::n'] = String(n);
  cache.putAll(batch, ttl);
}

function SRM__cacheGetChunked_(cache, key) {
  var meta = cache.get(key + '::n');
  if (!meta) return null;
  var n = parseInt(meta, 10), ks = [];
  for (var i = 0; i < n; i++) ks.push(key + '::' + i);
  var parts = cache.getAll(ks), json = '';
  for (var j = 0; j < n; j++) {
    var p = parts[key + '::' + j];
    if (p == null) return null;               // 청크 유실 → 미스 처리
    json += p;
  }
  return SRM_safe_(function () { return JSON.parse(json); }, null);
}

/* ---------------------------------------------------------------------
 * 3. O(1) 조인용 해시맵 빌더  (개발계획서 §4-1 Hash Map 조회)
 *    rows(2D 배열 또는 객체 배열) → Map(key -> row)
 * ------------------------------------------------------------------- */
function SRM_indexBy_(rows, keyFn) {
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var k = keyFn(rows[i], i);
    if (k === '' || k == null) continue;
    if (!(k in map)) map[k] = rows[i];
  }
  return map;
}

/** 헤더행 1줄 + 데이터행들을 {헤더명: 값} 객체 배열로 변환 */
function SRM_toObjects_(values, headerRowIndex, dataStartRowIndex) {
  headerRowIndex = headerRowIndex || 0;
  dataStartRowIndex = dataStartRowIndex || (headerRowIndex + 1);
  if (!values || values.length <= dataStartRowIndex) return [];
  var headers = values[headerRowIndex].map(function (h) { return String(h).trim(); });
  var out = [];
  for (var r = dataStartRowIndex; r < values.length; r++) {
    var row = values[r], o = {};
    for (var c = 0; c < headers.length; c++) o[headers[c]] = row[c];
    out.push(o);
  }
  return out;
}

/* ---------------------------------------------------------------------
 * 4. 세션 이메일 (활성 → 유효 사용자 → 폴백)
 * ------------------------------------------------------------------- */
function SRM_getActiveUserEmail_() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!email || email.trim() === '' || email === 'Unknown') {
    try { email = Session.getEffectiveUser().getEmail(); } catch (e) {}
  }
  if (!email || email.trim() === '') email = SRM_FALLBACK_EMAIL;
  return email;
}

/* ---------------------------------------------------------------------
 * 5. 사용자정보 조회 (사용자정보 시트 1회 읽기, 캐시 적용)
 *    반환: { email, id, name, dept, hasFullAccess }
 * ------------------------------------------------------------------- */
function SRM_getUserContext_() {
  return SRM_safe_(function () {
    var email = SRM_getActiveUserEmail_();
    var userId = (email && email.indexOf('@') !== -1) ? email.split('@')[0] : '';
    var ctx = { email: email || '', id: userId || 'Unknown', name: '', dept: '', hasFullAccess: false };
    if (!userId) { ctx.name = 'User'; ctx.dept = '삼성메디슨'; return ctx; }

    ctx.name = userId.toUpperCase();
    var rows = SRM_safe_(function () {
      return SRM_readSheet_('사용자정보', { ttl: 600 });
    }, []);

    var lc = userId.toLowerCase().trim();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] && String(rows[i][0]).toLowerCase().trim() === lc) {
        ctx.name = String(rows[i][1] || '').trim() || userId.toUpperCase();
        ctx.dept = String(rows[i][2] || '').trim();
        ctx.hasFullAccess = String(rows[i][3] || '').trim().toUpperCase() === 'Y';
        return ctx;
      }
    }
    // 시트 미등록 → 로컬 폴백 맵
    var fb = SRM_USER_FALLBACK[userId];
    if (fb) { ctx.name = fb.name; ctx.dept = fb.dept; }
    return ctx;
  }, { email: '', id: 'Unknown', name: 'User', dept: '삼성메디슨', hasFullAccess: false });
}

/* ---------------------------------------------------------------------
 * 6. 프론트에서 google.script.run 으로 호출하는 공용 API
 *    (기존 각 모듈 Code.gs 의 동명 함수를 대체)
 * ------------------------------------------------------------------- */

/** M-SRM index 가 사용: 전체 사용자 컨텍스트 객체 */
function getUserDetails() {
  var c = SRM_getUserContext_();
  return { email: c.email, id: c.id, name: c.name, dept: c.dept };
}

/** 통합_header 가 사용: "부서 이름" 표시 문자열 */
function fetchUserDeptAndName() {
  var c = SRM_getUserContext_();
  if (c.dept && c.name) return c.dept + ' ' + c.name;
  if (c.name) return c.name;
  return '구매그룹 ' + (c.id && c.id !== 'Unknown' ? c.id.toUpperCase() : 'USER');
}

/** 보안 데이터 조회 권한(D열 Y) 여부 */
function checkUserFullAccess(email) {
  return SRM_safe_(function () {
    var userId = email
      ? email.split('@')[0]
      : SRM_getActiveUserEmail_().split('@')[0];
    var rows = SRM_readSheet_('사용자정보', { ttl: 600 });
    var lc = String(userId).toLowerCase().trim();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] && String(rows[i][0]).toLowerCase().trim() === lc) {
        return String(rows[i][3] || '').trim().toUpperCase() === 'Y';
      }
    }
    return false;
  }, false);
}

/** 접속로그 1행 append (실패해도 조용히 무시 — 화면 흐름 우선) */
function logAccess(moduleTitle) {
  SRM_safe_(function () {
    var ss = SpreadsheetApp.openById(SRM_ADMIN_SS_ID);
    var sheet = ss.getSheetByName('접속로그') || ss.insertSheet('접속로그');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['접속시간', '접속자ID', '접속모듈']);
      sheet.getRange('A1:C1').setBackground('#D3D3D3').setFontWeight('bold');
    }
    var email = SRM_getActiveUserEmail_();
    sheet.appendRow([new Date(), email ? email.split('@')[0] : 'Unknown', moduleTitle]);
  }, null);
}

/** 바로가기 시트 → 그룹번호(A열)로 묶은 메뉴 그룹 배열 (캐시 적용) */
function getMenuLinks() {
  return SRM_safe_(function () {
    var data = SRM_readSheet_('바로가기', { ttl: 3600 });
    var groups = [], cur = [], curNum = null;
    for (var i = 1; i < data.length; i++) {
      var groupNum = data[i][0];
      var title = data[i][1];
      var url = data[i][2];
      var icon = data[i][3] ? String(data[i][3]).trim() : 'bi-gear-wide-connected';
      var inactive = data[i][4] ? String(data[i][4]).trim().toUpperCase() === 'Y' : false;
      if (!title || !url) continue;
      if (curNum !== groupNum) {
        if (cur.length) groups.push(cur);
        cur = []; curNum = groupNum;
      }
      cur.push({ title: title, url: url, icon: icon, isInactiveFromSheet: inactive });
    }
    if (cur.length) groups.push(cur);
    return groups;
  }, []);
}

/* ---------------------------------------------------------------------
 * 7. 부트스트랩 통합 호출 — 페이지당 3회(사용자/이름/메뉴) 직렬 왕복을 1회로.
 *    사용자정보·바로가기 시트는 SRM_readSheet_ 캐시라서 사실상 0ms.
 *    프론트: google.script.run.getBootstrap() 한 번 → 헤더+포털버튼 즉시 렌더.
 * ------------------------------------------------------------------- */
function getBootstrap() {
  var c = SRM_getUserContext_();
  return {
    user: { email: c.email, id: c.id, name: c.name, dept: c.dept },
    displayName: (c.dept && c.name) ? (c.dept + ' ' + c.name) : (c.name || ('구매그룹 ' + (c.id && c.id !== 'Unknown' ? c.id.toUpperCase() : 'USER'))),
    hasFullAccess: c.hasFullAccess,
    menu: getMenuLinks()
  };
}
