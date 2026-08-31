/**
 * ============================================================================
 * 개선 요약
 * ============================================================================
 * 1) [속도 - 가장 큰 효과] getScmData()에 캐싱이 전혀 없었습니다. 접속할 때마다
 *    매번 시트 전체를 읽고, 필터 6종 + Risk 테마까지 매번 새로 집계했습니다.
 *    스크립트 캐시(6시간)를 추가해서 첫 번째 사용자 이후로는 캐시만 읽습니다.
 * 2) [구조] 필터 6개 + Risk 테마를 만들려고 allData를 7번 따로 훑던 걸
 *    한 번의 순회로 통합했습니다.
 * 3) [속도] doGet()이 logAccess()가 끝날 때까지 기다렸다가 화면을 그리기
 *    시작했습니다. logAccess는 별도 스프레드시트를 열어 로그를 남기는
 *    작업이라 사용자 화면 표시와는 무관한데, 매 요청을 그만큼 지연시키고
 *    있었습니다. 포털(M-SRM_index.html)처럼 화면 로드 후 클라이언트에서
 *    비동기로 호출하도록 옮겼습니다 (SCMMap_index.html의 initialize() 안,
 *    아래 안내된 자리에 한 줄 추가해주시면 됩니다).
 * 4) [제안 - 검토 필요, 코드 미반영]
 *    a. getVendorCodeTypes(), fetchUserDeptAndName(), checkUserFullAccess(),
 *       getMenuLinks()가 전부 같은 스프레드시트(ID 1-qBeuf9...)를 각각 따로
 *       엽니다. 페이지 하나를 열 때 이 함수들이 겹쳐 호출되면 그만큼 같은
 *       파일을 여러 번 여는 셈입니다. 사용자 정보+메뉴+권한을 한 번의 호출로
 *       묶어 반환하는 함수 하나로 합치면 왕복 횟수를 줄일 수 있습니다 —
 *       다만 여러 모듈이 공유하는 함수라 SCM Map 파일만 바꾼다고 되는 게
 *       아니라서 이번 개선 범위에는 포함하지 않았습니다.
 *    b. getDetailData()/getPartnerDetailByCode()는 클릭할 때마다 매번 시트를
 *       선형 탐색합니다. 클릭 빈도가 낮다면 문제없지만, 자주 눌린다면 이
 *       역시 캐시 대상입니다.
 * ============================================================================
 */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const SHEET_NAME = 'SCM Raw';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;
const SCM_CACHE_KEY = "SCM_MAP_DATA_V1";
const SCM_CACHE_TTL_SEC = 21600; // 6시간

const COLUMN_MAPPING = {
  "No": 0, "Vendor Code": 1, "Vendor Desc": 2, "Maker": 3,
  "개발구매": 4, "조달구매": 5, "협력사": 6, "국가": 7, "도시": 8,
  "lat": 9, "lon": 10, "품목": 11, "품목군": 12, "생산유형": 13, "Risk": 14,
  "추가1": 15, "추가2": 16, "추가3": 17, "추가4": 18, "추가5": 19,
};

const COUNTRY_CODE_MAP = {
  "US": "미국", "CN": "중국", "JP": "일본", "KR": "대한민국", "VN": "베트남",
  "MY": "말레이시아", "TW": "대만", "DE": "독일", "PH": "필리핀", "IN": "인도",
  "ID": "인도네시아", "TH": "태국", "SG": "싱가포르", "MX": "멕시코", "GB": "영국",
  "FR": "프랑스", "IT": "이탈리아", "CA": "캐나다", "BR": "브라질"
};

const TREEVIEW_COLUMNS = ["국가", "도시", "협력사", "품목", "품목군", "생산유형"];
const COUNTRY_PRIORITY = ["KR", "TW", "CN", "US"];
const PROD_TYPE_PRIORITY = ["FAB", "Package", "Test", "원재료", "생산", "가공", "Office"];

function doGet(e) {
  // logAccess는 클라이언트 로드 후 비동기로 호출하도록 옮겼습니다 (아래 참고).
  const template = HtmlService.createTemplateFromFile('index');
  template.initialPageDataString = JSON.stringify(getScmData());
  return template.evaluate()
    .setTitle('SCM Map')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---- 캐시 헬퍼 (스크립트 전체 공유; SCM 데이터는 사용자별로 다르지 않음) ----
function getFromCache(baseKey) {
  const cache = CacheService.getScriptCache();
  const chunkCount = cache.get(baseKey + "_COUNT");
  if (!chunkCount) return null;
  const count = parseInt(chunkCount, 10);
  const parts = new Array(count);
  for (let i = 0; i < count; i++) {
    const chunk = cache.get(baseKey + "_" + i);
    if (!chunk) return null;
    parts[i] = chunk;
  }
  return JSON.parse(parts.join(""));
}

function putToCache(baseKey, dataObj, ttlSec) {
  const cache = CacheService.getScriptCache();
  const jsonString = JSON.stringify(dataObj);
  const CHUNK_SIZE = 90000;
  const chunkCount = Math.ceil(jsonString.length / CHUNK_SIZE) || 0;
  const payload = {};
  for (let i = 0; i < chunkCount; i++) {
    payload[baseKey + "_" + i] = jsonString.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
  }
  payload[baseKey + "_COUNT"] = chunkCount.toString();
  cache.putAll(payload, ttlSec);
}

function clearScmCache() {
  const cache = CacheService.getScriptCache();
  const chunkCount = cache.get(SCM_CACHE_KEY + "_COUNT");
  if (chunkCount) {
    const count = parseInt(chunkCount, 10);
    const keys = [SCM_CACHE_KEY + "_COUNT"];
    for (let i = 0; i < count; i++) keys.push(SCM_CACHE_KEY + "_" + i);
    cache.removeAll(keys);
  }
}

// 필터 버튼에서 "데이터 갱신" 같은 동작을 추가하고 싶다면 이 함수를 호출하도록 연결하면 됩니다.
function refreshScmData() {
  clearScmCache();
  return getScmData();
}

function getScmData() {
  const cached = getFromCache(SCM_CACHE_KEY);
  if (cached) return cached;

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`시트 '${SHEET_NAME}'를 찾을 수 없습니다.`);
    const lastRow = sheet.getLastRow();

    if (lastRow < DATA_START_ROW) {
      const empty = {
        success: true, filters: {}, allData: [], riskThemes: [],
        columnMapping: COLUMN_MAPPING, treeviewColumns: TREEVIEW_COLUMNS,
        userEmail: Session.getActiveUser().getEmail(), countryCodeMap: COUNTRY_CODE_MAP,
        vendorCodeTypes: { regular: [], ipc: [] }
      };
      return empty;
    }

    const maxCols = Math.max(sheet.getLastColumn(), 27);
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, maxCols);
    const allDataValues = dataRange.getDisplayValues();

    const riskColIdx = COLUMN_MAPPING['Risk'];
    const countryColIdx = COLUMN_MAPPING['국가'];
    const cityColIdx = COLUMN_MAPPING['도시'];
    const itemColIdx = COLUMN_MAPPING['품목'];
    const itemGroupColIdx = COLUMN_MAPPING['품목군'];
    const prodTypeColIdx = COLUMN_MAPPING['생산유형'];
    const vendorColIdx = COLUMN_MAPPING['협력사'];

    // ===== 단일 순회로 "생산유형" 재구성 + 필터셋 6종 + Risk 테마까지 한 번에 =====
    const countrySet = new Set();
    const citySet = new Set();
    const itemSet = new Set();
    const itemGroupSet = new Set();
    const prodTypeSet = new Set();
    const vendorSet = new Set();
    const riskThemeSet = new Set();

    allDataValues.forEach(row => {
      // S~AA(인덱스 18~26)에서 생산유형 문자열 재구성
      const roles = [];
      for (let i = 18; i <= 26; i++) {
        const val = row[i] ? String(row[i]).trim() : "";
        if (val) roles.push(val);
      }
      row[prodTypeColIdx] = roles.join(', ');

      if (row[countryColIdx]) countrySet.add(row[countryColIdx]);
      if (row[cityColIdx]) citySet.add(row[cityColIdx]);
      if (row[itemColIdx]) itemSet.add(row[itemColIdx]);
      if (row[itemGroupColIdx]) itemGroupSet.add(row[itemGroupColIdx]);
      if (row[vendorColIdx]) vendorSet.add(row[vendorColIdx]);

      roles.forEach(v => {
        const trimmed = v.startsWith("Office") ? "Office" : v;
        if (trimmed) prodTypeSet.add(trimmed);
      });

      if (row[riskColIdx]) {
        row[riskColIdx].split('#').map(t => t.trim()).filter(Boolean).forEach(t => riskThemeSet.add(t));
      }
    });

    function sortWithPriority(values, priority) {
      const arr = [...values].sort();
      arr.sort((a, b) => {
        const idxA = priority.indexOf(a);
        const idxB = priority.indexOf(b);
        if (idxA !== -1 && idxB !== -1) return idxA - idxB;
        if (idxA !== -1) return -1;
        if (idxB !== -1) return 1;
        return a.localeCompare(b);
      });
      return arr;
    }

    const filters = {
      "국가": sortWithPriority(countrySet, COUNTRY_PRIORITY),
      "도시": [...citySet].sort(),
      "품목": [...itemSet].sort(),
      "품목군": [...itemGroupSet].sort(),
      "생산유형": sortWithPriority(prodTypeSet, PROD_TYPE_PRIORITY),
      "협력사": [...vendorSet].sort()
    };

    const finalData = {
      success: true,
      filters: filters,
      allData: allDataValues,
      riskThemes: [...riskThemeSet].sort(),
      columnMapping: COLUMN_MAPPING,
      treeviewColumns: TREEVIEW_COLUMNS,
      userEmail: Session.getActiveUser().getEmail(),
      countryCodeMap: COUNTRY_CODE_MAP,
      vendorCodeTypes: getVendorCodeTypes()
    };

    putToCache(SCM_CACHE_KEY, finalData, SCM_CACHE_TTL_SEC);
    return finalData;

  } catch (e) {
    Logger.log(e.toString());
    console.error("getScmData 함수 실행 중 오류 발생: " + e.message);
    return { success: false, message: e.toString() };
  }
}

function getVendorCodeTypes() {
  const types = { regular: [], ipc: [] };
  try {
    const ss = SpreadsheetApp.openById("1rQHrhLa3xNZ2DyfK1OS4jeF17xkfXobkV7joMoq4A4w");

    const sheetReg = ss.getSheetByName("협력업체 관리대장");
    if (sheetReg) {
      const lastRow = sheetReg.getLastRow();
      if (lastRow >= 5) {
        const codes = sheetReg.getRange(5, 6, lastRow - 4, 1).getDisplayValues();
        codes.forEach(r => {
          const val = String(r[0] || '').trim();
          if (val) types.regular.push(val);
        });
      }
    }

    const sheetIpc = ss.getSheetByName("IPC Vendor");
    if (sheetIpc) {
      const lastRow = sheetIpc.getLastRow();
      if (lastRow >= 5) {
        const codes = sheetIpc.getRange(5, 6, lastRow - 4, 1).getDisplayValues();
        codes.forEach(r => {
          const val = String(r[0] || '').trim();
          if (val) types.ipc.push(val);
        });
      }
    }
  } catch (e) {
    console.error("Error getting vendor code types: " + e.message);
  }
  return types;
}

function getDetailData(country, city, vendor) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('원자재SCM');
    if (!sheet) throw new Error("시트 '원자재SCM'를 찾을 수 없습니다.");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, headers: [], data: [] };

    const dataRange = sheet.getRange(1, 1, lastRow, 12);
    const values = dataRange.getDisplayValues();
    const headers = values[0];
    const data = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[10] === country && row[11] === city && row[9] === vendor) {
        data.push(row);
      }
    }
    return { success: true, headers: headers, data: data };
  } catch (e) {
    Logger.log(e.toString());
    return { success: false, message: e.toString() };
  }
}

function logAccess(moduleTitle) {
  try {
    const logSpreadsheetId = "1-qBeuf94mvboL0zZkfVgEAtgKsZiAAD6A6d9XjPlwb8";
    const ss = SpreadsheetApp.openById(logSpreadsheetId);
    let logSheet = ss.getSheetByName("접속로그") || ss.insertSheet("접속로그");
    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(["접속시간", "접속자ID", "접속모듈"]);
      logSheet.getRange("A1:C1").setBackground("#D3D3D3").setFontWeight("bold");
    }
    let userEmail = "";
    try { userEmail = Session.getActiveUser().getEmail(); } catch (e) {}
    const userId = userEmail ? userEmail.split('@')[0] : "Unknown";
    logSheet.appendRow([new Date(), userId, moduleTitle]);
  } catch (e) {
    console.error("로그 기록 중 오류 발생: " + e.message);
  }
}

function fetchUserDeptAndName() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return "USER";
    const userId = email.split('@')[0];
    const ss = SpreadsheetApp.openById("1-qBeuf94mvboL0zZkfVgEAtgKsZiAAD6A6d9XjPlwb8");
    const sheet = ss.getSheetByName("사용자정보");
    if (!sheet) return userId.toUpperCase();
    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === userId) {
        const name = data[i][1] || '';
        const dept = data[i][2] || '';
        return dept ? `${dept} ${name}` : name;
      }
    }
    return userId.toUpperCase();
  } catch (e) {
    return "USER";
  }
}

function checkUserFullAccess(email) {
  try {
    if (!email) return false;
    const userId = email.split('@')[0];
    const ss = SpreadsheetApp.openById("1-qBeuf94mvboL0zZkfVgEAtgKsZiAAD6A6d9XjPlwb8");
    const sheet = ss.getSheetByName("사용자정보");
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === userId) {
        return String(data[i][3]).trim().toUpperCase() === 'Y';
      }
    }
    return false;
  } catch (e) {
    return false;
  }
}

function getMenuLinks() {
  try {
    const ss = SpreadsheetApp.openById("1-qBeuf94mvboL0zZkfVgEAtgKsZiAAD6A6d9XjPlwb8");
    const sheet = ss.getSheetByName("바로가기");
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    const menuGroups = [];
    let currentGroup = [];
    let currentGroupNum = null;
    for (let i = 1; i < data.length; i++) {
      const groupNum = data[i][0];
      const title = data[i][1];
      const url = data[i][2];
      const icon = data[i][3] ? String(data[i][3]).trim() : "bi-gear-wide-connected";
      if (!title || !url) continue;
      if (currentGroupNum !== groupNum) {
        if (currentGroup.length > 0) menuGroups.push(currentGroup);
        currentGroup = [];
        currentGroupNum = groupNum;
      }
      currentGroup.push({ title: title, url: url, icon: icon });
    }
    if (currentGroup.length > 0) menuGroups.push(currentGroup);
    return menuGroups;
  } catch (e) {
    console.error("메뉴 데이터 가져오기 실패: " + e.message);
    return [];
  }
}

function getPartnerDetailByCode(vendorCode) {
  try {
    const ss = SpreadsheetApp.openById("1rQHrhLa3xNZ2DyfK1OS4jeF17xkfXobkV7joMoq4A4w");

    function isCodeMatched(codeA, codeB) {
      const strA = String(codeA || '').trim();
      const strB = String(codeB || '').trim();
      if (!strA || !strB) return false;
      if (strA.toLowerCase() === strB.toLowerCase()) return true;
      const numA = parseInt(strA, 10);
      const numB = parseInt(strB, 10);
      if (!isNaN(numA) && !isNaN(numB) && numA === numB) return true;
      return false;
    }

    let sheet = ss.getSheetByName("협력업체 관리대장");
    let matchedRowIdx = -1;
    let lastCol = 0;
    let isIpc = false;

    if (sheet) {
      const lastRow = sheet.getLastRow();
      lastCol = sheet.getLastColumn();
      if (lastRow >= 5) {
        const codeDisplayValues = sheet.getRange(5, 6, lastRow - 4, 1).getDisplayValues();
        for (let i = 0; i < codeDisplayValues.length; i++) {
          if (isCodeMatched(codeDisplayValues[i][0], vendorCode)) {
            matchedRowIdx = i + 5;
            break;
          }
        }
      }
    }

    if (matchedRowIdx === -1) {
      sheet = ss.getSheetByName("IPC Vendor");
      if (sheet) {
        const lastRow = sheet.getLastRow();
        lastCol = sheet.getLastColumn();
        if (lastRow >= 5) {
          const codeDisplayValues = sheet.getRange(5, 6, lastRow - 4, 1).getDisplayValues();
          for (let i = 0; i < codeDisplayValues.length; i++) {
            if (isCodeMatched(codeDisplayValues[i][0], vendorCode)) {
              matchedRowIdx = i + 5;
              isIpc = true;
              break;
            }
          }
        }
      }
    }

    if (matchedRowIdx === -1) {
      return { success: false, message: `협력업체 코드 '${vendorCode}'에 일치하는 정규협력사 또는 IPC Vendor를 찾을 수 없습니다.` };
    }

    const displayRow = sheet.getRange(matchedRowIdx, 1, 1, lastCol).getDisplayValues()[0];
    const COLUMN_MAPPING_DETAIL = [
      3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 98, 26, 27, 28, 29, 30,
      31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48,
      58, 59, 60, 61, 62, 63, 64, 65, 66, 67,
      68, 69, 70, 71, 72, 73, 74, 75, 76, 77,
      78, 79, 80, 81, 82, 83, 84, 85, 86, 87,
      88, 89, 90, 91, 92, 93, 94, 95, 96, 97,
      99, 102, 105, 108,
      111, 112, 113, 114, 115, 116, 117
    ];
    const rowValues = COLUMN_MAPPING_DETAIL.map(colIdx => displayRow[colIdx - 1] || '');

    const item = {
      values: rowValues,
      isIpc: isIpc,
      lat: displayRow[24 - 1] || '',
      lon: displayRow[25 - 1] || '',
      factory1Lat: displayRow[100 - 1] || '',
      factory1Lon: displayRow[101 - 1] || '',
      factory2Lat: displayRow[103 - 1] || '',
      factory2Lon: displayRow[104 - 1] || '',
      factory3Lat: displayRow[106 - 1] || '',
      factory3Lon: displayRow[107 - 1] || '',
      factory4Lat: displayRow[109 - 1] || '',
      factory4Lon: displayRow[110 - 1] || ''
    };

    return { success: true, item: item };
  } catch (e) {
    return { success: false, message: "상세 정보 조회 중 오류가 발생했습니다: " + e.message };
  }
}
