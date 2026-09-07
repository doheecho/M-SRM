/* =====================================================================
 * SCM Risk관리_Code.gs  ·  지정학적 SCM Map Risk 대시보드 백엔드 (G.A.S V8)
 * ---------------------------------------------------------------------
 * 이 파일은 통합_common.gs 공용 헬퍼와 바인딩되어 동작합니다.
 * SCM Map 및 지정학적Risk 마스터 구글 시트 데이터베이스와 결합하여
 * 초고속 지정학적 구매 조달 리스크 센싱 시각화를 선사합니다.
 *
 * ─────────────────────────────────────────────────────────────────────
 * [이번 최적화 패스에서 변경된 부분] (기능·반환값·함수 시그니처는 원본과 100% 동일)
 *
 *  1) BUG FIX: getCurrencyData()
 *     원본은 catch 블록에서 try 블록 안에서 const로 선언된 sheet / usdToday /
 *     eurToday / jpyToday 를 다시 참조했습니다. try{}는 블록 스코프이므로
 *     catch{}에서는 해당 변수들이 애초에 존재하지 않아 "sheet is not defined"류의
 *     ReferenceError가 새로 발생 → 원래 에러 메시지가 완전히 가려지는 죽은 코드였습니다.
 *     → 정상적으로 에러를 그대로 반환하도록 단순화했습니다.
 *
 *  2) 하드코딩된 스프레드시트 ID / 캐시 키 통합
 *     동일한 ID·키 문자열이 여러 함수에 중복 하드코딩되어 있었고,
 *     clearScmRiskCache()의 CACHE_KEY와 getScmRiskData()의 CACHE_KEY가
 *     "같은 문자열을 각자 따로" 들고 있어 향후 한쪽만 수정하면 캐시 불일치가
 *     날 수 있는 구조였습니다. → 파일 상단 상수로 단일화.
 *
 *  3) getScmRiskData() 내부 4개 병렬 Map 통합
 *     riskLevelMap / riskDescMap / riskCountryMap / riskTypeMap 4개를
 *     같은 key(리스크명)로 따로 관리하던 것을 riskMetaMap 1개 객체로 통합
 *     (리스크당 4번의 개별 맵 접근 → 1번의 객체 접근).
 *
 *  4) SCM Raw 시트 컬럼 접근에 매직넘버 대신 이름 상수(SCM_RAW_COL) 부여
 *     row[5], row[6], row[13] 같은 숫자 인덱스만으로는 어떤 컬럼인지
 *     주석 없이는 알 수 없었던 부분을 명시적으로 표기.
 *
 *  5) getCostMarketRiskData()의 3중 try/catch 스프레드시트 탐색 로직을
 *     후보 ID 배열 순회로 단순화 (동일한 열기 시도 패턴이 3번 복붙되어 있었음).
 *
 *  6) include(filename) 중복 선언 제거
 *     같은 GAS 프로젝트에 함께 배포되는 통합_common.gs가 이미 include()를
 *     정의하고 있고(에러 시 빈 문자열 반환하는 더 안전한 버전), common.gs
 *     주석에도 "모듈 Code.gs의 중복 정의는 삭제" 하라고 명시돼 있어 제거했습니다.
 *
 *  변경하지 않은 부분: getPartnerDetailByCode()의 PC_MAP 컬럼 매핑은
 *  시트 실물 구조를 알아야 검증 가능하므로 그대로 유지했습니다.
 *  통합_common.gs / 통합_theme.html / 통합_header.html은 공용 파일이라
 *  이번 작업에서 손대지 않았습니다.
 * ===================================================================== */

// ── 공용 상수 (여기저기 흩어져 있던 하드코딩 값을 한곳으로 통일) ──────────
const SCM_RISK_SS_ID     = "1mrMQ7B09ubu_5agloTKMZV_tjQbN0tHzUezEIMMuVko"; // 지정학적Risk / 원가시황Risk
const SCM_MAP_SS_ID      = "1IcQjU0Ya21SifETdcRFFEB6mBQysXJ5MIflFsFg_hhQ"; // SCM Raw / 원자재SCM
const PARTNER_DB_ID      = "1rQHrhLa3xNZ2DyfK1OS4jeF17xkfXobkV7joMoq4A4w"; // 협력업체 관리대장 / IPC Vendor
const CURRENCY_SS_ID     = "15mDVNS3jFIX4mNdu0OEHMTaHgDbwvNDSmBp7OxHsH9k"; // 회사 환율 관리대장
const SCM_RISK_CACHE_KEY = "SCM_RISK_DATA_CACHE_KEY_V1";
const SCM_RISK_CACHE_TTL_SEC = 21600; // 6시간

// SCM Raw 시트 컬럼 인덱스 (getRange(2,2,...,25) 기준, B열 = index 0)
const SCM_RAW_COL = {
  CODE: 0, SITE: 1, VENDOR: 5, COUNTRY: 6, CITY: 7,
  LAT: 8, LON: 9, ITEM: 10, ITEM_GROUP: 11, RISK: 13,
  TYPE_START: 17, TYPE_END: 24 // S~Z: 생산유형 개별 컬럼
};

function doGet(e) {
  logAccess('SCM Risk 관리');
  const template = HtmlService.createTemplateFromFile('index');
  template.srmTheme = (e && e.parameter && (e.parameter.srm_theme || e.parameter.theme)) || '';

  const bootstrapData = getBootstrap();
  template.bootstrapData = JSON.stringify(bootstrapData);

  return template.evaluate()
      .setTitle('SCM Risk관리 대시보드')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// include(filename)는 통합_common.gs에 이미 정의되어 있어 (같은 GAS 프로젝트에 함께
// 배포되므로 전역에서 그대로 사용 가능) 여기서는 중복 선언하지 않습니다.
// 통합_common.gs 주석에도 "중복 정의는 각 모듈 Code.gs에서 삭제" 하도록 명시되어 있습니다.

/**
 * 유저 캐시(getUserCache) 청크 방식 로딩용 헬퍼 함수
 */
function getScmFromCache_(baseKey) {
  try {
    const cache = CacheService.getUserCache();
    const chunkCount = cache.get(baseKey + "_COUNT");
    if (!chunkCount) return null;
    let fullString = "";
    for (let i = 0; i < parseInt(chunkCount); i++) {
      const chunk = cache.get(baseKey + "_" + i);
      if (!chunk) return null;
      fullString += chunk;
    }
    return JSON.parse(fullString);
  } catch (e) {
    Logger.log("Cache read error for " + baseKey + ": " + e.message);
    return null;
  }
}

/**
 * 유저 캐시(getUserCache) 청크 방식 저장용 헬퍼 함수
 */
function putScmToCache_(baseKey, dataObj) {
  try {
    const cache = CacheService.getUserCache();
    const jsonString = JSON.stringify(dataObj);
    const chunks = jsonString.match(/.{1,90000}/g) || [];
    for (let i = 0; i < chunks.length; i++) {
      cache.put(baseKey + "_" + i, chunks[i], SCM_RISK_CACHE_TTL_SEC);
    }
    cache.put(baseKey + "_COUNT", chunks.length.toString(), SCM_RISK_CACHE_TTL_SEC);
  } catch (e) {
    Logger.log("Cache write error for " + baseKey + ": " + e.message);
  }
}

/**
 * SCM 지정학 캐시를 강제로 비우고 신규 원본을 적재하는 클리어 함수
 */
function clearScmRiskCache() {
  try {
    const cache = CacheService.getUserCache();
    const count = cache.get(SCM_RISK_CACHE_KEY + "_COUNT");
    if (count) {
      for (let i = 0; i < parseInt(count); i++) {
        cache.remove(SCM_RISK_CACHE_KEY + "_" + i);
      }
      cache.remove(SCM_RISK_CACHE_KEY + "_COUNT");
    }
  } catch (e) {
    Logger.log("Cache clear error for SCM: " + e.message);
  }
  return getScmRiskData(true); // force reload fresh data
}

/**
 * 지정학적 SCM Map 및 지정학적Risk 교차 매핑 데이터를 수합하여 반환하는 핵심 API (캐싱 지원)
 */
function getScmRiskData(forceReload) {
  try {
    // 강제 갱신이 아닌 경우 캐시 데이터 존재 유무를 확인해 즉시 초고속 리턴!
    if (!forceReload) {
      const cached = getScmFromCache_(SCM_RISK_CACHE_KEY);
      if (cached) {
        Logger.log("SCM 지정학 대시보드 데이터 캐시 적재 성공!");
        return cached;
      }
    }

    let mapLocations = [];
    let allScmLocations = [];
    let riskStats = [];

    // 1) SCM Risk 관리 시트의 '지정학적Risk' B4:F 읽기
    //    (B: Name, C: Level, D: Description, E: Target Country, F: Risk Type)
    const riskSs = SpreadsheetApp.openById(SCM_RISK_SS_ID);
    const classSheet = riskSs.getSheetByName("지정학적Risk");
    let activeRisks = [];
    let riskMetaMap = {}; // { [name]: { level, desc, country, typeText } }

    if (classSheet) {
      const lastClassRow = classSheet.getLastRow();
      if (lastClassRow >= 4) {
        const classValues = classSheet.getRange(4, 2, lastClassRow - 3, 5).getValues();
        classValues.forEach(r => {
          const name = String(r[0]).trim();
          if (!name) return;
          activeRisks.push(name);
          riskMetaMap[name] = {
            level: String(r[1]).trim() || '하', // 기본은 '하'로 세팅
            desc: String(r[2] || '').trim(),
            country: String(r[3] || '').trim(),
            typeText: String(r[4] || '').trim()
          };
        });
      }
    }

    // 2) SCM Map 시트의 'SCM Raw' 읽기
    const mapSs = SpreadsheetApp.openById(SCM_MAP_SS_ID);
    const scmSheet = mapSs.getSheetByName("SCM Raw");

    if (scmSheet && activeRisks.length > 0) {
      const scmLastRow = scmSheet.getLastRow();
      if (scmLastRow >= 2) {
        // B열(Vendor Code)부터 Z열(생산유형)까지 25개 열을 로드
        const scmRawValues = scmSheet.getRange(2, 2, scmLastRow - 1, 25).getValues();

        // 각 Risk별 통계 집계를 위한 임시 맵
        const statsMap = {};
        activeRisks.forEach(r => {
          statsMap[r] = { vendors: new Set(), cities: new Set(), itemGroups: new Set(), items: [] };
        });

        scmRawValues.forEach(row => {
          const code = String(row[SCM_RAW_COL.CODE] || '').trim();
          const site = String(row[SCM_RAW_COL.SITE] || '').trim();
          const vendor = String(row[SCM_RAW_COL.VENDOR] || '').trim();
          const country = String(row[SCM_RAW_COL.COUNTRY] || '').trim();
          const city = String(row[SCM_RAW_COL.CITY] || '').trim();
          const lat = parseFloat(row[SCM_RAW_COL.LAT]);
          const lon = parseFloat(row[SCM_RAW_COL.LON]);
          const item = String(row[SCM_RAW_COL.ITEM] || '').trim();
          const itemGroup = String(row[SCM_RAW_COL.ITEM_GROUP] || '').trim();

          // S~Z열 개별 생산유형 셀 값들을 안전 파싱하여 결합 (중복 소거)
          const typeVals = [];
          for (let c = SCM_RAW_COL.TYPE_START; c <= SCM_RAW_COL.TYPE_END; c++) {
            const val = String(row[c] || '').trim();
            if (val && val !== '-' && val !== '미분류') typeVals.push(val);
          }
          const typeCombined = typeVals.join(", ");
          const riskStr = String(row[SCM_RAW_COL.RISK] || '').trim();

          // 지진 시뮬레이션용: Risk 분류 여부와 관계없이 실질 좌표가 있는 모든 SCM Location 수합
          if (!isNaN(lat) && !isNaN(lon)) {
            allScmLocations.push({
              code: code, site: site, vendor: vendor, country: country, city: city,
              lat: lat, lon: lon, item: item, itemGroup: itemGroup,
              type: typeCombined, risk: riskStr || '-'
            });
          }

          if (!riskStr) return;

          // 이 행의 Risk 테마들 추출 (# 단위 파싱)
          const rowThemes = riskStr.split('#').map(t => t.trim()).filter(Boolean);

          // 지정학적Risk의 액티브 Risk들과 교차 대조
          rowThemes.forEach(t => {
            if (statsMap[t] && !isNaN(lat) && !isNaN(lon)) {
              if (vendor && vendor !== '-' && vendor !== '미분류') statsMap[t].vendors.add(vendor);
              if (city && city !== '-' && city !== '미분류') statsMap[t].cities.add(city);
              if (itemGroup && itemGroup !== '-' && itemGroup !== '미분류') statsMap[t].itemGroups.add(itemGroup);

              const locObj = {
                code: code, site: site, vendor: vendor, country: country, city: city,
                lat: lat, lon: lon, item: item, itemGroup: itemGroup,
                type: typeCombined, risk: t
              };
              statsMap[t].items.push(locObj);
              mapLocations.push(locObj);
            }
          });
        });

        activeRisks.forEach(r => {
          const entry = statsMap[r];
          const meta = riskMetaMap[r] || {};
          if (entry.items.length > 0) {
            riskStats.push({
              name: r,
              level: meta.level || '하',
              desc: meta.desc || '',
              country: meta.country || '',
              typeText: meta.typeText || '',
              vendorCount: entry.vendors.size,
              cityCount: entry.cities.size,
              itemGroupCount: entry.itemGroups.size,
              count: entry.items.length
            });
          }
        });
      }
    }

    const context = SRM_getUserContext_();

    const resultObj = {
      success: true,
      mapLocations: mapLocations,       // 지정학 탭 지도용 교차 마커 리스트 반환
      allScmLocations: allScmLocations, // 지진 시뮬레이션용 전체 마커 리스트 반환
      riskStats: riskStats,             // 지정학 탭 좌측 카테고리 요약표용 통계 반환
      userEmail: context.email,
      userName: context.name,
      userDept: context.text || context.dept
    };

    // 새로 읽어온 데이터 세트를 유저 캐시에 고밀도 이중 적재
    putScmToCache_(SCM_RISK_CACHE_KEY, resultObj);

    return resultObj;
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * 정규 협력사 및 IPC 거래선 마스터 DB에서 세부 인적, 구매, 재무 지표를 실시간 인출하는 연동 API
 */
function getPartnerDetailByCode(vendorCode) {
  try {
    const ss = SpreadsheetApp.openById(PARTNER_DB_ID);

    function isCodeMatched(codeA, codeB) {
      const strA = String(codeA || '').trim();
      const strB = String(codeB || '').trim();
      if (!strA || !strB) return false;
      if (strA.toLowerCase() === strB.toLowerCase()) return true;
      const numA = parseInt(strA, 10), numB = parseInt(strB, 10);
      return (!isNaN(numA) && !isNaN(numB) && numA === numB);
    }

    let sheet = ss.getSheetByName("협력업체 관리대장");
    let matchedRowIdx = -1, lastCol = 0, isIpc = false;

    if (sheet) {
      const lastRow = sheet.getLastRow();
      lastCol = sheet.getLastColumn();
      if (lastRow >= 5) {
        const codes = sheet.getRange(5, 6, lastRow - 4, 1).getDisplayValues();
        for (let i = 0; i < codes.length; i++) {
          if (isCodeMatched(codes[i][0], vendorCode)) { matchedRowIdx = i + 5; break; }
        }
      }
    }
    if (matchedRowIdx === -1) {
      sheet = ss.getSheetByName("IPC Vendor");
      if (sheet) {
        const lastRow = sheet.getLastRow();
        lastCol = sheet.getLastColumn();
        if (lastRow >= 5) {
          const codes = sheet.getRange(5, 6, lastRow - 4, 1).getDisplayValues();
          for (let i = 0; i < codes.length; i++) {
            if (isCodeMatched(codes[i][0], vendorCode)) { matchedRowIdx = i + 5; isIpc = true; break; }
          }
        }
      }
    }
    if (matchedRowIdx === -1) {
      return { success: false, message: `협력업체 코드 '${vendorCode}'에 일치하는 정규협력사 또는 IPC Vendor를 찾을 수 없습니다.` };
    }

    const displayRow = sheet.getRange(matchedRowIdx, 1, 1, lastCol).getDisplayValues()[0];
    const PC_MAP = [
      3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 98, 26, 27, 28, 29, 30,
      31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48,
      58, 59, 60, 61, 62, 63, 64, 65, 66, 67,
      68, 69, 70, 71, 72, 73, 74, 75, 76, 77,
      78, 79, 80, 81, 82, 83, 84, 85, 86, 87,
      88, 89, 90, 91, 92, 93, 94, 95, 96, 97,
      99, 102, 105, 108,
      111, 112, 113, 114, 115, 116, 117
    ];
    const item = {
      values: PC_MAP.map(c => displayRow[c - 1] || ''),
      isIpc: isIpc,
      lat: displayRow[24 - 1] || '', lon: displayRow[25 - 1] || '',
      factory1Lat: displayRow[100 - 1] || '', factory1Lon: displayRow[101 - 1] || '',
      factory2Lat: displayRow[103 - 1] || '', factory2Lon: displayRow[104 - 1] || '',
      factory3Lat: displayRow[106 - 1] || '', factory3Lon: displayRow[107 - 1] || '',
      factory4Lat: displayRow[109 - 1] || '', factory4Lon: displayRow[110 - 1] || ''
    };
    return { success: true, item: item };
  } catch (e) {
    return { success: false, message: "상세 정보 조회 중 오류가 발생했습니다: " + e.message };
  }
}

/**
 * SCM Map original site detailed produced items list in 원자재SCM sheet
 */
function getDetailData(country, city, vendor) {
  try {
    const rawV = SRM_readSheet_('원자재SCM', { id: SCM_MAP_SS_ID, ttl: 3600, useDisplay: false });
    const values = rawV.map(r => r.map(v => (v == null ? '' : (typeof v === 'string' ? v : String(v)))));
    if (!values || values.length < 2) return { success: true, headers: [], data: [] };

    const headers = values[0].slice(0, 12);
    const data = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[10] === country && row[11] === city && row[9] === vendor) {
        data.push(row.slice(0, 12));
      }
    }
    return { success: true, headers: headers, data: data };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * 깃허브에서 AI 원자재 가격 전망 실시간 데이터를 직접 가져와 바인딩하는 API
 */
function getMarketForecastData() {
  try {
    const url = "https://doheecho.github.io/market-forecast/raw_materials_forecast.json";
    const response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });

    if (response.getResponseCode() !== 200) {
      throw new Error("깃허브 서버 응답 오류 (HTTP " + response.getResponseCode() + ")");
    }

    const jsonText = response.getContentText("UTF-8");
    const forecastData = JSON.parse(jsonText);

    return { success: true, forecastData: forecastData };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * 전용 회사 환율 관리 워크시트에서 USD, EUR, JPY 환율 데이터를 실시간 가져오는 API
 *
 * [BUG FIX] 원본은 catch(e) 안에서 try 블록 내부(const 선언, 블록 스코프)의
 * sheet / usdToday / eurToday / jpyToday를 다시 참조하고 있었습니다.
 * try{}는 별도 블록 스코프이므로 catch{}에서는 해당 변수가 존재하지 않아
 * "sheet is not defined" 형태의 새로운 ReferenceError가 발생하면서
 * 원래 발생한 진짜 에러 메시지를 완전히 삼켜버리는 죽은 코드였습니다.
 * → 정상적으로 에러 메시지를 그대로 반환하도록 단순화했습니다.
 */
function getCurrencyData() {
  try {
    const ss = SpreadsheetApp.openById(CURRENCY_SS_ID);
    const sheet = ss.getSheets()[0]; // 첫 번째 워크시트 자동 포인팅

    // 1) 오늘 환율 가져오기 (D6:D8)
    const todayRange = sheet.getRange("D6:D8").getValues();
    const usdToday = todayRange[0][0];
    const eurToday = todayRange[1][0];
    const jpyToday = todayRange[2][0];

    // 2) 과거 환율 가져오기 (J6:M8 -> Row 6: USD, Row 7: EUR, Row 8: JPY)
    const pastRange = sheet.getRange("J6:M8").getValues();

    return {
      success: true,
      data: {
        usd: {
          today: usdToday,
          past1M: pastRange[0][0], // J6 (1M)
          past3M: pastRange[0][1], // K6 (3M)
          past6M: pastRange[0][2], // L6 (6M)
          past1Y: pastRange[0][3]  // M6 (1Y)
        },
        eur: {
          today: eurToday,
          past1M: pastRange[1][0],
          past3M: pastRange[1][1],
          past6M: pastRange[1][2],
          past1Y: pastRange[1][3]
        },
        jpy: {
          today: jpyToday,
          past1M: pastRange[2][0],
          past3M: pastRange[2][1],
          past6M: pastRange[2][2],
          past1Y: pastRange[2][3]
        }
      }
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * 원가시황Risk 시트에서 민감도, 단가인상 요청내역, 리스크 협력사 데이터를 수합하는 API
 * (Date 객체 등 비직렬화 요소를 문자열로 강제 포맷 변환하여 구글 Apps Script 전송 크래시 완벽 박멸)
 *
 * [최적화] 원본은 "지정학Risk 시트 열기 → 실패 시 SCM맵 시트 열기 → 실패 시 활성 시트 열기"를
 * 세 번의 개별 try/catch 블록으로 복붙해 두었습니다. 동일한 시도를 후보 ID 배열 순회로
 * 단순화했습니다. 탐색 순서·최종 동작은 원본과 동일합니다.
 */
function getCostMarketRiskData() {
  try {
    let sheet = null;

    // 1) 원가시황Risk 시트를 보유한 스프레드시트를 순서대로 탐색
    const candidateIds = [SCM_RISK_SS_ID, SCM_MAP_SS_ID];
    for (const id of candidateIds) {
      try {
        const candidateSheet = SpreadsheetApp.openById(id).getSheetByName("원가시황Risk");
        if (candidateSheet) { sheet = candidateSheet; break; }
      } catch (eCandidate) {
        // 접근 권한 없음 등은 무시하고 다음 후보로 계속 진행
      }
    }

    // 2) 위에서 못 찾았을 경우 최후의 수단으로 현재 활성 스프레드시트 확인
    if (!sheet) {
      try {
        const activeSs = SpreadsheetApp.getActiveSpreadsheet();
        sheet = activeSs ? activeSs.getSheetByName("원가시황Risk") : null;
      } catch (eActive) {
        // no-op
      }
    }

    if (!sheet) {
      throw new Error("'원가시황Risk' 시트를 포함한 스프레드시트 접근 권한이 없거나, 시트명이 일치하지 않습니다.");
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) {
      return { success: true, sensitivity: [], requestList: [], riskPartners: [] };
    }

    // G.A.S JSON 직렬화 장애 방지용 셀 포맷터 (Date 객체 -> String 변환)
    const serializeCell = function(val) {
      if (val instanceof Date) {
        const y = val.getFullYear();
        const m = String(val.getMonth() + 1).padStart(2, '0');
        const d = String(val.getDate()).padStart(2, '0');
        return y + "-" + m + "-" + d;
      }
      return val;
    };

    // 1) 품목군별 지수 민감도 (M3:R)
    const sensRange = sheet.getRange(3, 13, lastRow - 2, 6).getValues();
    const sensitivity = sensRange.filter(r => r.some(cell => String(cell).trim() !== ""))
                                 .map(row => row.map(serializeCell));

    // 2) 단가인상 요청내역 (B3:I)
    const reqRange = sheet.getRange(3, 2, lastRow - 2, 8).getValues();
    const requestList = reqRange.filter(r => r.some(cell => String(cell).trim() !== ""))
                                 .map(row => row.map(serializeCell));

    // 3) 원가·시황측면 Risk 협력사 (AA3:AJ)
    const partnerRange = sheet.getRange(3, 27, lastRow - 2, 10).getValues();
    const riskPartners = partnerRange.filter(r => r.some(cell => String(cell).trim() !== ""))
                                     .map(row => row.map(serializeCell));

    return {
      success: true,
      sensitivity: sensitivity,
      requestList: requestList,
      riskPartners: riskPartners
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
