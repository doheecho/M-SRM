/* =====================================================================
 * SCM Risk관리_Code.gs  ·  지정학적 SCM Map Risk 대시보드 백엔드 (G.A.S V8)
 * ---------------------------------------------------------------------
 * 이 파일은 통합_common.gs 공용 헬퍼와 바인딩되어 동작합니다.
 * SCM Map 및 지정학적Risk 마스터 구글 시트 데이터베이스와 결합하여
 * 초고속 지정학적 구매 조달 리스크 센싱 시각화를 선사합니다.
 * ===================================================================== */

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

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

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
      cache.put(baseKey + "_" + i, chunks[i], 21600); // 6시간 동안 보존
    }
    cache.put(baseKey + "_COUNT", chunks.length.toString(), 21600);
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
    const CACHE_KEY = "SCM_RISK_DATA_CACHE_KEY_V1";
    const count = cache.get(CACHE_KEY + "_COUNT");
    if (count) {
      for (let i = 0; i < parseInt(count); i++) {
        cache.remove(CACHE_KEY + "_" + i);
      }
      cache.remove(CACHE_KEY + "_COUNT");
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
    const CACHE_KEY = "SCM_RISK_DATA_CACHE_KEY_V1";
    
    // 강제 갱신이 아닌 경우 캐시 데이터 존재 유무를 확인해 즉시 초고속 리턴!
    if (!forceReload) {
      const cached = getScmFromCache_(CACHE_KEY);
      if (cached) {
        Logger.log("SCM 지정학 대시보드 데이터 캐시 적재 성공!");
        return cached;
      }
    }

    let mapLocations = [];
    let allScmLocations = [];
    let riskStats = [];
    
    // 1) SCM Risk 관리 시트의 '지정학적Risk' B4:F 읽기 (B는 Name, C는 Level, D는 Description, E는 Target Country, F는 Risk Type)
    const riskSs = SpreadsheetApp.openById("1mrMQ7B09ubu_5agloTKMZV_tjQbN0tHzUezEIMMuVko");
    const classSheet = riskSs.getSheetByName("지정학적Risk");
    let activeRisks = [];
    let riskLevelMap = {};
    let riskDescMap = {}; 
    let riskCountryMap = {}; // E열 대상 국가 매핑 추가!
    let riskTypeMap = {}; // F열 시뮬레이션 유형 매핑 추가!
    
    if (classSheet) {
      const lastClassRow = classSheet.getLastRow();
      if (lastClassRow >= 4) {
        // column B (col 2)부터 column F (col 6)까지 5개 컬럼을 인출 (Name, Level, Description, Country, TypeText)
        const classValues = classSheet.getRange(4, 2, lastClassRow - 3, 5).getValues();
        classValues.forEach(r => {
          const name = String(r[0]).trim();
          const level = String(r[1]).trim(); // '상', '중', '하'
          const desc = String(r[2] || '').trim();  // Risk Description!
          const country = String(r[3] || '').trim(); // Target Country (E열)
          const typeText = String(r[4] || '').trim(); // Risk Type (F열)
          if (name) {
            activeRisks.push(name);
            riskLevelMap[name] = level || '하'; // 기본은 '하'로 세팅
            riskDescMap[name] = desc;
            riskCountryMap[name] = country;
            riskTypeMap[name] = typeText;
          }
        });
      }
    }
    
    // 2) SCM Map 시트 '1IcQjU0Ya21SifETdcRFFEB6mBQysXJ5MIflFsFg_hhQ'의 'SCM Raw' 읽기
    const mapSs = SpreadsheetApp.openById("1IcQjU0Ya21SifETdcRFFEB6mBQysXJ5MIflFsFg_hhQ");
    const scmSheet = mapSs.getSheetByName("SCM Raw");
    
    if (scmSheet && activeRisks.length > 0) {
      const scmLastRow = scmSheet.getLastRow();
      if (scmLastRow >= 2) {
        // B열(Vendor Code)부터 Z열(생산유형)까지 25개 열을 로드 (C is index 2, Z is index 26)
        // Row mapping: B(0)=Code, C(1)=Desc, D(2)=Maker, G(5)=Vendor, H(6)=Country, I(7)=City, J(8)=Lat, K(9)=Lon, L(10)=Item, M(11)=Group, O(13)=Risk
        // S(17)부터 Z(24)까지는 생산유형 개별 세부 열들이 포진되어 있습니다.
        const scmRawValues = scmSheet.getRange(2, 2, scmLastRow - 1, 25).getValues();
        
        // 각 Risk별 통계 집계를 위한 임시 맵
        const statsMap = {};
        activeRisks.forEach(r => {
          statsMap[r] = { vendors: new Set(), cities: new Set(), itemGroups: new Set(), items: [] };
        });
        
        scmRawValues.forEach(row => {
          const code = String(row[0] || '').trim();
          const site = String(row[1] || '').trim();
          const vendor = String(row[5] || '').trim();
          const country = String(row[6] || '').trim();
          const city = String(row[7] || '').trim();
          const lat = parseFloat(row[8]);
          const lon = parseFloat(row[9]);
          const item = String(row[10] || '').trim();
          const itemGroup = String(row[11] || '').trim();
          
          // S열(index 17)부터 Z열(index 24)까지의 개별 컬럼 셀 값들을 안전 파싱 수렴하여 생산유형 결합
          const typeVals = [];
          for (let c = 17; c <= 24; c++) {
            const val = String(row[c] || '').trim();
            if (val && val !== '-' && val !== '미분류') {
              typeVals.push(val);
            }
          }
          const typeCombined = typeVals.join(", ");
          
          const riskStr = String(row[13] || '').trim(); // O열 is index 13
          
          // 지진 시뮬레이션용: Risk 분류 여부와 관계없이 실질 좌표가 있는 모든 SCM Location 수합
          if (!isNaN(lat) && !isNaN(lon)) {
            allScmLocations.push({
              code: code,
              site: site,
              vendor: vendor,
              country: country,
              city: city,
              lat: lat,
              lon: lon,
              item: item,
              itemGroup: itemGroup,
              type: typeCombined,
              risk: riskStr || '-'
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
                code: code,
                site: site,
                vendor: vendor,
                country: country,
                city: city,
                lat: lat,
                lon: lon,
                item: item,
                itemGroup: itemGroup,
                type: typeCombined, // 수집된 콤마 결합 생산유형 패킷 주입!
                risk: t
              };
              statsMap[t].items.push(locObj);
              mapLocations.push(locObj);
            }
          });
        });
        
        activeRisks.forEach(r => {
          const entry = statsMap[r];
          if (entry.items.length > 0) {
            riskStats.push({
              name: r,
              level: riskLevelMap[r] || '하', // '상', '중', '하' 이식
              desc: riskDescMap[r] || '',     // 지정학적Risk D열 Description 결합 반환!
              country: riskCountryMap[r] || '', // E열 추가!
              typeText: riskTypeMap[r] || '', // F열 추가!
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
      mapLocations: mapLocations, // 지정학 탭 지도용 교차 마커 리스트 반환
      allScmLocations: allScmLocations, // 지진 시뮬레이션용 전체 마커 리스트 반환
      riskStats: riskStats,       // 지정학 탭 좌측 카테고리 요약표용 통계 반환
      userEmail: context.email,
      userName: context.name,
      userDept: context.text || context.dept
    };
    
    // 새로 읽어온 데이터 세트를 유저 캐시에 고밀도 이중 적재
    putScmToCache_(CACHE_KEY, resultObj);
    
    return resultObj;
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

const PARTNER_DB_ID = '1rQHrhLa3xNZ2DyfK1OS4jeF17xkfXobkV7joMoq4A4w';

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
    const ssId = "1IcQjU0Ya21SifETdcRFFEB6mBQysXJ5MIflFsFg_hhQ"; // SCM Map Spreadsheet ID
    const rawV = SRM_readSheet_('원자재SCM', { id: ssId, ttl: 3600, useDisplay: false });
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
    
    return {
      success: true,
      forecastData: forecastData
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * 전용 회사 환율 관리 워크시트에서 USD, EUR, JPY 환율 데이터를 실시간 가져오는 API
 */
function getCurrencyData() {
  try {
    const ssId = "15mDVNS3jFIX4mNdu0OEHMTaHgDbwvNDSmBp7OxHsH9k";
    const ss = SpreadsheetApp.openById(ssId);
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
          past1M: pastRange[1][0], // J7
          past3M: pastRange[1][1], // K7
          past6M: pastRange[1][2], // L7
          past1Y: pastRange[1][3]  // M7
        },
        jpy: {
          today: jpyToday,
          past1M: pastRange[2][0], // J8
          past3M: pastRange[2][1], // K8
          past6M: pastRange[2][2], // L8
          past1Y: pastRange[2][3]  // M8
        }
      }
    };
  } catch (e) {
    // Re-verify eur past3M alignment
    try {
      const pastRange = sheet.getRange("J6:M8").getValues();
      return {
        success: true,
        data: {
          usd: { today: usdToday, past1M: pastRange[0][0], past3M: pastRange[0][1], past6M: pastRange[0][2], past1Y: pastRange[0][3] },
          eur: { today: eurToday, past1M: pastRange[1][0], past3M: pastRange[1][1], past6M: pastRange[1][2], past1Y: pastRange[1][3] },
          jpy: { today: jpyToday, past1M: pastRange[2][0], past3M: pastRange[2][1], past6M: pastRange[2][2], past1Y: pastRange[2][3] }
        }
      };
    } catch(errInner) {
      return { success: false, message: e.toString() };
    }
  }
}

/**
 * 원가시황Risk 시트에서 민감도, 단가인상 요청내역, 리스크 협력사 데이터를 수합하는 API
 * (Date 객체 등 비직렬화 요소를 문자열로 강제 포맷 변환하여 구글 Apps Script 전송 크래시 완벽 박멸)
 */
function getCostMarketRiskData() {
  try {
    let ss = null;
    let sheet = null;
    
    // 1) 지형학 리스크 스프레드시트 열기 시도
    try {
      const riskSs = SpreadsheetApp.openById("1mrMQ7B09ubu_5agloTKMZV_tjQbN0tHzUezEIMMuVko");
      sheet = riskSs.getSheetByName("원가시황Risk");
      if (sheet) ss = riskSs;
    } catch (e1) {}
    
    // 2) SCM 맵 스프레드시트 열기 시도
    if (!sheet) {
      try {
        const mapSs = SpreadsheetApp.openById("1IcQjU0Ya21SifETdcRFFEB6mBQysXJ5MIflFsFg_hhQ");
        sheet = mapSs.getSheetByName("원가시황Risk");
        if (sheet) ss = mapSs;
      } catch (e2) {}
    }
    
    // 3) G.A.S 활성 현재 스프레드시트 열기 시도 (3중 디펜스)
    if (!sheet) {
      try {
        const activeSs = SpreadsheetApp.getActiveSpreadsheet();
        sheet = activeSs.getSheetByName("원가시황Risk");
        if (sheet) ss = activeSs;
      } catch (e3) {}
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
    
    // 3) 원가·시황측면 Risk 협력사 (AA3:AJ) <-- Row 3 is the header, Row 4+ is content!
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
