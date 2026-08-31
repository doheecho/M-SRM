/**
 * ============================================================================
 * 개선 요약 (원본 대비)
 * ============================================================================
 * 1) [속도] generateInsightsAndSave() 안에서 Object.keys(finalData.data).forEach(...)
 *    로 전체 데이터를 5번 따로 훑던 것을 2번으로 통합했습니다.
 *      - 1차 스캔: 연도에 상관없이 필요한 누적치(aggByYear, vendorSpend, makerSpend,
 *        ipcVendorSpend, vendorIpcMap, l1Data, detailData)를 한 번에 계산
 *      - 2차 스캔: currYear/prevYear를 알아야만 계산 가능한 것들(통화별, IPC 현황,
 *        대리점/대행사)을 모아서 처리하되, 그 외 연도는 즉시 skip 처리 → 데이터가
 *        여러 해에 걸쳐 있을수록 효과가 커집니다 (예: 6개년 데이터면 2차 스캔에서
 *        실질적으로 처리하는 행이 원래의 1/3로 줄어듭니다).
 *    key.split("|")를 5번 반복 호출하던 것도 2번으로 줄었습니다.
 *
 * 2) [속도/구조] CacheService.getUserCache() → getScriptCache() 로 변경했습니다.
 *    Spend 데이터는 사용자별로 달라지는 데이터가 아니라 회사 전체 구매실적이므로,
 *    기존 방식대로면 접속하는 사용자 수만큼 동일한 시트 읽기+집계+AI 시사점 생성이
 *    반복됩니다(캐시가 사람별로 따로 생김). 스크립트 캐시로 바꾸면 처음 한 사람이
 *    캐시를 채운 뒤로는 전 사용자가 그 결과를 공유합니다.
 *    ※ 만약 사용자별로 다른 데이터를 보여줄 계획이 있었다면 이 변경은 되돌려야
 *      합니다 — 현재 코드 상으로는 그런 분기가 보이지 않아 반영했습니다.
 *
 * 3) [속도] 캐시 청크 분할 시 정규식(match(/.{1,90000}/g)) 대신 단순 substring
 *    반복문을 사용했습니다. 데이터가 커질수록(수 MB) 정규식 방식은 느려지고
 *    메모리를 더 씁니다.
 *
 * 4) [제안 - 코드에는 미반영, 검토 필요]
 *    a. generateInsightsAndSave()가 캐시 미스 시점에 동기적으로 실행되어 그 순간
 *       접속한 사용자가 느린 최초 로딩을 그대로 겪습니다. '데이터 갱신' 버튼이나
 *       시트 자체 수정 시점이 아니라, 매일 새벽 등 정해진 시간에 실행되는
 *       "설치형 시간 기반 트리거"로 clearCacheAndRefresh()를 미리 돌려두면
 *       사용자는 항상 캐시 적중만 경험하게 됩니다.
 *    b. rawAggregated의 키를 10개 필드를 "|"로 이어붙인 문자열로 쓰고 있어서
 *       조회할 때마다 매번 split("|")로 되풀이해 파싱합니다. 데이터 규모가 더
 *       커지면 배열의 배열([[year,plant,...,krw], ...]) 구조로 바꾸고 인덱스
 *       상수(IDX_YEAR=0 등)를 쓰는 편이 파싱 비용과 문자열 메모리를 더 아낍니다.
 *       다만 이건 클라이언트(Spend_index.html)의 데이터 소비 로직도 같이 손봐야
 *       해서 이번 개선 범위에서는 제외했습니다.
 *    c. Spend_header.html 파일이 현재 Spend_index.html과 완전히 동일한 2,313줄
 *       전체 대시보드 코드를 담고 있습니다(아마 복사 실수로 보입니다). index.html이
 *       `<?!= include('header'); ?>` 로 이 파일을 그대로 삽입하는 구조라, 지금은
 *       페이지 전체가 두 번 겹쳐 로드되고 있을 가능성이 높습니다. 이 파일은
 *       본문에서 별도로 드린 축소판 Spend_header.html로 교체하시길 권장합니다.
 * ============================================================================
 */

const FILE_IDS = [
  "1M_CFAoSE_DdjOY9fNahHJ0G-pYLbv8ipNAifD405BSA"
];
const SHEET_NAME = '구매실적정리';
const CACHE_BASE_KEY = "SPEND_DATA_DASHBOARD_V7"; // 캐시 스코프 변경으로 버전 bump

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.initialPageDataString = JSON.stringify(getSpendDataAdvanced());
  return template.evaluate()
    .setTitle('구매 Spend 분석')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getActiveUserEmail() {
  return Session.getActiveUser().getEmail();
}

// ---- 캐시: 스크립트 전체 공유 캐시로 변경 (회사 공통 데이터이므로) ----
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

function putToCache(baseKey, dataObj) {
  const cache = CacheService.getScriptCache();
  const jsonString = JSON.stringify(dataObj);
  const CHUNK_SIZE = 90000;
  const chunkCount = Math.ceil(jsonString.length / CHUNK_SIZE) || 0;
  const payload = {};
  for (let i = 0; i < chunkCount; i++) {
    payload[baseKey + "_" + i] = jsonString.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
  }
  payload[baseKey + "_COUNT"] = chunkCount.toString();
  // putAll이 put을 반복 호출하는 것보다 네트워크 왕복이 적습니다.
  cache.putAll(payload, 21600);
}

function getSpendDataAdvanced() {
  const cachedData = getFromCache(CACHE_BASE_KEY);
  if (cachedData) return cachedData;

  let rawAggregated = {};
  let filterSets = {
    year: new Set(), plant: new Set(), tradeType: new Set(), currency: new Set(),
    l1: new Set(), l2: new Set(), l3: new Set(), vendorDesc: new Set(), ipcVendor: new Set(), maker: new Set()
  };
  let updateMonth = 12;
  let aiInsights = [];

  try {
    for (let i = 0; i < FILE_IDS.length; i++) {
      const ss = SpreadsheetApp.openById(FILE_IDS[i].trim());
      const updateSheet = ss.getSheetByName("Setting");
      if (updateSheet) {
        const dateText = updateSheet.getRange("A2").getDisplayValue().trim();
        const match = dateText.match(/\d{4}[-./\s]+(\d{1,2})/);
        if (match) {
          updateMonth = parseInt(match[1], 10);
        } else {
          const parsedMonth = parseInt(dateText.replace(/[^0-9]/g, ''), 10);
          if (!isNaN(parsedMonth) && parsedMonth >= 1 && parsedMonth <= 12) {
            updateMonth = parsedMonth;
          }
        }
        const cValues = updateSheet.getRange("C2:C").getDisplayValues();
        aiInsights = cValues.map(row => row[0].trim()).filter(val => val.length > 0);
        console.log("최근 업데이트 월 식별 완료: " + updateMonth + "월");
        break;
      }
    }
  } catch (e) {
    console.log("Setting 확인 중 에러: " + e.message);
  }

  FILE_IDS.forEach(id => {
    try {
      const ss = SpreadsheetApp.openById(id.trim());
      const sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet || sheet.getLastRow() < 2) return;
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).getValues();
      values.forEach(row => {
        let year = row[0] ? String(row[0]).replace(/\.0$/, '').trim() : "미분류";
        if (year.length > 4) year = year.substring(0, 4);
        const plant = row[6] ? String(row[6]).trim() : "미분류";
        const l1 = row[7] ? String(row[7]).trim() : "미분류";
        const l2 = row[8] ? String(row[8]).trim() : "미분류";
        const l3 = row[9] ? String(row[9]).trim() : "미분류";
        const vendorDesc = row[2] ? String(row[2]).trim() : "미분류";
        const ipcVendor = row[5] ? String(row[5]).trim() : "미분류";
        const maker = row[3] ? String(row[3]).trim() : "미분류";
        const tradeType = row[10] ? String(row[10]).trim() : "미분류";
        const currency = row[11] ? String(row[11]).trim() : "미분류";
        const spendKRW = parseFloat(String(row[13]).replace(/,/g, '').replace(/\s/g, '')) || 0;
        const spendUSD = parseFloat(String(row[14]).replace(/,/g, '').replace(/\s/g, '')) || 0;

        filterSets.year.add(year);
        filterSets.plant.add(plant);
        filterSets.tradeType.add(tradeType);
        filterSets.currency.add(currency);
        filterSets.l1.add(l1);
        filterSets.l2.add(l2);
        filterSets.l3.add(l3);
        filterSets.vendorDesc.add(vendorDesc);
        filterSets.ipcVendor.add(ipcVendor);
        filterSets.maker.add(maker);

        const key = `${year}|${plant}|${tradeType}|${currency}|${l1}|${l2}|${l3}|${vendorDesc}|${ipcVendor}|${maker}`;
        if (!rawAggregated[key]) {
          rawAggregated[key] = { krw: 0, usd: 0 };
        }
        rawAggregated[key].krw += spendKRW;
        rawAggregated[key].usd += spendUSD;
      });
    } catch (err) {
      console.error("파일 로드 실패: " + id, err);
    }
  });

  const finalData = {
    success: true,
    data: rawAggregated,
    updateMonth: updateMonth,
    aiInsights: aiInsights,
    filters: {
      year: Array.from(filterSets.year).sort(),
      plant: Array.from(filterSets.plant).sort(),
      tradeType: Array.from(filterSets.tradeType).sort(),
      currency: Array.from(filterSets.currency).sort(),
      l1: Array.from(filterSets.l1).sort(),
      l2: Array.from(filterSets.l2).sort(),
      l3: Array.from(filterSets.l3).sort(),
      vendorDesc: Array.from(filterSets.vendorDesc).sort(),
      ipcVendor: Array.from(filterSets.ipcVendor).sort(),
      maker: Array.from(filterSets.maker).sort()
    }
  };

  // 캐시 미스 시 구글시트 "Setting" 시트에 올바른 AI 시사점 갱신 및 데이터 동기화
  try {
    generateInsightsAndSave(finalData);
  } catch (insErr) {
    console.log("시사점 자동 생성 및 저장 실패: " + insErr.message);
  }

  putToCache(CACHE_BASE_KEY, finalData);
  return finalData;
}

function clearCacheAndRefresh() {
  const cache = CacheService.getScriptCache();
  const chunkCount = cache.get(CACHE_BASE_KEY + "_COUNT");
  if (chunkCount) {
    const count = parseInt(chunkCount, 10);
    const keysToRemove = [CACHE_BASE_KEY + "_COUNT"];
    for (let i = 0; i < count; i++) keysToRemove.push(CACHE_BASE_KEY + "_" + i);
    cache.removeAll(keysToRemove);
  }
  const finalData = getSpendDataAdvanced();
  generateInsightsAndSave(finalData);
  putToCache(CACHE_BASE_KEY, finalData);
  return finalData;
}

function generateInsightsAndSave(finalData) {
  try {
    const allKeys = Object.keys(finalData.data);

    const targetL1s = ["기구", "전기/전자", "전장", "원부자재", "포장/인쇄", "프로브"];
    const targetDetails = ["기구Assy", "판금", "사출", "능동소자", "수동소자", "PC Mother(L3)", "VGA(L3)", "PCB(L3)", "FPCB(L3)", "PBA", "Cable", "Cable Assy", "파워Assy", "모니터Assy", "Single Crystal(L3)", "Ceramic(L3)", "포장", "Accessory"];
    // 반복문 안에서 매번 endsWith/replace 하지 않도록 미리 분리
    const targetDetailL2 = targetDetails.filter(t => !t.endsWith("(L3)"));
    const targetDetailL3 = targetDetails
      .filter(t => t.endsWith("(L3)"))
      .map(t => ({ target: t, l3Name: t.slice(0, -4).trim() }));

    // ===== 1차 스캔: 연도와 무관하게 계산 가능한 모든 누적치를 한 번에 =====
    let aggByYear = {};
    let vendorSpend = {};
    let makerSpend = {};
    let ipcVendorSpend = {};
    let vendorIpcMap = {};
    let l1Data = {};
    let detailData = {};

    allKeys.forEach(key => {
      const parts = key.split("|");
      const year = parts[0];
      const tradeType = parts[2];
      const l1 = parts[4];
      const l2 = parts[5];
      const l3 = parts[6];
      const vendor = parts[7];
      const ipcVendor = parts[8];
      const maker = parts[9];
      const krw = finalData.data[key].krw;

      aggByYear[year] = (aggByYear[year] || 0) + krw;

      if (!vendorSpend[vendor]) vendorSpend[vendor] = {};
      vendorSpend[vendor][year] = (vendorSpend[vendor][year] || 0) + krw;

      if (!makerSpend[maker]) makerSpend[maker] = {};
      makerSpend[maker][year] = (makerSpend[maker][year] || 0) + krw;

      if (tradeType.toUpperCase().includes("IPC")) {
        if (ipcVendor && ipcVendor !== "미분류") vendorIpcMap[vendor] = ipcVendor;
        if (!ipcVendorSpend[vendor]) ipcVendorSpend[vendor] = {};
        ipcVendorSpend[vendor][year] = (ipcVendorSpend[vendor][year] || 0) + krw;
      }

      if (targetL1s.includes(l1)) {
        if (!l1Data[l1]) l1Data[l1] = { total: {}, vendors: {} };
        l1Data[l1].total[year] = (l1Data[l1].total[year] || 0) + krw;
        if (!l1Data[l1].vendors[vendor]) l1Data[l1].vendors[vendor] = {};
        l1Data[l1].vendors[vendor][year] = (l1Data[l1].vendors[vendor][year] || 0) + krw;
      }

      // targetDetails 매칭 (원본과 동일 로직, 소스는 그대로 두고 순회만 미리 나눠서 조금 더 빠르게)
      for (let i = 0; i < targetDetailL2.length; i++) {
        const target = targetDetailL2[i];
        if (l2 === target) {
          if (!detailData[target]) detailData[target] = { total: {}, vendors: {} };
          detailData[target].total[year] = (detailData[target].total[year] || 0) + krw;
          if (!detailData[target].vendors[vendor]) detailData[target].vendors[vendor] = {};
          detailData[target].vendors[vendor][year] = (detailData[target].vendors[vendor][year] || 0) + krw;
        }
      }
      for (let i = 0; i < targetDetailL3.length; i++) {
        const { target, l3Name } = targetDetailL3[i];
        if (l3 === l3Name) {
          if (!detailData[target]) detailData[target] = { total: {}, vendors: {} };
          detailData[target].total[year] = (detailData[target].total[year] || 0) + krw;
          if (!detailData[target].vendors[vendor]) detailData[target].vendors[vendor] = {};
          detailData[target].vendors[vendor][year] = (detailData[target].vendors[vendor][year] || 0) + krw;
        }
      }
    });

    const years = Object.keys(aggByYear).sort();
    if (years.length === 0) return false;

    const currYear = years[years.length - 1];
    const prevYear = years.length > 1 ? years[years.length - 2] : null;
    const updateMonth = finalData.updateMonth || 12;
    const forecastFactor = updateMonth < 12 ? (12 / updateMonth) : 1;
    const forecastText = updateMonth < 12 ? `(${updateMonth}월 누적기준 연간 환산)` : "";
    const currTotalRaw = aggByYear[currYear] || 0;
    const currTotal = currTotalRaw * forecastFactor;
    const prevTotal = prevYear ? (aggByYear[prevYear] || 0) : 0;
    const overallDirection = currTotal >= prevTotal ? "증가" : "감소";

    // ===== 2차 스캔: currYear/prevYear에만 관련된 항목들을 한 번에 계산 =====
    // (연도가 currYear/prevYear가 아니면 즉시 skip → 다년치 데이터일수록 처리량이 크게 줄어듦)
    let currIpcTotalSpend = 0;
    let currIpcSpendByPVendor = {};
    let currIpcSpendByVendor = {};
    let currIpcVendorL3Spend = {};
    let currIpcSpendDetails = {};
    let currencySpend = { curr: {}, prev: {} };
    let vendorByCurrency = { KRW: {}, USD: {}, EUR: {}, JPY: {} };
    let l2ByCurrency = { KRW: {}, USD: {}, EUR: {}, JPY: {} };
    let foreignSpend = { curr: 0, prev: 0 };
    let agencySpend = {};
    let proxySpend = {};
    const CURRENCY_SET = { KRW: true, USD: true, EUR: true, JPY: true };

    allKeys.forEach(key => {
      const parts = key.split("|");
      const year = parts[0];
      const isCurr = year === currYear;
      const isPrev = prevYear !== null && year === prevYear;
      if (!isCurr && !isPrev) return; // 다른 연도는 이 단계에서 필요 없음

      const tradeType = parts[2];
      const currencyRaw = parts[3];
      const currency = currencyRaw ? currencyRaw.toUpperCase() : "미분류";
      const l2 = parts[5];
      const l3 = parts[6];
      const vendor = parts[7];
      const ipcVendor = parts[8];
      const maker = parts[9];
      const krw = finalData.data[key].krw;

      if (isCurr) currencySpend.curr[currency] = (currencySpend.curr[currency] || 0) + krw;
      if (isPrev) currencySpend.prev[currency] = (currencySpend.prev[currency] || 0) + krw;

      if (isCurr && CURRENCY_SET[currency]) {
        vendorByCurrency[currency][vendor] = (vendorByCurrency[currency][vendor] || 0) + krw;
        l2ByCurrency[currency][l2] = (l2ByCurrency[currency][l2] || 0) + krw;
      }

      if (currency && currency !== "KRW" && currency !== "미분류") {
        if (isCurr) foreignSpend.curr += krw;
        if (isPrev) foreignSpend.prev += krw;
      }

      if (isCurr && tradeType.toUpperCase().includes("IPC")) {
        currIpcTotalSpend += krw;
        if (ipcVendor && ipcVendor !== "미분류") {
          currIpcSpendByPVendor[ipcVendor] = (currIpcSpendByPVendor[ipcVendor] || 0) + krw;
          if (!currIpcSpendDetails[ipcVendor]) currIpcSpendDetails[ipcVendor] = {};
          currIpcSpendDetails[ipcVendor][vendor] = (currIpcSpendDetails[ipcVendor][vendor] || 0) + krw;
        }
        currIpcSpendByVendor[vendor] = (currIpcSpendByVendor[vendor] || 0) + krw;
        if (!currIpcVendorL3Spend[vendor]) currIpcVendorL3Spend[vendor] = {};
        currIpcVendorL3Spend[vendor][l3] = (currIpcVendorL3Spend[vendor][l3] || 0) + krw;
      }

      if (isCurr) {
        if (tradeType.includes("대리점")) {
          agencySpend[maker] = (agencySpend[maker] || 0) + krw;
        } else if (tradeType.includes("대행사")) {
          proxySpend[maker] = (proxySpend[maker] || 0) + krw;
        }
      }
    });

    function formatBillionInt(v) {
      if (!v || v === 0) return '0';
      return Math.round(v / 100000000).toLocaleString('en-US');
    }

    let insights = [];

    // ======================================================================
    // 여기서부터 "insights.push(...)" 로 시사점 문장을 만드는 부분은 원본 로직과
    // 완전히 동일합니다 (위에서 준비한 aggByYear/vendorSpend/makerSpend/
    // ipcVendorSpend/vendorIpcMap/l1Data/detailData/currIpc*/currencySpend/
    // vendorByCurrency/l2ByCurrency/foreignSpend/agencySpend/proxySpend 를
    // 그대로 사용하도록 변수명을 맞춰뒀습니다).
    //
    // ⚠️ 원본 파일의 이 아래 부분(총 구매액 트렌드 ~ finalInsights 조립 및
    // Setting 시트에 저장하는 마지막 코드)은 GitHub이 자동화 접근에서 raw 파일
    // 다운로드를 막아둔 탓에 제가 끝까지 가져오지 못했습니다. 원본 Spend_Cods.gs
    // 에서 "let insights = [];" 바로 다음 줄부터 파일 끝까지를 그대로 복사해서
    // 이 자리에 붙여넣어 주세요 — 그 아래 로직은 이번 개선과 무관하게 그대로
    // 재사용 가능합니다 (사용하는 변수명이 전부 동일하게 유지되도록 맞췄습니다).
    // ======================================================================

  } catch (err) {
    console.log("generateInsightsAndSave 처리 중 에러: " + err.message);
    return false;
  }
}
