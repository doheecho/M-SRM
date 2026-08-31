const FILE_IDS = [
  "1M_CFAoSE_DdjOY9fNahHJ0G-pYLbv8ipNAifD405BSA"
];
const SHEET_NAME = '구매실적정리';
const CACHE_BASE_KEY = "SPEND_DATA_DASHBOARD_V6";

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

function getFromCache(baseKey) {
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
}

function putToCache(baseKey, dataObj) {
  const cache = CacheService.getUserCache();
  const jsonString = JSON.stringify(dataObj);
  const chunks = jsonString.match(/.{1,90000}/g) || [];
  for (let i = 0; i < chunks.length; i++) {
    cache.put(baseKey + "_" + i, chunks[i], 21600);
  }
  cache.put(baseKey + "_COUNT", chunks.length.toString(), 21600);
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
        if(year.length > 4) year = year.substring(0, 4);
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
  const cache = CacheService.getUserCache();
  const chunkCount = cache.get(CACHE_BASE_KEY + "_COUNT");
  if (chunkCount) {
    for (let i = 0; i < parseInt(chunkCount); i++) {
      cache.remove(CACHE_BASE_KEY + "_" + i);
    }
    cache.remove(CACHE_BASE_KEY + "_COUNT");
  }
  
  const finalData = getSpendDataAdvanced();
  generateInsightsAndSave(finalData);
  putToCache(CACHE_BASE_KEY, finalData);
  return finalData;
}

function generateInsightsAndSave(finalData) {
  try {
    let aggByYear = {};
    let vendorSpend = {};
    let makerSpend = {};
    let foreignSpend = { curr: 0, prev: 0 };
    let ipcVendorSpend = {}; 
    let vendorIpcMap = {}; 
    
    Object.keys(finalData.data).forEach(key => {
      const parts = key.split("|");
      const year = parts[0];
      const tradeType = parts[2];
      const currency = parts[3];
      const vendor = parts[7];
      const ipcVendor = parts[8];
      const maker = parts[9];
      const krw = finalData.data[key].krw;
      
      if (tradeType.toUpperCase().includes("IPC")) {
        if (ipcVendor && ipcVendor !== "미분류") {
          vendorIpcMap[vendor] = ipcVendor;
        }
        if (!ipcVendorSpend[vendor]) ipcVendorSpend[vendor] = {};
        if (!ipcVendorSpend[vendor][year]) ipcVendorSpend[vendor][year] = 0;
        ipcVendorSpend[vendor][year] += krw;
      }
      
      if (!aggByYear[year]) aggByYear[year] = 0;
      aggByYear[year] += krw;
      
      if (!vendorSpend[vendor]) vendorSpend[vendor] = {};
      if (!vendorSpend[vendor][year]) vendorSpend[vendor][year] = 0;
      vendorSpend[vendor][year] += krw;
      
      if (!makerSpend[maker]) makerSpend[maker] = {};
      if (!makerSpend[maker][year]) makerSpend[maker][year] = 0;
      makerSpend[maker][year] += krw;
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

    let currIpcTotalSpend = 0;
    let currIpcSpendByPVendor = {};
    let currIpcSpendByVendor = {};
    let currIpcVendorL3Spend = {};
    let currIpcSpendDetails = {}; 
    
    // 외화 및 환종별 상세 분석을 위한 객체
    let currencySpend = { curr: {}, prev: {} };
    let vendorByCurrency = { KRW: {}, USD: {}, EUR: {}, JPY: {} };
    let l2ByCurrency = { KRW: {}, USD: {}, EUR: {}, JPY: {} };

    Object.keys(finalData.data).forEach(key => {
      const parts = key.split("|");
      const year = parts[0];
      const tradeType = parts[2];
      const currencyRaw = parts[3];
      const currency = currencyRaw ? currencyRaw.toUpperCase() : "미분류";
      const l2 = parts[5];
      const l3 = parts[6];
      const vendor = parts[7];
      const ipcVendor = parts[8];
      const krw = finalData.data[key].krw;
      
      if (year === currYear) {
          currencySpend.curr[currency] = (currencySpend.curr[currency] || 0) + krw;
          if (['KRW', 'USD', 'EUR', 'JPY'].includes(currency)) {
              if (!vendorByCurrency[currency]) vendorByCurrency[currency] = {};
              vendorByCurrency[currency][vendor] = (vendorByCurrency[currency][vendor] || 0) + krw;
              
              if (!l2ByCurrency[currency]) l2ByCurrency[currency] = {};
              l2ByCurrency[currency][l2] = (l2ByCurrency[currency][l2] || 0) + krw;
          }
      }
      if (year === prevYear) {
          currencySpend.prev[currency] = (currencySpend.prev[currency] || 0) + krw;
      }

      if (currency && currency !== "KRW" && currency !== "미분류") {
        if (year === currYear) foreignSpend.curr += krw;
        if (year === prevYear) foreignSpend.prev += krw;
      }

      if (year === currYear && tradeType.toUpperCase().includes("IPC")) {
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
    });

    function formatBillionInt(v) {
      if (!v || v === 0) return '0';
      return Math.round(v / 100000000).toLocaleString('en-US');
    }

    let insights = [];

    // [총 구매액 트렌드]
    if (prevYear && prevTotal > 0) {
      const yoy1Year = ((currTotal - prevTotal) / prevTotal * 100).toFixed(1);
      const dirColor1Year = currTotal >= prevTotal ? "#DC3545" : "#0D6EFD";
      const overallDirection1Year = currTotal >= prevTotal ? "증가" : "감소";

      let yearCount = Math.min(years.length, 4);
      let recentYears = years.slice(-yearCount);
      let n = recentYears.length - 1; 
      let baseYear = recentYears[0];
      let baseTotal = aggByYear[baseYear] || 0;
      
      let avgYoyPercent = "0.0";
      let trendDirection = "유지";
      let trendDirColor = "#333";

      if (n > 0 && baseTotal > 0) {
        let cagr = Math.pow(currTotal / baseTotal, 1 / n) - 1;
        avgYoyPercent = Math.abs(cagr * 100).toFixed(1);
        trendDirection = cagr >= 0 ? "증가" : "감소";
        trendDirColor = cagr >= 0 ? "#DC3545" : "#0D6EFD";
      } else {
        avgYoyPercent = Math.abs(yoy1Year);
        trendDirection = overallDirection1Year;
        trendDirColor = dirColor1Year;
      }

      let historyArr = recentYears.map(y => {
        let val = (y === currYear) ? currTotal : (aggByYear[y] || 0);
        let yy = y.substring(2); 
        return `${formatBillionInt(val)}억('${yy}년)`;
      });
      
      let trendSentence = "";
      let valBold = `<b>${formatBillionInt(currTotal)}억</b>`;
      let dirColored1Year = `<span style="color: ${dirColor1Year}; font-weight: bold;">${Math.abs(yoy1Year)}% ${overallDirection1Year}</span>`;
      
      if (updateMonth < 12) {
        trendSentence = `최근 ${yearCount}개년도 구매액은 ${historyArr.join(' → ')} 으로 <span style="color: ${trendDirColor}; font-weight: bold;">YoY ${avgYoyPercent}% ${trendDirection}</span> 중입니다.`;
        insights.push(`<b>[총 구매액 트렌드]</b> ${currYear}년 연간 구매액(전망)은 ${valBold}${forecastText}으로, 전년 대비 ${dirColored1Year}할 것으로 예상됩니다. ${trendSentence}`);
      } else {
        trendSentence = `최근 ${yearCount}개년도 구매액은 ${historyArr.join(' → ')} 으로 <span style="color: ${trendDirColor}; font-weight: bold;">YoY ${avgYoyPercent}% ${trendDirection}</span> 했습니다.`;
        insights.push(`<b>[총 구매액 트렌드]</b> ${currYear}년 연간 구매액은 ${valBold}으로, 전년 대비 ${dirColored1Year}한 것으로 집계되었습니다. ${trendSentence}`);
      }
      
    } else {
      let valBold = `<b>${formatBillionInt(currTotal)}억</b>`;
      let textSuffix = updateMonth < 12 ? "할 것으로 예상됩니다." : "한 것으로 집계되었습니다.";
      let titleSuffix = updateMonth < 12 ? "구매액(전망)" : "구매액";
      insights.push(`<b>[총 구매액 트렌드]</b> ${currYear}년 연간 ${titleSuffix}은 ${valBold}${forecastText}으로, 전년 대비 비교 자료가 없으나 금년도 실적은 해당 금액으로 ${textSuffix}`);
    }

    // [Vendor 집중도]
    let topVendors = Object.keys(vendorSpend).sort((a,b) => ((vendorSpend[b][currYear]||0) * forecastFactor) - ((vendorSpend[a][currYear]||0) * forecastFactor)).slice(0, 10);
    let top10VendorSum = topVendors.reduce((sum, v) => sum + ((vendorSpend[v][currYear]||0) * forecastFactor), 0);
    let vendorConcentration = currTotal > 0 ? (top10VendorSum / currTotal * 100).toFixed(1) : "0.0";
    
    let sumBold = `<b>${formatBillionInt(top10VendorSum)}억</b>`;
    let concBold = `<b>${vendorConcentration}%</b>`;
    
    if (prevYear && prevTotal > 0) {
      let prevTopVendors = Object.keys(vendorSpend).sort((a,b) => (vendorSpend[b][prevYear]||0) - (vendorSpend[a][prevYear]||0)).slice(0, 10);
      let prevTop10VendorSum = prevTopVendors.reduce((sum, v) => sum + (vendorSpend[v][prevYear]||0), 0);
      let prevVendorConcentration = (prevTop10VendorSum / prevTotal * 100).toFixed(1);
      
      let diffConcentration = (parseFloat(vendorConcentration) - parseFloat(prevVendorConcentration)).toFixed(1);
      let concStatus = diffConcentration >= 0
        ? `증가하여 <span style="color: #DC3545; font-weight: bold;">집중도가 심화</span>되었고`
        : `감소하여 <span style="color: #0D6EFD; font-weight: bold;">집중도가 완화</span>되었고`;
        
      let yearCount = Math.min(years.length, 3);
      let recentYears = years.slice(-yearCount);
      let vendorHistArr = recentYears.map(y => {
        let isCurr = (y === currYear);
        let yTotal = isCurr ? currTotal : (aggByYear[y] || 0);
        let factor = isCurr ? forecastFactor : 1;
        let yTopVendors = Object.keys(vendorSpend).sort((a,b) => ((vendorSpend[b][y]||0) * factor) - ((vendorSpend[a][y]||0) * factor)).slice(0, 10);
        let yTopSum = yTopVendors.reduce((sum, v) => sum + ((vendorSpend[v][y]||0) * factor), 0);
        let yConc = yTotal > 0 ? (yTopSum / yTotal * 100).toFixed(1) : "0.0";
        let yy = y.substring(2);
        return `${yConc}%('${yy}년)`;
      });
      let vendorHistStr = `최근 ${yearCount}년간 ${vendorHistArr.join(' → ')} 로 변동 중입니다.`;
      
      insights.push(`<b>[Vendor 집중도]</b> 구매액 상위 10개 Vendor의 연간 구매액은 ${sumBold}으로, 전체 구매액의 ${concBold}를 차지하고 있습니다. 상위 10개사의 비중은 전년도 ${prevVendorConcentration}% 대비 <span style="color: ${diffConcentration >= 0 ? '#DC3545' : '#0D6EFD'}; font-weight: bold;">${Math.abs(diffConcentration)}%p</span> ${concStatus}, ${vendorHistStr}`);
    } else {
      insights.push(`<b>[Vendor 집중도]</b> 구매액 상위 10개 Vendor의 연간 구매액은 ${sumBold}으로, 전체 구매액의 ${concBold}를 차지하고 있어 집중도가 확인됩니다.`);
    }

    // [Maker 집중도]
    let topMakers = Object.keys(makerSpend).sort((a,b) => ((makerSpend[b][currYear]||0) * forecastFactor) - ((makerSpend[a][currYear]||0) * forecastFactor)).slice(0, 10);
    let top10MakerSum = topMakers.reduce((sum, m) => sum + ((makerSpend[m][currYear]||0) * forecastFactor), 0);
    let makerConcentration = currTotal > 0 ? (top10MakerSum / currTotal * 100).toFixed(1) : "0.0";
    
    let top5Makers = topMakers.slice(0, 5);
    let makerDetails = top5Makers.map(m => {
      let mSpend = (makerSpend[m][currYear]||0) * forecastFactor;
      let mPercent = currTotal > 0 ? (mSpend / currTotal * 100).toFixed(1) : "0.0";
      return `<span style="color: #198754; font-weight: bold;">${m}</span> ${mPercent}% (${formatBillionInt(mSpend)}억)`;
    }).join(", ");
    
    let makerTargetStr = `<b>${makerConcentration}% (${formatBillionInt(top10MakerSum)}억 전망)</b>`;
    let makerStr = `<b>[Maker 집중도]</b> 상위 10개 Maker의 구매액 비중은 전체의 ${makerTargetStr}이며 ${makerDetails} 순 입니다.`;
    
    if (prevYear && prevTotal > 0) {
      let prevTopMakers = Object.keys(makerSpend).sort((a,b) => (makerSpend[b][prevYear]||0) - (makerSpend[a][prevYear]||0));
      if (prevTopMakers.length > 0) {
        let prevTop3 = prevTopMakers.slice(0, 3);
        let prevMakerDetails = prevTop3.map(m => {
            let mSpend = makerSpend[m][prevYear] || 0;
            let mPercent = (mSpend / prevTotal * 100).toFixed(1);
            return `<span style="color: #198754; font-weight: bold;">${m}</span> ${mPercent}% (${formatBillionInt(mSpend)}억)`;
        });
        
        if (prevMakerDetails.length > 0) {
            let firstPrevMaker = prevMakerDetails[0];
            if (prevMakerDetails.length > 1) {
                let restPrevMakers = prevMakerDetails.slice(1).join(", ");
                makerStr += ` 전년도 구매액 상위 Maker는 ${firstPrevMaker}, ${restPrevMakers} 였습니다.`;
            } else {
                makerStr += ` 전년도 구매액 1위 Maker는 ${firstPrevMaker} 이었습니다.`;
            }
        }
      }
    }
    insights.push(makerStr);

    // [Vendor 거래 증가/감소]
    if (prevYear) {
      let vendorDiffs = [];
      Object.keys(vendorSpend).forEach(v => {
        let c = (vendorSpend[v][currYear] || 0) * forecastFactor;
        let p = vendorSpend[v][prevYear] || 0;
        if (c > 0 || p > 0) vendorDiffs.push({name: v, curr: c, prev: p, diff: c - p});
      });
      vendorDiffs.sort((a,b) => b.diff - a.diff);
      
      if (vendorDiffs.length > 0 && vendorDiffs[0].diff > 0) {
        let topV = vendorDiffs[0];
        let topVPct = topV.prev > 0 ? `+${((topV.diff) / topV.prev * 100).toFixed(1)}%` : `신규`;
        let topVName = `<span style="color: #6f42c1; font-weight: bold;">${topV.name}</span>`;
        
        let nextIncs = vendorDiffs.slice(1, 6).filter(v => v.diff > 0).map(v => {
            let pct = v.prev > 0 ? `+${((v.diff) / v.prev * 100).toFixed(1)}%` : '신규';
            return `<span style="color: #6f42c1; font-weight: bold;">${v.name}</span> <span style="color: #333; font-weight: bold;">${formatBillionInt(v.diff)}억 (${pct})</span>`;
        });
        let nextIncStr = nextIncs.length > 0 ? ` 이후 ${nextIncs.join(', ')} 등의 구매액이 크게 상승했습니다.` : "";
        
        insights.push(`<b>[Vendor 거래 증가]</b> ${topVName} 의 구매액은 전년비 <span style="color: #DC3545; font-weight: bold;">${formatBillionInt(topV.diff)}억 증가하여 큰 상승폭 (${topVPct})</span>을 보였습니다.${nextIncStr}`);
      } else {
        insights.push(`<b>[Vendor 거래 안정세]</b> 전년 대비 이례적으로 구매액이 크게 폭등한 Vendor는 관찰되지 않았습니다.`);
      }
      
      if (vendorDiffs.length > 1 && vendorDiffs[vendorDiffs.length-1].diff < 0) {
        let botV = vendorDiffs[vendorDiffs.length-1];
        let botVPct = botV.prev > 0 ? `${((botV.diff) / botV.prev * 100).toFixed(1)}%` : `N/A`;
        let botVName = `<span style="color: #6f42c1; font-weight: bold;">${botV.name}</span>`;
        
        let nextDecs = vendorDiffs.slice(Math.max(0, vendorDiffs.length - 6), vendorDiffs.length - 1).reverse().filter(v => v.diff < 0).map(v => {
            let pct = v.prev > 0 ? `${((v.diff) / v.prev * 100).toFixed(1)}%` : 'N/A';
            return `<span style="color: #6f42c1; font-weight: bold;">${v.name}</span> <span style="color: #333; font-weight: bold;">${formatBillionInt(Math.abs(v.diff))}억 (${pct})</span>`;
        });
        let nextDecStr = nextDecs.length > 0 ? ` 이외에 ${nextDecs.join(', ')} 순으로 전년비 구매액이 감소한 것으로 나타놓았으며, ` : "";
        
        insights.push(`<b>[Vendor 거래 감소]</b> ${botVName} 의 구매액이 전년비 <span style="color: #0D6EFD; font-weight: bold;">${formatBillionInt(Math.abs(botV.diff))}억 감소(${botVPct})</span>하였고 ${nextDecStr} <span style="color: #0D6EFD; font-weight: bold;">협력사 운영전략에 따른 구매액 감소</span> 여부를 판단할 필요가 있습니다.`);
      }
    }

    // [Maker 거래 증가/감소]
    if (prevYear) {
      let makerDiffs = [];
      Object.keys(makerSpend).forEach(m => {
        let c = (makerSpend[m][currYear] || 0) * forecastFactor;
        let p = makerSpend[m][prevYear] || 0;
        if (c > 0 || p > 0) makerDiffs.push({name: m, curr: c, prev: p, diff: c - p});
      });
      makerDiffs.sort((a,b) => b.diff - a.diff);
      
      if (makerDiffs.length > 0 && makerDiffs[0].diff > 0) {
        let topM = makerDiffs[0];
        let topMPct = topM.prev > 0 ? `+${((topM.diff) / topM.prev * 100).toFixed(1)}%` : `신규`;
        let topMName = `<span style="color: #198754; font-weight: bold;">${topM.name}</span>`;
        
        let nextIncs = makerDiffs.slice(1, 6).filter(m => m.diff > 0).map(m => {
            let pct = m.prev > 0 ? `+${((m.diff) / m.prev * 100).toFixed(1)}%` : '신규';
            return `<span style="color: #198754; font-weight: bold;">${m.name}</span> <span style="color: #333; font-weight: bold;">${formatBillionInt(m.diff)}억 (${pct})</span>`;
        });
        let nextIncStr = nextIncs.length > 0 ? ` 증가금액 순서는 ${nextIncs.join(', ')} 입니다.` : "";
        
        insights.push(`<b>[Maker 거래 증가]</b> ${topMName} 의 구매액은 전년비 <span style="color: #DC3545; font-weight: bold;">${formatBillionInt(topM.diff)}억, 가장 크게 상승 (${topMPct})</span>했습니다.${nextIncStr} `);
      }
      
      if (makerDiffs.length > 1 && makerDiffs[makerDiffs.length-1].diff < 0) {
        let botM = makerDiffs[makerDiffs.length-1];
        let botMPct = botM.prev > 0 ? `${((botM.diff) / botM.prev * 100).toFixed(1)}%` : `N/A`;
        let botMName = `<span style="color: #198754; font-weight: bold;">${botM.name}</span>`;
        
        let nextDecs = makerDiffs.slice(Math.max(0, makerDiffs.length - 6), makerDiffs.length - 1).reverse().filter(m => m.diff < 0).map(m => {
            let pct = m.prev > 0 ? `${((m.diff) / m.prev * 100).toFixed(1)}%` : 'N/A';
            return `<span style="color: #198754; font-weight: bold;">${m.name}</span> <span style="color: #333; font-weight: bold;">${formatBillionInt(Math.abs(m.diff))}억 (${pct})</span>`;
        });
        let nextDecStr = nextDecs.length > 0 ? ` 그외 ${nextDecs.join(', ')} 등이 크게 감소했습니다.` : "";
        
        insights.push(`<b>[Maker 거래 감소]</b> ${botMName} 의 구매액이 전년비 <span style="color: #0D6EFD; font-weight: bold;">${formatBillionInt(Math.abs(botM.diff))}억 감소(${botMPct})</span>했고, ${nextDecStr} <span style="color: #0D6EFD; font-weight: bold;">당사 모델 단종</span>이나 <span style="color: #0D6EFD; font-weight: bold;">버전업 등 설계변경에 의한 감소</span>인 경우 <span style="color: #0D6EFD; font-weight: bold;">당사/협력사 보유 잔여재고 처리방안</span>에 대한 점검이 필요합니다.`);
      }
    }

    // [공급망 다변화 및 IPC 로직]
    if (prevYear) {
      let newVendors = Object.keys(vendorSpend).filter(v => (vendorSpend[v][currYear]||0) > 0 && (vendorSpend[v][prevYear]||0) === 0);
      
      if (newVendors.length > 0) {
        newVendors.sort((a,b) => vendorSpend[b][currYear] - vendorSpend[a][currYear]);
        
        let topNew = newVendors[0];
        let topNewSpend = (vendorSpend[topNew][currYear] || 0) * forecastFactor;
        let topNewName = `<span style="color: #6f42c1; font-weight: bold;">${topNew}</span>`;
        
        let newCount = newVendors.length;
        let nextNewStr = "";
        
        if (newCount > 1) {
            let nextNewList = newVendors.slice(1, 6).map(v => {
                let vSpend = (vendorSpend[v][currYear] || 0) * forecastFactor;
                return `<span style="color: #6f42c1; font-weight: bold;">${v}</span><span style="font-weight: bold;"> (${formatBillionInt(vSpend)}억)</span>`;
            });
            nextNewStr = ` 그외 ${nextNewList.join(', ')} 등 ${newCount}개 Vendor와 신규거래가 발생했습니다.`;
        } else {
            nextNewStr = ` 올해 추가된 신규 Vendor는 ${topNewName} 1개사 입니다.`;
        }
        
        insights.push(`<b>[공급망 다변화 - 신규]</b> 올해 신규거래가 발생한 Vendor 중 ${topNewName}의 구매액이 <b>${formatBillionInt(topNewSpend)}억</b>으로 가장 높습니다.${nextNewStr}`);
      } else {
        insights.push(`<b>[공급망 구조 유지]</b> 올해 신규 Vendor는 없으며, 기존 공급망 체제가 유지되고 있습니다.`);
      }

      let droppedVendors = Object.keys(vendorSpend).filter(v => (vendorSpend[v][prevYear]||0) > 0 && (vendorSpend[v][currYear]||0) === 0);
      if (droppedVendors.length > 0) {
          droppedVendors.sort((a,b) => vendorSpend[b][prevYear] - vendorSpend[a][prevYear]);
          let droppedCount = droppedVendors.length;
          
          let droppedDetails = droppedVendors.slice(0, 5).map((v, idx) => {
              let pSpend = vendorSpend[v][prevYear] || 0;
              let prefix = idx === 0 ? "전년 구매액 " : "";
              return `<span style="color: #0D6EFD; font-weight: bold;">${v}</span> (<b style="color: #333;">${prefix}${formatBillionInt(pSpend)}억</b>)`;
          }).join(", ");
          
          let droppedEtc = droppedCount > 5 ? " 등" : "";
          
          insights.push(`<b>[공급망 다변화 - 축소]</b> 작년比 올해 거래이력이 없는 Vendor는 ${droppedDetails}${droppedEtc} ${droppedCount}개사 입니다. 거래종결이 아닌 경우, 거래가 발생하지 않은 사유 점검이 필요합니다.`);
      }
      
      const ipcNameMap = {
          "SIPC": "싱가폴IPC",
          "TWIPC": "대만IPC",
          "HKIPC": "중국IPC",
          "USIPC": "미주IPC",
          "UKIPC": "구주IPC"
      };
      
      let ipcPVendors = Object.keys(currIpcSpendByPVendor).sort((a, b) => currIpcSpendByPVendor[b] - currIpcSpendByPVendor[a]);
      if (ipcPVendors.length > 0) {
          let ipcPVendorCount = ipcPVendors.length;
          let ipcTotalSpendFormatted = formatBillionInt(currIpcTotalSpend * forecastFactor);
          
          let ipcPVendorDetails = ipcPVendors.map(pv => {
              let spend = currIpcSpendByPVendor[pv] * forecastFactor;
              let mappedPv = ipcNameMap[pv] || pv;
              return `<span style="color: #6f42c1; font-weight: bold;">${mappedPv}</span> <b style="color: #333;">(${formatBillionInt(spend)}억)</b>`;
          }).join(", ");

          let top2Ipcs = ipcPVendors.slice(0, 2);
          let extraIpcStr = "";

          const getTopVendorsStr = (ipcCode) => {
              if (!currIpcSpendDetails[ipcCode]) return "";
              let sortedVendors = Object.keys(currIpcSpendDetails[ipcCode])
                  .sort((a, b) => currIpcSpendDetails[ipcCode][b] - currIpcSpendDetails[ipcCode][a])
                  .slice(0, 3);
                  
              return sortedVendors.map(v => {
                  let vSpend = currIpcSpendDetails[ipcCode][v] * forecastFactor;
                  return `<span style="color: #6f42c1; font-weight: bold;">${v}</span> <b style="color: #333;">(${formatBillionInt(vSpend)}억)</b>`;
              }).join(", ");
          };

          if (top2Ipcs.length === 2) {
              let ipc1 = top2Ipcs[0];
              let ipc2 = top2Ipcs[1];
              let mappedIpc1 = `<span style="color: #6f42c1; font-weight: bold;">${ipcNameMap[ipc1] || ipc1}</span>`;
              let mappedIpc2 = `<span style="color: #6f42c1; font-weight: bold;">${ipcNameMap[ipc2] || ipc2}</span>`;
              
              extraIpcStr = ` ${mappedIpc1}는 ${getTopVendorsStr(ipc1)}의 금액비중이 높고, ${mappedIpc2}는 ${getTopVendorsStr(ipc2)} 순입니다.`;
          } else if (top2Ipcs.length === 1) {
              let ipc1 = top2Ipcs[0];
              let mappedIpc1 = `<span style="color: #6f42c1; font-weight: bold;">${ipcNameMap[ipc1] || ipc1}</span>`;
              
              extraIpcStr = ` ${mappedIpc1}는 ${getTopVendorsStr(ipc1)}의 금액비중이 높습니다.`;
          }
          insights.push(`<b>[IPC 거래현황]</b> 당사는 ${ipcPVendorCount}개의 IPC와 <b style="color: #333;">${ipcTotalSpendFormatted}억원</b> 거래중이고, 거래액 기준 ${ipcPVendorDetails} 입니다.${extraIpcStr}`);
          
          let ipcVendorsList = Object.keys(currIpcSpendByVendor).sort((a, b) => currIpcSpendByVendor[b] - currIpcSpendByVendor[a]);
          let ipcVendorCount = ipcVendorsList.length;
          let topIpcVendors = ipcVendorsList.slice(0, 5); 
          let vendorStrList = topIpcVendors.map(v => {
              let spend = currIpcSpendByVendor[v] * forecastFactor;
              return `<span style="color: #6f42c1; font-weight: bold;">${v}</span> (${formatBillionInt(spend)}억)`;
          }).join(", ");
          
          const l3SpendDetails = {}; 
          for (const vendor in currIpcVendorL3Spend) {
              for (const l3 in currIpcVendorL3Spend[vendor]) {
                  const spend = currIpcVendorL3Spend[vendor][l3];
                  if (!l3SpendDetails[l3]) {
                      l3SpendDetails[l3] = { totalSpend: 0, vendors: {} };
                  }
                  l3SpendDetails[l3].totalSpend += spend;
                  l3SpendDetails[l3].vendors[vendor] = spend;
              }
          }

          const top3L3s = Object.keys(l3SpendDetails)
              .sort((a, b) => l3SpendDetails[b].totalSpend - l3SpendDetails[a].totalSpend)
              .slice(0, 5);

          const newL3DetailsStr = top3L3s.map(l3 => {
              const l3Data = l3SpendDetails[l3];
              const totalL3SpendFormatted = formatBillionInt(l3Data.totalSpend * forecastFactor);
              const sortedVendors = Object.keys(l3Data.vendors)
                  .sort((a, b) => l3Data.vendors[b] - l3Data.vendors[a]);
              
              const vendorCount = sortedVendors.length;
              const top3Vendors = sortedVendors.slice(0, 3);
              
              const top3VendorsStyled = top3Vendors.map(v => `<span style="color: #6f42c1; font-weight: bold;">${v}</span>`);
              const vendorsStr = top3VendorsStyled.join(", ") + (vendorCount > 3 ? " 등" : "");
              return `<b style="color: #333;">${l3}</b> (<b style="color: #333;">${totalL3SpendFormatted}억</b>. ${vendorsStr} ${vendorCount}개사)`;
          }).join(", ");

          insights.push(`<b>[IPC 거래 Vendor]</b> IPC 거래 Vendor는 ${vendorStrList} 등 ${ipcVendorCount}개사 이며, 주요 거래품목은 ${newL3DetailsStr} 입니다.`);
      }

      let newIpcVendors = Object.keys(ipcVendorSpend).filter(v => 
          (ipcVendorSpend[v][currYear] || 0) > 0 && 
          (ipcVendorSpend[v][prevYear] || 0) === 0
      );
      
      if (newIpcVendors.length > 0) {
          newIpcVendors.sort((a,b) => ipcVendorSpend[b][currYear] - ipcVendorSpend[a][currYear]);
          
          let ipcCount = newIpcVendors.length;
          
          if (ipcCount === 1) {
              let v = newIpcVendors[0];
              let vSpend = (ipcVendorSpend[v][currYear] || 0) * forecastFactor;
              let ipcCodeRaw = vendorIpcMap[v] || '알수없음';
              let ipcCodeName = ipcNameMap[ipcCodeRaw] || ipcCodeRaw; 
              
              let ipcCode = `<span style="color: #6f42c1; font-weight: bold;">${ipcCodeName}</span>`;
              let vName = `<span style="color: #6f42c1; font-weight: bold;">${v}</span>`;
              
              let l3SpendObj = currIpcVendorL3Spend[v] || {};
              let sortedL3s = Object.keys(l3SpendObj).sort((a, b) => l3SpendObj[b] - l3SpendObj[a]);
              let topL3s = sortedL3s.slice(0, 3);
              let suffix = sortedL3s.length > 3 ? " 등" : "";
              let l3Str = topL3s.length > 0 ? ` (${topL3s.join(", ")}${suffix})` : "";
              
              insights.push(`<b>[IPC 거래선 증가]</b> 올해 신규거래가 발생한 IPC 거래 Vendor는 ${vName} 1개사이며, ${ipcCode}와 <b style="color: #333;">${formatBillionInt(vSpend)}억원</b>${l3Str} 거래 중입니다.`);
          } else {
              let displayCount = Math.min(ipcCount, 5);
              let ipcDetails = newIpcVendors.slice(0, displayCount).map(v => {
                  let vSpend = (ipcVendorSpend[v][currYear] || 0) * forecastFactor;
                  let ipcCodeRaw = vendorIpcMap[v] || '알수없음';
                  let ipcCodeName = ipcNameMap[ipcCodeRaw] || ipcCodeRaw; 
                  
                  let ipcCode = `<span style="color: #6f42c1; font-weight: bold;">${ipcCodeName}</span>`;
                  let vName = `<span style="color: #6f42c1; font-weight: bold;">${v}</span>`;
                  
                  let l3SpendObj = currIpcVendorL3Spend[v] || {};
                  let sortedL3s = Object.keys(l3SpendObj).sort((a, b) => l3SpendObj[b] - l3SpendObj[a]);
                  let topL3s = sortedL3s.slice(0, 3);
                  let suffix = sortedL3s.length > 3 ? " 등" : "";
                  let l3Str = topL3s.length > 0 ? ` ${topL3s.join(", ")}${suffix}` : "";
                  
                  return `${vName} <b style="color: #333;">(${formatBillionInt(vSpend)}억)</b>, ${ipcCode},${l3Str}`;
              }).join(', ');
              
              let etcStr = " 등"; 
              
              insights.push(`<b>[IPC 거래선 증가]</b> 전년도 거래가 없었다가 올해 신규거래가 발생한 Vendor 중 IPC 거래선은 ${ipcCount}개 Vendor 이며, 금액순서대로 ${ipcDetails}${etcStr} 입니다.`);
          }
      }
    }

    let agencySpend = {}; 
    let proxySpend = {};  
    
    Object.keys(finalData.data).forEach(key => {
      const parts = key.split("|");
      const year = parts[0];
      const tradeType = parts[2];
      const maker = parts[9];
      const krw = finalData.data[key].krw;
      
      if (year === currYear) {
        if (tradeType.includes("대리점")) {
          if (!agencySpend[maker]) agencySpend[maker] = 0;
          agencySpend[maker] += krw;
        } else if (tradeType.includes("대행사")) {
          if (!proxySpend[maker]) proxySpend[maker] = 0;
          proxySpend[maker] += krw;
        }
      }
    });

    let agencyHigh = []; 
    let proxyHigh = [];  
    Object.keys(agencySpend).forEach(m => {
      let spend = agencySpend[m] * forecastFactor;
      if (spend >= 300000000) agencyHigh.push({name: m, spend: spend});
    });
    Object.keys(proxySpend).forEach(m => {
      let spend = proxySpend[m] * forecastFactor;
      if (spend >= 300000000) proxyHigh.push({name: m, spend: spend});
    });
    
    agencyHigh.sort((a,b) => b.spend - a.spend);
    proxyHigh.sort((a,b) => b.spend - a.spend);
    
    if (agencyHigh.length > 0) {
      let topStr = agencyHigh.slice(0, 5).map(m => `<span style="color: #198754; font-weight: bold;">${m.name}</span> <b style="color: #333;">(${formatBillionInt(m.spend)}억)</b>`).join(", ");
      insights.push(`<b>[대리점→직거래 전환 검토]</b> 대리점 거래 Maker 중 연간 거래액이 3억 이상인 Maker는 ${topStr} 등 ${agencyHigh.length}개사 입니다. "협력업체 관리기준" 상 연간 거래액이 3억 이상인 Maker는 직거래 전환하는 것이 기본 원칙이며, 직거래할 수 없는 사유를 매년 점검해야합니다.`);
    }
    
    if (proxyHigh.length > 0) {
      let topStr = proxyHigh.slice(0, 5).map(m => `<span style="color: #198754; font-weight: bold;">${m.name}</span> <b style="color: #333;">(${formatBillionInt(m.spend)}억)</b>`).join(", ");
      insights.push(`<b>[대행사→직거래 전환 검토]</b> 대행사 거래 Maker 중 연간 거래액이 3억 이상인 Maker는 ${topStr} 등 ${proxyHigh.length}개사 입니다. 대행사 거래 Maker도 직거래 전환이 기본 원칙이며, 구매대행 수수료가 약 0.3% 발생하는 것을 염두에 두어야합니다.`);
    }

    let microMakers = [];
    Object.keys(makerSpend).forEach(m => {
      let spend = (makerSpend[m][currYear] || 0) * forecastFactor;
      if (spend > 0 && spend < 10000000) { 
        microMakers.push({name: m, spend: spend});
      }
    });
    
    microMakers.sort((a,b) => b.spend - a.spend); 
    if (microMakers.length > 0) {
      let topStr = microMakers.slice(0, 5).map(m => `<span style="color: #198754; font-weight: bold;">${m.name}</span> <b style="color: #333;">(${Math.round(m.spend / 10000).toLocaleString()}만)</b>`).join(", ");
      insights.push(`<b>[소액 거래선 점검]</b> 연간 거래액 천만원 미만의 소액거래 Maker는 ${topStr} 등 ${microMakers.length}개사이며, Maker 운영전략을 확인하여 거래선 통합/정리 등 필요여부를 점검해주시길 바랍니다.`);
    }

    let currForeign = foreignSpend.curr * forecastFactor;
    let prevForeign = foreignSpend.prev;
    
    let currForeignPct = currTotal > 0 ? (currForeign / currTotal * 100).toFixed(1) : "0.0";
    let prevForeignPct = prevTotal > 0 ? (prevForeign / prevTotal * 100).toFixed(1) : "0.0";
    
    if (prevYear && prevTotal > 0) {
      let diffForeign = currForeign - prevForeign;
      let isForeignInc = diffForeign >= 0;
      let dirForeign = isForeignInc ? "증가" : "감소";
      let colorForeign = isForeignInc ? "#DC3545" : "#0D6EFD";
      let yoyForeign = prevForeign > 0 ? (Math.abs(diffForeign) / prevForeign * 100).toFixed(1) : "0.0";
      
      insights.push(`<b>[외화 거래품목 비중]</b> 외화 거래 품목의 올해 구매액은 ${formatBillionInt(currForeign)}억(${currForeignPct}%) 이며, 전년도 ${formatBillionInt(prevForeign)}억(${prevForeignPct}%) 대비 <span style="color: ${colorForeign}; font-weight: bold;">${yoyForeign}% ${dirForeign}</span> 했습니다. 직접구매 자재는 환차손/익에 영향을 받지 않지만, 전체 구매액은 해당월 환율 기준으로 환산된 원화금액으로 표시되므로, 최근의 구매액 ${overallDirection}에 환율의 영향이 있을 수 있습니다.`);
    } else {
      insights.push(`<b>[외화 거래품목 비중]</b> 외화 거래 품목의 올해 구매액은 ${formatBillionInt(currForeign)}억(${currForeignPct}%) 입니다. 직접구매 자재는 환차손/익에 영향을 받지 않지만, 전체 구매액은 해당월 환율 기준으로 환산된 원화금액으로 표시되므로, 최근의 구매액 추이에 환율의 영향이 있을 수 있습니다.`);
    }

    let currKRW = (currencySpend.curr['KRW'] || 0) * forecastFactor;
    let currUSD = (currencySpend.curr['USD'] || 0) * forecastFactor;
    let currEUR = (currencySpend.curr['EUR'] || 0) * forecastFactor;
    let currJPY = (currencySpend.curr['JPY'] || 0) * forecastFactor;

    let prevKRW = currencySpend.prev['KRW'] || 0;
    let prevUSD = currencySpend.prev['USD'] || 0;

    let pctCurrKRW = currTotal > 0 ? (currKRW / currTotal * 100).toFixed(1) : "0.0";
    let pctCurrUSD = currTotal > 0 ? (currUSD / currTotal * 100).toFixed(1) : "0.0";
    let pctCurrEUR = currTotal > 0 ? (currEUR / currTotal * 100).toFixed(1) : "0.0";
    let pctCurrJPY = currTotal > 0 ? (currJPY / currTotal * 100).toFixed(1) : "0.0";

    let pctPrevKRW = prevTotal > 0 ? (prevKRW / prevTotal * 100).toFixed(1) : "0.0";
    let pctPrevUSD = prevTotal > 0 ? (prevUSD / prevTotal * 100).toFixed(1) : "0.0";

    let diffKRW = parseFloat(pctCurrKRW) - parseFloat(pctPrevKRW);
    let diffUSD = parseFloat(pctCurrUSD) - parseFloat(pctPrevUSD);

    let diffKRWStr = diffKRW >= 0 ? `<span style="color: #DC3545; font-weight: bold;">${diffKRW.toFixed(1)}%p 증가</span>` : `<span style="color: #0D6EFD; font-weight: bold;">${Math.abs(diffKRW).toFixed(1)}%p 감소</span>`;
    let diffUSDStr = diffUSD >= 0 ? `<span style="color: #DC3545; font-weight: bold;">${diffUSD.toFixed(1)}%p 증가</span>` : `<span style="color: #0D6EFD; font-weight: bold;">${Math.abs(diffUSD).toFixed(1)}%p 감소</span>`;

    if (prevYear && prevTotal > 0) {
        insights.push(`<b>[환종별 거래 비중]</b> 올해 환종별 구매액은 KRW <b style="color: #333;">${formatBillionInt(currKRW)}억</b> (${pctCurrKRW}%), USD <b style="color: #333;">${formatBillionInt(currUSD)}억</b> (${pctCurrUSD}%), EUR <b style="color: #333;">${formatBillionInt(currEUR)}억</b> (${pctCurrEUR}%), JPY <b style="color: #333;">${formatBillionInt(currJPY)}억</b> (${pctCurrJPY}%) 입니다. 전년비 KRW 비중은 ${diffKRWStr}, USD 비중은 ${diffUSDStr} 했습니다.`);
    } else {
        insights.push(`<b>[환종별 거래 비중]</b> 올해 환종별 구매액은 KRW <b style="color: #333;">${formatBillionInt(currKRW)}억</b> (${pctCurrKRW}%), USD <b style="color: #333;">${formatBillionInt(currUSD)}억</b> (${pctCurrUSD}%), EUR <b style="color: #333;">${formatBillionInt(currEUR)}억</b> (${pctCurrEUR}%), JPY <b style="color: #333;">${formatBillionInt(currJPY)}억</b> (${pctCurrJPY}%) 입니다.`);
    }

    ['KRW', 'USD', 'EUR', 'JPY'].forEach(curr => {
         let curVendorObj = vendorByCurrency[curr] || {};
         let curL2Obj = l2ByCurrency[curr] || {};

         let topVendors = Object.keys(curVendorObj).map(v => ({name: v, spend: curVendorObj[v] * forecastFactor})).filter(x => x.spend > 0).sort((a,b) => b.spend - a.spend).slice(0, 5);
         let topL2s = Object.keys(curL2Obj).map(l => ({name: l, spend: curL2Obj[l] * forecastFactor})).filter(x => x.spend > 0).sort((a,b) => b.spend - a.spend).slice(0, 5);

         if (topVendors.length > 0) {
             let vStr = topVendors.map(x => `<span style="color: #6f42c1; font-weight: bold;">${x.name}</span> (<b style="color: #333;">${formatBillionInt(x.spend)}억</b>)`).join(", ");
             let lStr = topL2s.map(x => `<span style="color: #198754; font-weight: bold;">${x.name}</span> (<b style="color: #333;">${formatBillionInt(x.spend)}억</b>)`).join(", ");
             
             let vendorSuffix = topVendors.length >= 5 ? " 등" : "";
             let l2Suffix = topL2s.length >= 5 ? " 등" : "";

             insights.push(`<b>[${curr}]</b> Vendor 기준 ${vStr}${vendorSuffix} 순서입니다. 품목 기준으로는 ${lStr}${l2Suffix} 순입니다.`);
         }
    });

    const targetL1s = ["기구", "전기/전자", "전장", "원부자재", "포장/인쇄", "프로브"];
    let l1Data = {};
    Object.keys(finalData.data).forEach(key => {
      const parts = key.split("|");
      const year = parts[0];
      const l1 = parts[4];     
      const vendor = parts[7]; 
      const krw = finalData.data[key].krw;
      if (targetL1s.includes(l1)) {
        if (!l1Data[l1]) l1Data[l1] = { total: {}, vendors: {} };
        if (!l1Data[l1].total[year]) l1Data[l1].total[year] = 0;
        l1Data[l1].total[year] += krw;
        if (!l1Data[l1].vendors[vendor]) l1Data[l1].vendors[vendor] = {};
        if (!l1Data[l1].vendors[vendor][year]) l1Data[l1].vendors[vendor][year] = 0;
        l1Data[l1].vendors[vendor][year] += krw;
      }
    });

    if (prevYear && prevTotal > 0) {
      targetL1s.forEach(l1 => {
        if (!l1Data[l1]) return;
        
        let cTotal = (l1Data[l1].total[currYear] || 0) * forecastFactor;
        let pTotal = l1Data[l1].total[prevYear] || 0;
        
        if (cTotal === 0 && pTotal === 0) return;
        
        let diff = cTotal - pTotal;
        let dir = diff >= 0 ? "증가" : "감소";
        let sign = diff >= 0 ? "+" : "-";
        let pct = pTotal > 0 ? ((Math.abs(diff) / pTotal) * 100).toFixed(1) : "N/A";
        let color = diff >= 0 ? "#DC3545" : "#0D6EFD";
        
        let allCurrSorted = Object.keys(l1Data[l1].vendors)
            .map(v => ({ name: v, spend: (l1Data[l1].vendors[v][currYear] || 0) * forecastFactor }))
            .sort((a,b) => b.spend - a.spend);
            
        let allPrevSorted = Object.keys(l1Data[l1].vendors)
            .map(v => ({ name: v, spend: l1Data[l1].vendors[v][prevYear] || 0 }))
            .sort((a,b) => b.spend - a.spend);

        let currTop5 = allCurrSorted.filter(v => v.spend > 0).slice(0, 5);
        let prevTop5 = allPrevSorted.filter(v => v.spend > 0).slice(0, 5);
            
        let cStr = currTop5.map(v => `<b>${v.name}</b> (${formatBillionInt(v.spend)}억)`).join(", ");
        let pStr = prevTop5.map(v => `<b>${v.name}</b> (${formatBillionInt(v.spend)}억)`).join(", ");
        
        let unionVendors = [...new Set([...currTop5.map(v=>v.name), ...prevTop5.map(v=>v.name)])];
        let rankChanges = [];
        unionVendors.forEach(vName => {
            let cIdx = allCurrSorted.findIndex(v => v.name === vName && v.spend > 0);
            let pIdx = allPrevSorted.findIndex(v => v.name === vName && v.spend > 0);
            
            let cRankStr = cIdx >= 0 ? (cIdx + 1) : "-";
            let pRankStr = pIdx >= 0 ? (pIdx + 1) : "-";
            
            if (cRankStr !== pRankStr) {
                let colorCode = "#333";
                let rankLabel = "";
                
                if (pRankStr === "-") {
                    colorCode = "#DC3545";
                    rankLabel = "신규진입";
                } else if (cRankStr === "-") {
                    colorCode = "#0D6EFD";
                    rankLabel = `${pRankStr}위→제외`;
                } else {
                    if (cIdx < pIdx) colorCode = "#DC3545";
                    else if (cIdx > pIdx) colorCode = "#0D6EFD";
                    rankLabel = `${pRankStr}→${cRankStr}위`;
                }
                rankChanges.push(`<span style="color: ${colorCode};"><b>${vName}</b> (${rankLabel})</span>`);
            }
        });

        let vendorChangeStr = rankChanges.length > 0 
            ? `. 변동내역은 ${rankChanges.join(", ")}.` 
            : ` , 상위사의 순위 변동은 없습니다.`;
        
        let insightText = `<b>[품목 동향 - ${l1}]</b> 구매액 <b>${formatBillionInt(cTotal)}억</b>, 전년비 <span style="color: ${color}; font-weight: bold;">${formatBillionInt(Math.abs(diff))}억 (${sign}${pct}%) ${dir}</span>. ${cStr} 순서로, 전년도 상위권은 ${pStr}${vendorChangeStr}`;
        
        insights.push(insightText);
      });
    }

    const targetDetails = ["기구Assy", "판금", "사출", "능동소자", "수동소자","PC Mother(L3)", "VGA(L3)", "PCB(L3)", "FPCB(L3)", "PBA", "Cable", "Cable Assy", "파워Assy", "모니터Assy", "Single Crystal(L3)", "Ceramic(L3)", "포장", "Accessory"]; 
    let detailData = {};
    Object.keys(finalData.data).forEach(key => {
      const parts = key.split("|");
      const year = parts[0];
      const l2 = parts[5];     
      const l3 = parts[6];     
      const vendor = parts[7]; 
      const krw = finalData.data[key].krw;

      targetDetails.forEach(target => {
        let isMatch = false;
        if (target.endsWith("(L3)")) {
            let actualL3Name = target.replace("(L3)", "").trim(); 
            if (l3 === actualL3Name) isMatch = true;
        } else {
            if (l2 === target) isMatch = true;
        }
        
        if (isMatch) {
            if (!detailData[target]) detailData[target] = { total: {}, vendors: {} };
            if (!detailData[target].total[year]) detailData[target].total[year] = 0;
            detailData[target].total[year] += krw;
            if (!detailData[target].vendors[vendor]) detailData[target].vendors[vendor] = {};
            if (!detailData[target].vendors[vendor][year]) detailData[target].vendors[vendor][year] = 0;
            detailData[target].vendors[vendor][year] += krw;
        }
      });
    });

    if (prevYear && prevTotal > 0) {
      targetDetails.forEach(target => {
        if (!detailData[target]) return;
        
        let cTotal = (detailData[target].total[currYear] || 0) * forecastFactor;
        let pTotal = detailData[target].total[prevYear] || 0;
        
        if (cTotal === 0 && pTotal === 0) return;
        
        let diff = cTotal - pTotal;
        let dir = diff >= 0 ? "증가" : "감소";
        let sign = diff >= 0 ? "+" : "-";
        let pct = pTotal > 0 ? ((Math.abs(diff) / pTotal) * 100).toFixed(1) : "N/A";
        let color = diff >= 0 ? "#DC3545" : "#0D6EFD";
        
        let allCurrSorted = Object.keys(detailData[target].vendors)
            .map(v => ({ name: v, spend: (detailData[target].vendors[v][currYear] || 0) * forecastFactor }))
            .sort((a,b) => b.spend - a.spend);
            
        let allPrevSorted = Object.keys(detailData[target].vendors)
            .map(v => ({ name: v, spend: detailData[target].vendors[v][prevYear] || 0 }))
            .sort((a,b) => b.spend - a.spend);

        let currTop5 = allCurrSorted.filter(v => v.spend > 0).slice(0, 5);
        let prevTop5 = allPrevSorted.filter(v => v.spend > 0).slice(0, 5);
            
        let cStr = currTop5.map(v => `<b>${v.name}</b> (${formatBillionInt(v.spend)}억)`).join(", ");
        let pStr = prevTop5.map(v => `<b>${v.name}</b> (${formatBillionInt(v.spend)}억)`).join(", ");
        
        let unionVendors = [...new Set([...currTop5.map(v=>v.name), ...prevTop5.map(v=>v.name)])];
        let rankChanges = [];
        unionVendors.forEach(vName => {
            let cIdx = allCurrSorted.findIndex(v => v.name === vName && v.spend > 0);
            let pIdx = allPrevSorted.findIndex(v => v.name === vName && v.spend > 0);
            
            let cRankStr = cIdx >= 0 ? (cIdx + 1) : "-";
            let pRankStr = pIdx >= 0 ? (pIdx + 1) : "-";
            
            if (cRankStr !== pRankStr) {
                let colorCode = "#333";
                let rankLabel = "";
                
                if (pRankStr === "-") {
                    colorCode = "#DC3545";
                    rankLabel = "신규진입";
                } else if (cRankStr === "-") {
                    colorCode = "#0D6EFD";
                    rankLabel = `${pRankStr}위→제외`;
                } else {
                    if (cIdx < pIdx) colorCode = "#DC3545";
                    else if (cIdx > pIdx) colorCode = "#0D6EFD";
                    rankLabel = `${pRankStr}→${cRankStr}위`;
                }
                rankChanges.push(`<span style="color: ${colorCode};"><b>${vName}</b> (${rankLabel})</span>`);
            }
        });

        let vendorChangeStr = rankChanges.length > 0 
            ? `. 변동내역은 ${rankChanges.join(", ")}.` 
            : ` 으로, 상위사의 순위 변동은 없습니다.`;
        
        let displayTarget = target.replace("(L3)", "").trim();
        let insightText = `<b>[세부품목 동향 - ${displayTarget}]</b> 구매액 <b>${formatBillionInt(cTotal)}억</b>, 전년비 <span style="color: ${color}; font-weight: bold;">${formatBillionInt(Math.abs(diff))}억 (${sign}${pct}%) ${dir}</span>. ${cStr} 순서로, 전년도 상위권은 ${pStr}${vendorChangeStr}`;
        
        insights.push(insightText);
      });
    }

    let finalInsights = insights.slice(0, 50);
    while (finalInsights.length < 10) {
      finalInsights.push(`<b>[추가 모니터링 요망]</b> 품목군(L1/L2/L3) 필터를 통해 상세 카테고별 재고 및 구매액 추이를 교차 검증해주시기 바랍니다.`);
    }

    if (finalInsights.length > 0) {
      for (let i = 0; i < FILE_IDS.length; i++) {
        const ss = SpreadsheetApp.openById(FILE_IDS[i].trim());
        const updateSheet = ss.getSheetByName("Setting");
        if (updateSheet) {
          updateSheet.getRange("C3:C").clearContent();
          
          const writeData = finalInsights.map(ins => [ins]);
          updateSheet.getRange(3, 3, writeData.length, 1).setValues(writeData);
          
          finalData.aiInsights = finalInsights;
          break;
        }
      }
      return true;
    }
  } catch(e) {
    console.error("내부 트렌드 분석 및 시사점 생성 중 에러: ", e);
  }
  return false;
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
    try { userEmail = Session.getActiveUser().getEmail(); } catch(e) {}
    const userId = userEmail ? userEmail.split('@')[0] : "Unknown";
    
    logSheet.appendRow([new Date(), userId, moduleTitle]);
  } catch (e) {
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
        if (currentGroup.length > 0) {
          menuGroups.push(currentGroup);
        }
        currentGroup = [];
        currentGroupNum = groupNum;
      }
      currentGroup.push({ title: title, url: url, icon: icon });
    }
    
    if (currentGroup.length > 0) {
      menuGroups.push(currentGroup);
    }
    
    return menuGroups;
  } catch (e) {
    console.error("메뉴 데이터 가져오기 실패: " + e.message);
    return [];
  }
}