function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
      .setTitle('원자재 재고실적 분석')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getActiveUserEmail() {
  return Session.getActiveUser().getEmail();
}

function get6MonthsLabels(baseDateRaw) {
  let year = 2026;
  let month = 6; // July is index 6
  if (baseDateRaw && baseDateRaw.length === 8) {
    year = parseInt(baseDateRaw.substring(0, 4), 10);
    month = parseInt(baseDateRaw.substring(4, 6), 10) - 1; // 0-indexed
  }
  let labels = [];
  for (let i = 5; i >= 0; i--) {
    let d = new Date(year, month - i, 1);
    let y = d.getFullYear();
    let m = d.getMonth() + 1;
    let mStr = m < 10 ? "0" + m : "" + m;
    labels.push(`${y}-${mStr}`);
  }
  return labels;
}

function getRawAnalysisData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName('Info Record');
    if (!infoSheet) {
      throw new Error("'Info Record' 시트를 찾을 수 없습니다.");
    }
    const infoLastRow = infoSheet.getLastRow();
    const infoData = infoSheet.getRange("A1:N" + infoLastRow).getValues();
    const infoMap = {};
    for (let i = 1; i < infoData.length; i++) {
      const row = infoData[i];
      const key = String(row[0]).trim(); // A열: 구분자(Key)
      if (!key) continue;
      let partVal = String(row[3]).trim(); // D열: 파트
      infoMap[key] = {
        part: partVal ? (partVal.includes("파트") ? partVal : partVal + "파트") : "미분류",
        purchaser: String(row[4]).trim(),   // E열: Purchaser
        l1: String(row[5]).trim(),          // F열: L1
        l2: String(row[6]).trim(),          // G열: L2
        l3: String(row[7]).trim(),          // H열: L3
        stockType: String(row[8]).trim(),   // I열: 일반, LTB, 전략비축
        stockReason: String(row[9]).trim(), // J열: 전략비축 세부
        moq: String(row[10]).trim(),        // K열: MOQ
        lt: String(row[11]).trim(),         // L열: L/T
        vendor: String(row[12]).trim(),     // M열: Vendor Code
        vendorDesc: String(row[13]).trim()  // N열: Vendor Name
      };
    }

    const processRawSheetToMap = (sheetName) => {
      const rawSheet = ss.getSheetByName(sheetName);
      if (!rawSheet) return null;
      const dateRange = rawSheet.getRange("B2");
      const dateRaw = dateRange ? String(dateRange.getDisplayValue()).trim() : "";
      let formattedDate = dateRaw.length === 8
          ? `${dateRaw.substring(0,4)}-${dateRaw.substring(4,6)}-${dateRaw.substring(6,8)}`
          : dateRaw;
      const reportTitle = `원자재 재고실적 (${formattedDate})`;
      
      const lastRow = rawSheet.getLastRow();
      if (lastRow < 2) return null;
      
      const rawData = rawSheet.getRange("A1:CK" + lastRow).getValues();
      const dataMap = {};
      const parseNumber = value => {
        if (typeof value === 'number') return value;
        const parsed = parseFloat(value);
        return isNaN(parsed) ? 0 : parsed;
      };
      for (let i = 1; i < rawData.length; i++) {
        const row = rawData[i];
        const key = String(row[0]).trim();
        if (!key) continue;
        const info = infoMap[key] || {};
        const purchaser = info.purchaser || "미지정";
        if (purchaser.toUpperCase() === "DR") continue;

        // ==========================================
        // 1. 수량(Q) 
        // ==========================================
        const inTransitQ = parseNumber(row[24]); // Y열 (수량 In-Transit) -> Index 24
        const onHandQ = parseNumber(row[28]);    // AC열 (수량 On-Hand) -> Index 28
        const totalQ = inTransitQ + onHandQ;
        const openPoQ = parseNumber(row[26]);    // AA열 (수량 Open PO) -> Index 26

        // 수량 Aging 데이터 수집 (AK열 ~ AQ열 -> Index 36 ~ 42)
        const agingQ = []; 
        let deadQ = 0;
        for (let c = 36; c <= 42; c++) {
          const valQ = parseNumber(row[c]);
          agingQ.push(valQ);
          if (c >= 39 && c <= 42) {
              deadQ += valQ; 
          }
        }

        let mrp8Q = 0;
        let mrp4Q = 0;
        for (let c = 55; c <= 62; c++) { mrp8Q += parseNumber(row[c]); }
        for (let c = 55; c <= 58; c++) { mrp4Q += parseNumber(row[c]); }
        const invDaysQ = (mrp8Q > 0) ? (totalQ / mrp8Q) * 60 : 0;

        // ==========================================
        // 2. 금액(A)
        // ==========================================
        const inTransitA = parseNumber(row[25]); // Z열 (금액 In-Transit) -> Index 25
        const onHandA = parseNumber(row[29]);    // AD열 (금액 On-Hand) -> Index 29
        const totalA = inTransitA + onHandA;
        const openPoA = parseNumber(row[27]);    // AB열 (금액 Open PO) -> Index 27

        // 금액 Aging 데이터 수집 (AV열 ~ BB열 -> Index 47 ~ 53)
        const agingA = [];
        let deadA = 0;
        for (let c = 47; c <= 53; c++) {
          const valA = parseNumber(row[c]);
          agingA.push(valA);
          if (c >= 50 && c <= 53) {
              deadA += valA; 
          }
        }

        let mrp8A = 0;
        let mrp4A = 0;
        for (let c = 71; c <= 78; c++) { mrp8A += parseNumber(row[c]); }
        for (let c = 71; c <= 74; c++) { mrp4A += parseNumber(row[c]); }
        const invDaysA = (mrp8A > 0) ? (totalA / mrp8A) * 60 : 0;

        dataMap[key] = {
          plant: String(row[2]).trim(),
          material: String(row[3]).trim(),
          materialDesc: String(row[4]).trim(),
          purchaser: purchaser,
          part: info.part || "미분류",
          l1: info.l1 || "미분류", l2: info.l2 || "미분류", l3: info.l3 || "미분류",
          stockType: info.stockType || "일반",
          stockReason: info.stockReason || "",
          vendor: info.vendor || "", vendorDesc: info.vendorDesc || "미지정",
          moq: info.moq || "", lt: info.lt || "",
          totalA, deadA, mrp8A, mrp4A, invDaysA, agingA, inTransitA, onHandA, openPoA,
          totalQ, deadQ, mrp8Q, mrp4Q, invDaysQ, agingQ, inTransitQ, onHandQ, openPoQ
        };
      }
      return { reportTitle, dataMap, dateRaw };
    };

    const processHistorySheet = () => {
      const rawSheet = ss.getSheetByName('2-5개월전Raw');
      if (!rawSheet) return {};
      
      const lastRow = rawSheet.getLastRow();
      if (lastRow < 2) return {};
      
      const rawData = rawSheet.getRange("A2:CK" + lastRow).getValues();
      const historyMap = {};
      
      const parseNumber = value => {
        const parsed = parseFloat(value);
        return isNaN(parsed) ? 0 : parsed;
      };
      
      for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        const key = String(row[0]).trim();
        if (!key) continue;
        
        const dateRaw = String(row[1]).trim();
        if (dateRaw.length < 6) continue;
        const ym = `${dateRaw.substring(0,4)}-${dateRaw.substring(4,6)}`;
        
        // 금액(A)
        const inTransitA = parseNumber(row[25]);
        const onHandA = parseNumber(row[29]);
        const totalA = inTransitA + onHandA;
        let deadA = 0;
        for (let c = 50; c <= 53; c++) {
          deadA += parseNumber(row[c]);
        }
        let mrp8A = 0;
        for (let c = 71; c <= 78; c++) {
          mrp8A += parseNumber(row[c]);
        }
        const malA = parseNumber(row[53]); // Aging 181일 이상
        
        // 수량(Q)
        const inTransitQ = parseNumber(row[24]);
        const onHandQ = parseNumber(row[28]);
        const totalQ = inTransitQ + onHandQ;
        let deadQ = 0;
        for (let c = 39; c <= 42; c++) {
          deadQ += parseNumber(row[c]);
        }
        let mrp8Q = 0;
        for (let c = 55; c <= 62; c++) {
          mrp8Q += parseNumber(row[c]);
        }
        const malQ = parseNumber(row[42]); // Aging 181일 이상
        
        if (!historyMap[key]) {
          historyMap[key] = {};
        }
        
        historyMap[key][ym] = {
          totalA, deadA, mrp8A, malA,
          totalQ, deadQ, mrp8Q, malQ
        };
      }
      return historyMap;
    };

    const currentResult = processRawSheetToMap('재고실적Raw');
    const prevResult = processRawSheetToMap('전월재고Raw');
    if (!currentResult) throw new Error("'재고실적Raw' 시트를 찾을 수 없습니다.");

    let summaryData = null;
    let b8Title = "";
    const mainSheet = ss.getSheetByName('Main');
    if (mainSheet) {
      const summaryRange = mainSheet.getRange("B2:B5").getDisplayValues();
      const trendRange = mainSheet.getRange("G2:L4").getDisplayValues();
      summaryData = {
        tot: summaryRange[0][0],   // B2: 총재고
        dead: summaryRange[1][0],  // B3: 부진재고
        rate: summaryRange[2][0],  // B4: 부진율
        days: summaryRange[3][0],  // B5: 재고일수
        trendLabels: trendRange[0], // G2:L2 (월)
        trendTot: trendRange[1],    // G3:L3 (총재고 금액)
        trendDead: trendRange[2]    // G4:L4 (부진재고 금액)
      };
      
      const b8Value = String(mainSheet.getRange("B8").getDisplayValue()).trim();
      if (b8Value) {
        const nums = b8Value.match(/\d+/g);
        if (nums && nums.length >= 2) {
          const yyyy = nums[0];
          const m = parseInt(nums[1], 10);
          b8Title = `${yyyy}년 ${m}월 원자재 재고 분석 리포트`;
        } else {
          b8Title = b8Value.includes("리포트") ? b8Value : `${b8Value} 원자재 재고 분석 리포트`;
        }
      }
    }

    const mergedData = [];
    const currentKeys = Object.keys(currentResult.dataMap);
    const prevKeys = prevResult ? Object.keys(prevResult.dataMap) : [];
    const allKeys = new Set([...currentKeys, ...prevKeys]);

    allKeys.forEach(key => {
      const curr = currentResult.dataMap[key];
      const prev = prevResult ? prevResult.dataMap[key] : null;
      const base = curr || prev;
      
      const currentTotalQty = curr ? curr.totalQ : 0;
      const currentDeadQty = curr ? curr.deadQ : 0; 
      
      mergedData.push({
        ...base,
        key: key,
        isDeadStock: (currentTotalQty > 0) && (currentDeadQty > 0), 
        
        totalA: curr ? curr.totalA : 0,
        deadA: curr ? curr.deadA : 0,
        mrp8A: curr ? curr.mrp8A : 0,
        totalQ: curr ? curr.totalQ : 0,
        currAgingA: curr ? curr.agingA : [0,0,0,0,0,0,0],
        currAgingQ: curr ? curr.agingQ : [0,0,0,0,0,0,0],
        
        prevTotalA: prev ? prev.totalA : 0,
        prevDeadA: prev ? prev.deadA : 0,
        prevMrp8A: prev ? prev.mrp8A : 0,
        prevTotalQ: prev ? prev.totalQ : 0,
        prevDeadQ: prev ? prev.deadQ : 0,
        prevMrp8Q: prev ? prev.mrp8Q : 0,
        prevAgingA: prev ? prev.agingA : [0,0,0,0,0,0,0],
        prevAgingQ: prev ? prev.agingQ : [0,0,0,0,0,0,0],
        
        diffTotalA: (curr ? curr.totalA : 0) - (prev ? prev.totalA : 0),
        diffDeadA: (curr ? curr.deadA : 0) - (prev ? prev.deadA : 0),
        
        history: {}
      });
    });

    const filters = {
      plant: new Set(), part: new Set(), l1: new Set(),
      l2: new Set(), l3: new Set(), purchaser: new Set(), vendorDesc: new Set()
    };

    mergedData.forEach(d => {
        if (d.plant) filters.plant.add(d.plant);
        if (d.part) filters.part.add(d.part);
        if (d.l1) filters.l1.add(d.l1);
        if (d.l2) filters.l2.add(d.l2);
        if (d.l3) filters.l3.add(d.l3);
        if (d.purchaser) filters.purchaser.add(d.purchaser);
        if (d.vendorDesc) filters.vendorDesc.add(d.vendorDesc);
    });

    const monthLabels = get6MonthsLabels(currentResult.dateRaw);

    return {
      success: true,
      userEmail: Session.getActiveUser().getEmail(),
      reportTitle: b8Title || currentResult.reportTitle,
      agingLabels: ['30일 이하', '60일 이하', '90일 이하', '120일 이하', '150일 이하', '180일 이하', '180일 초과'],
      monthLabels: monthLabels, 
      data: mergedData, 
      summaryData: summaryData, 
      dateRaw: currentResult.dateRaw, 
      filters: {
         plant: [...filters.plant].sort(), part: [...filters.part].sort(),
         l1: [...filters.l1].sort(), l2: [...filters.l2].sort(), l3: [...filters.l3].sort(),
         purchaser: [...filters.purchaser].sort(), vendorDesc: [...filters.vendorDesc].sort()
      }
    };
  } catch (e) {
    Logger.log(e.stack);
    return { success: false, message: `스크립트 실행 실패: ${e.message}` };
  }
}

function getRawAnalysisDataPhase2(curDateRaw) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const rawSheet = ss.getSheetByName('2-5개월전Raw');
    const historyData = {};
    if (rawSheet) {
      const lastRow = rawSheet.getLastRow();
      if (lastRow >= 2) {
        const rawData = rawSheet.getRange("A2:CK" + lastRow).getValues();
        const parseNumber = value => {
          if (typeof value === 'number') return value;
          const parsed = parseFloat(value);
          return isNaN(parsed) ? 0 : parsed;
        };
        for (let i = 0; i < rawData.length; i++) {
          const row = rawData[i];
          const key = String(row[0]).trim();
          if (!key) continue;
          const dateRaw = String(row[1]).trim();
          if (dateRaw.length < 6) continue;
          const ym = `${dateRaw.substring(0,4)}-${dateRaw.substring(4,6)}`;
          
          // 금액(A)
          const inTransitA = parseNumber(row[25]);
          const onHandA = parseNumber(row[29]);
          const totalA = inTransitA + onHandA;
          let deadA = 0;
          for (let c = 50; c <= 53; c++) {
            deadA += parseNumber(row[c]);
          }
          let mrp8A = 0;
          for (let c = 71; c <= 78; c++) {
            mrp8A += parseNumber(row[c]);
          }
          const malA = parseNumber(row[53]); // Aging 181일 이상
          
          // 수량(Q)
          const inTransitQ = parseNumber(row[24]);
          const onHandQ = parseNumber(row[28]);
          const totalQ = inTransitQ + onHandQ;
          let deadQ = 0;
          for (let c = 39; c <= 42; c++) {
            deadQ += parseNumber(row[c]);
          }
          let mrp8Q = 0;
          for (let c = 55; c <= 62; c++) {
            mrp8Q += parseNumber(row[c]);
          }
          const malQ = parseNumber(row[42]); // Aging 181일 이상
          
          if (!historyData[key]) {
            historyData[key] = {};
          }
          historyData[key][ym] = {
            totalA, deadA, mrp8A, malA,
            totalQ, deadQ, mrp8Q, malQ
          };
        }
      }
    }

    let baseDate = new Date();
    if (curDateRaw && curDateRaw.length === 8) {
      baseDate = new Date(curDateRaw.substring(0,4), parseInt(curDateRaw.substring(4,6), 10)-1, curDateRaw.substring(6,8));
    }
    let cutoffDate = new Date(baseDate);
    cutoffDate.setFullYear(cutoffDate.getFullYear() - 1); // 기준일로부터 1년 전

    const issueSheet = ss.getSheetByName('출고이력');
    const movedKeys = new Set();
    if (issueSheet) {
      const issueLastRow = issueSheet.getLastRow();
      if (issueLastRow >= 2) {
        const issueData = issueSheet.getRange("A1:F" + issueLastRow).getValues();
        for (let i = 1; i < issueData.length; i++) {
          const row = issueData[i];
          const key = String(row[1]).trim();
          if (!key) continue;
          let issueDate = row[0];
          if (!(issueDate instanceof Date)) {
            const parts = String(issueDate).trim().split('-');
            if (parts.length >= 3) {
              issueDate = new Date(parts[0], parseInt(parts[1], 10) - 1, parts[2].substring(0, 2));
            }
          }
          if (issueDate instanceof Date && !isNaN(issueDate)) {
            if (issueDate > cutoffDate && issueDate <= baseDate) {
               let issueQty = row[5];
               if (typeof issueQty !== 'number') {
                   issueQty = parseFloat(String(issueQty).replace(/,/g, '')) || 0;
               }
               if (issueQty !== 0) {
                   movedKeys.add(key);
               }
            }
          }
        }
      }
    }

    return {
      success: true,
      historyData: historyData,
      movedKeys: Array.from(movedKeys)
    };
  } catch (e) {
    return {
      success: false,
      message: e.message
    };
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

function formatKoreanMoney(value) {
  const val = parseFloat(value);
  if (isNaN(val) || val === 0) return "0원";
  
  const absVal = Math.abs(val);
  const sign = val < 0 ? "-" : "";
  let formatted = "";
  
  if (absVal >= 1000000000) { 
    const scaled = Math.round(absVal / 100000000);
    formatted = scaled.toLocaleString() + "억";
  } else if (absVal >= 100000000) { 
    const scaled = (absVal / 100000000).toFixed(1);
    formatted = Number(scaled).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + "억";
  } else if (absVal >= 1000000) { 
    const scaled = Math.round(absVal / 1000000);
    formatted = scaled.toLocaleString() + "백만";
  } else { 
    const scaled = Math.round(absVal / 10000);
    formatted = scaled.toLocaleString() + "만원";
  }
  
  return sign + formatted;
}

function getTop10IncreasedItems(data, keyType) {
  let items = data.map(d => {
    let diffA = 0;
    let diffQ = 0;
    if (keyType === 'total') {
      diffA = (d.totalA || 0) - (d.prevTotalA || 0);
      diffQ = (d.totalQ || 0) - (d.prevTotalQ || 0);
    } else {
      diffA = (d.deadA || 0) - (d.prevDeadA || 0);
      diffQ = (d.deadQ || 0) - (d.prevDeadQ || 0);
    }
    
    let invDays = '-';
    if (d.mrp8A > 0) {
      invDays = Math.round(d.invDaysA).toLocaleString() + "일";
    } else {
      invDays = "계획 없음";
    }

    let stockAmt = keyType === 'total' ? (d.totalA || 0) : (d.deadA || 0);

    return {
      material: d.material || "",
      materialDesc: d.materialDesc || "",
      vendorDesc: d.vendorDesc || "미지정",
      stockAmt: stockAmt,
      diffA: diffA,
      diffQ: diffQ,
      invDays: invDays
    };
  });

  items = items.filter(it => it.diffA > 0);
  items.sort((a, b) => b.diffA - a.diffA);
  return items.slice(0, 10);
}

function buildIncreasedTableHtml(label, items) {
  if (!items || items.length === 0) {
    return `
    <div style="margin-top:8px;margin-bottom:15px;text-align:left;font-family:'굴림','Gulim',sans-serif;font-size:9pt;color:#333;">
        <span style="font-weight:bold;">- ${label} :</span>
        <div style="margin-top:6px;color:#6c757d;padding-left:5px;">해당 사항 없음</div>
    </div>`;
  }

  let tbodyHtml = items.map(it => {
    const formattedStockAmt = formatKoreanMoney(it.stockAmt || 0);
    const formattedAmt = it.diffA >= 0 ? "+" + formatKoreanMoney(it.diffA) : formatKoreanMoney(it.diffA);
    const formattedQty = it.diffQ >= 0 ? "+" + Math.round(it.diffQ).toLocaleString() + " <span style='color:#888888;font-weight:normal;'>EA</span>" : Math.round(it.diffQ).toLocaleString() + " <span style='color:#888888;font-weight:normal;'>EA</span>";
    return `
    <tr style="border-bottom:1px solid #dee2e6;">
        <td style="border:1px solid #dee2e6;color:#000;padding:6px;font-weight:bold;text-align:center;vertical-align:middle;">${it.material}</td>
        <td style="border:1px solid #dee2e6;color:#000;padding:6px;text-align:center;vertical-align:middle;">${it.materialDesc}</td>
        <td style="border:1px solid #dee2e6;color:#000;padding:6px;text-align:center;vertical-align:middle;">${it.vendorDesc}</td>
        <td style="border:1px solid #dee2e6;color:#dc3545;padding:6px;font-weight:bold;text-align:center;vertical-align:middle;">${formattedAmt}</td>
        <td style="border:1px solid #dee2e6;color:#dc3545;padding:6px;font-weight:bold;text-align:center;vertical-align:middle;">${formattedQty}</td>
        <td style="border:1px solid #dee2e6;color:#000;padding:6px;font-weight:bold;text-align:center;vertical-align:middle;">${formattedStockAmt}</td>
        <td style="border:1px solid #dee2e6;color:#000;padding:6px;font-weight:bold;text-align:center;vertical-align:middle;">${it.invDays}</td>
    </tr>`;
  }).join('');

  return `
  <div style="margin-top:8px;margin-bottom:15px;text-align:left;font-family:'굴림','Gulim',sans-serif;font-size:9pt;color:#333;">
      <span style="font-weight:bold;">- ${label} :</span>
      <div style="margin-top:6px;overflow-x:auto;">
          <table style="border-collapse:collapse;width:100%;color:#000;border:1px solid #dee2e6;font-size:9pt;font-family:'굴림','Gulim',sans-serif;margin-bottom:0;">
              <thead>
                  <tr style="background-color:#f8f9fa;font-weight:bold;color:#000;border-bottom:1px solid #dee2e6;">
                      <th style="background-color:#f1f3f5;font-weight:bold;width:110px;border:1px solid #dee2e6;color:#000;padding:6px;text-align:center;vertical-align:middle;">Material</th>
                      <th style="background-color:#f1f3f5;font-weight:bold;width:355px;border:1px solid #dee2e6;color:#000;padding:6px;text-align:center;vertical-align:middle;">Description</th>
                      <th style="background-color:#f1f3f5;font-weight:bold;width:110px;border:1px solid #dee2e6;color:#000;padding:6px;text-align:center;vertical-align:middle;">Vendor</th>
                      <th style="background-color:#f1f3f5;font-weight:bold;width:110px;border:1px solid #dee2e6;color:#000;padding:6px;text-align:center;vertical-align:middle;">증가금액</th>
                      <th style="background-color:#f1f3f5;font-weight:bold;width:110px;border:1px solid #dee2e6;color:#000;padding:6px;text-align:center;vertical-align:middle;">증가수량</th>
                      <th style="background-color:#f1f3f5;font-weight:bold;width:110px;border:1px solid #dee2e6;color:#000;padding:6px;text-align:center;vertical-align:middle;">재고금액</th>
                      <th style="background-color:#f1f3f5;font-weight:bold;width:110px;border:1px solid #dee2e6;color:#000;padding:6px;text-align:center;vertical-align:middle;">재고일수</th>
                  </tr>
              </thead>
              <tbody>
                  ${tbodyHtml}
              </tbody>
          </table>
      </div>
  </div>`;
}

function sendEmailReport(recipient, cc, subject, htmlBody, filterOptions) {
  try {
    let modifiedHtmlBody = htmlBody;
    const rawResult = getRawAnalysisData();
    if (rawResult && rawResult.success && rawResult.data) {
      let mergedData = rawResult.data;
      if (filterOptions && filterOptions.allowedKeys) {
        const allowedKeysSet = new Set(filterOptions.allowedKeys);
        mergedData = mergedData.filter(d => allowedKeysSet.has(d.key));
      }
      
      const topTotInc = getTop10IncreasedItems(mergedData, 'total');
      const topDeadInc = getTop10IncreasedItems(mergedData, 'dead');
      const tableTotHtml = buildIncreasedTableHtml("총재고 증가 상위품목", topTotInc);
      const tableDeadHtml = buildIncreasedTableHtml("부진재고 증가 상위품목", topDeadInc);
      
      const regexTot = /(- <b>총재고\s+상위품목\s*:[\s\S]*?<\/table>\s*<\/div>\s*<\/div>)/i;
      if (regexTot.test(modifiedHtmlBody)) {
        modifiedHtmlBody = modifiedHtmlBody.replace(regexTot, `$1\n${tableTotHtml}`);
      }
      
      const regexDead = /(- <b>부진재고\s+상위품목\s*:[\s\S]*?<\/table>\s*<\/div>\s*<\/div>)/i;
      if (regexDead.test(modifiedHtmlBody)) {
        modifiedHtmlBody = modifiedHtmlBody.replace(regexDead, `$1\n${tableDeadHtml}`);
      }
    }
    
    // 이메일 본문 압축 (불필요한 공백 및 탭, 줄바꿈 일괄 소거하여 데이터 크기 극단적 축소)
    modifiedHtmlBody = modifiedHtmlBody.replace(/\s+/g, ' ').replace(/>\s+</g, '><').trim();
    
    const options = {
      htmlBody: modifiedHtmlBody
    };
    if (cc) {
      options.cc = cc;
    }
    
    GmailApp.sendEmail(recipient, subject, "본 메일은 HTML 형식을 지원하는 메일 클라이언트에서 확인하실 수 있습니다.", options);
    return { success: true, message: `메일 발송 성공\n\n수신: ${recipient}${cc ? '\n참조: ' + cc : ''}` };
  } catch (e) {
    return { success: false, message: `메일 발송 실패: ${e.message}` };
  }
}

function getMailRouting(selectedPlant, selectedPurchaser, selectedPart) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let debugLog = [];
    let sheet = ss.getSheetByName('메일발송조건') || ss.getSheetByName('메일 발송 조건');
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        const name = sheets[i].getName();
        if (name.includes("발송조건") || name.includes("발송 조건")) {
          sheet = sheets[i];
          break;
        }
      }
    }
    
    if (!sheet) {
      debugLog.push("발송조건 시트를 찾을 수 없습니다.");
      return { recipient: "mh501.kim@samsung.com", cc: "dohee.cho@samsung.com", isDefault: true, debugLog: debugLog };
    }
    
    debugLog.push(`발송조건 시트 발견: ${sheet.getName()}`);
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) {
      return { recipient: "mh501.kim@samsung.com", cc: "dohee.cho@samsung.com", isDefault: true, debugLog: debugLog };
    }
    
    const data = sheet.getRange("B4:N" + lastRow).getValues();
    const toSet = new Set();
    const ccSet = new Set();
    
    const cleanId = id => {
      let str = String(id).trim();
      if (!str) return "";
      str = str.replace(/\(.*?\)/g, "").replace(/\[.*?\]/g, "").trim();
      if (!str) return "";
      if (str.includes("@")) return str.toLowerCase();
      return str.toLowerCase() + "@samsung.com";
    };
    
    const matchFilter = (filterVal, selectedVal) => {
      let f = String(filterVal).trim();
      f = f.replace(/\(.*?\)/g, "").replace(/\[.*?\]/g, "").trim().toUpperCase();
      const s = String(selectedVal).trim().toUpperCase();
      if (!f || f === "ALL" || f === "전체") return true;
      return f === s;
    };
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowPlant = String(row[0]).trim();
      const rowPurchaser = String(row[1]).trim();
      const rowPart = String(row[2]).trim();
      
      const pMatch = matchFilter(rowPlant, selectedPlant);
      const uMatch = matchFilter(rowPurchaser, selectedPurchaser);
      const rMatch = matchFilter(rowPart, selectedPart);
      
      if (pMatch && uMatch && rMatch) {
        for (let col = 3; col <= 7; col++) {
          const email = cleanId(row[col]);
          if (email) {
            toSet.add(email);
          }
        }
        for (let col = 8; col <= 12; col++) {
          const email = cleanId(row[col]);
          if (email) {
            ccSet.add(email);
          }
        }
      }
    }
    
    let recipients = [...toSet].join(",");
    let ccs = [...ccSet].join(",");
    let isDefault = false;
    if (!recipients) {
      recipients = "mh501.kim@samsung.com";
      ccs = "dohee.cho@samsung.com";
      isDefault = true;
    }
    return { recipient: recipients, cc: ccs, isDefault: isDefault, debugLog: debugLog };
  } catch (e) {
    return { recipient: "mh501.kim@samsung.com", cc: "dohee.cho@samsung.com", error: e.message, debugLog: [e.toString()] };
  }
}

function getMailRoutingList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('메일발송조건') || ss.getSheetByName('메일 발송 조건');
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        const name = sheets[i].getName();
        if (name.includes("발송조건") || name.includes("발송 조건")) {
          sheet = sheets[i];
          break;
        }
      }
    }
    
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    
    const data = sheet.getRange("B4:N" + lastRow).getValues();
    const list = [];
    
    const cleanId = id => {
      let str = String(id).trim();
      if (!str) return "";
      str = str.replace(/\(.*?\)/g, "").replace(/\[.*?\]/g, "").trim();
      if (!str) return "";
      if (str.includes("@")) return str.toLowerCase();
      return str.toLowerCase() + "@samsung.com";
    };
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const plant = String(row[0]).trim().replace(/\(.*?\)/g, "").trim();
      const purchaser = String(row[1]).trim().replace(/\(.*?\)/g, "").trim();
      const part = String(row[2]).trim().replace(/\(.*?\)/g, "").trim();
      
      const toSet = new Set();
      const ccSet = new Set();
      
      for (let col = 3; col <= 7; col++) {
        const email = cleanId(row[col]);
        if (email) toSet.add(email);
      }
      for (let col = 8; col <= 12; col++) {
        const email = cleanId(row[col]);
        if (email) ccSet.add(email);
      }
      
      const recipients = [...toSet].join(",");
      const ccs = [...ccSet].join(",");
      
      if (recipients) {
        list.push({
          plant: plant || "ALL",
          purchaser: purchaser || "ALL",
          part: part || "ALL",
          recipient: recipients,
          cc: ccs
        });
      }
    }
    return list;
  } catch (e) {
    return [];
  }
}