function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const SHEET_NAME = 'SCM Raw';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;
const COLUMN_MAPPING = {
    "No": 0,          // A열
    "Vendor Code": 1, // B열
    "Vendor Desc": 2, // C열
    "Maker": 3,       // D열
    "개발구매": 4,    // E열
    "조달구매": 5,    // F열
    "협력사": 6,      // G열
    "국가": 7,        // H열
    "도시": 8,        // I열
    "lat": 9,         // J열 (위도)
    "lon": 10,        // K열 (경도)
    "품목": 11,       // L열
    "품목군": 12,     // M열 (구분자)
    "생산유형": 13,   // N열 (형태)
    "Risk": 14,       // O열 (Risk 테마)
    "추가1": 15,      // P열
    "추가2": 16,      // Q열
    "추가3": 17,      // R열
    "추가4": 18,      // S열
    "추가5": 19,      // T열
};

const COUNTRY_CODE_MAP = {
  "US": "미국",
  "CN": "중국",
  "JP": "일본",
  "KR": "대한민국",
  "VN": "베트남",
  "MY": "말레이시아",
  "TW": "대만",
  "DE": "독일",
  "PH": "필리핀",
  "IN": "인도",
  "ID": "인도네시아",
  "TH": "태국",
  "SG": "싱가포르",
  "MX": "멕시코",
  "GB": "영국",
  "FR": "프랑스",
  "IT": "이탈리아",
  "CA": "캐나다",
  "BR": "브라질"
};

const TREEVIEW_COLUMNS = ["국가", "도시", "협력사", "품목", "품목군", "생산유형"];

function doGet(e) {

  logAccess('SCM Map');

  const template = HtmlService.createTemplateFromFile('index');
  template.initialPageDataString = JSON.stringify(getScmData());
  return template.evaluate()
      .setTitle('SCM Map')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getScmData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`시트 '${SHEET_NAME}'를 찾을 수 없습니다.`);
    
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) {
      return { success: true, filters: {}, allData: [], riskThemes: [], columnMapping: COLUMN_MAPPING, treeviewColumns: TREEVIEW_COLUMNS, userEmail: Session.getActiveUser().getEmail(), countryCodeMap: COUNTRY_CODE_MAP, vendorCodeTypes: { regular: [], ipc: [] } };
    }
    
    const maxCols = Math.max(sheet.getLastColumn(), 27);
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, maxCols);
    const allDataValues = dataRange.getDisplayValues();
    
    // Dynamically construct "생산유형" (index 13, column N) from S~AA columns (indexes 18 to 26)
    allDataValues.forEach(row => {
      const roles = [];
      for (let i = 18; i <= 26; i++) {
        const val = row[i] ? String(row[i]).trim() : "";
        if (val) {
          roles.push(val);
        }
      }
      // Overwrite column N (index 13) with comma-separated roles
      row[13] = roles.join(', ');
    });
    
    const filters = {};
    const filterableColumns = ["국가", "도시", "품목", "품목군", "생산유형", "협력사"];
    
    filterableColumns.forEach(colName => {
        const colIdx = COLUMN_MAPPING[colName];
        if (colIdx === undefined) {
          console.error("'" + colName + "'에 해당하는 열을 COLUMN_MAPPING에서 찾을 수 없습니다.");
          return;
        }
        let uniqueValues;
        if (colName === "생산유형") {
          const valuesSet = new Set();
          allDataValues.forEach(row => {
            const val = row[colIdx];
            if (val) {
              val.split(',').forEach(v => {
                let trimmed = v.trim();
                if (trimmed.startsWith("Office")) {
                  trimmed = "Office";
                }
                if (trimmed) valuesSet.add(trimmed);
              });
            }
          });
          uniqueValues = [...valuesSet];
          
          // Sort "생산유형" filters: FAB, Package, Test, 원재료, 생산, 가공, Office, and others alphabetically
          const priority = ["FAB", "Package", "Test", "원재료", "생산", "가공", "Office"];
          uniqueValues.sort((a, b) => {
            const idxA = priority.indexOf(a);
            const idxB = priority.indexOf(b);
            if (idxA !== -1 && idxB !== -1) return idxA - idxB;
            if (idxA !== -1) return -1;
            if (idxB !== -1) return 1;
            return a.localeCompare(b);
          });
        } else {
          uniqueValues = [...new Set(allDataValues.map(row => row[colIdx]))].filter(Boolean).sort();
          if (colName === "국가") {
            // Sort "국가" filters: KR, TW, CN, US first, and others alphabetically
            const priority = ["KR", "TW", "CN", "US"];
            uniqueValues.sort((a, b) => {
              const idxA = priority.indexOf(a);
              const idxB = priority.indexOf(b);
              if (idxA !== -1 && idxB !== -1) return idxA - idxB;
              if (idxA !== -1) return -1;
              if (idxB !== -1) return 1;
              return a.localeCompare(b);
            });
          }
        }
        filters[colName] = uniqueValues;
    });
    
    const riskColIdx = COLUMN_MAPPING['Risk'];
    const riskThemes = new Set();
    
    if (riskColIdx !== undefined) {
      allDataValues.forEach(row => {
          if (row[riskColIdx]) {
              row[riskColIdx].split('#').map(t => t.trim()).filter(Boolean).forEach(theme => riskThemes.add(theme));
          }
      });
    }
    
    return {
      success: true,
      filters: filters,
      allData: allDataValues,
      riskThemes: [...riskThemes].sort(),
      columnMapping: COLUMN_MAPPING,
      treeviewColumns: TREEVIEW_COLUMNS,
      userEmail: Session.getActiveUser().getEmail(),
      countryCodeMap: COUNTRY_CODE_MAP,
      vendorCodeTypes: getVendorCodeTypes()
    };
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
    
    // Regular vendors
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
    
    // IPC vendors
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
    try { userEmail = Session.getActiveUser().getEmail(); } catch(e) {}
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

function getPartnerDetailByCode(vendorCode) {
  try {
    const ss = SpreadsheetApp.openById("1rQHrhLa3xNZ2DyfK1OS4jeF17xkfXobkV7joMoq4A4w");
    
    // Helper function for ultra-robust matching
    function isCodeMatched(codeA, codeB) {
      const strA = String(codeA || '').trim();
      const strB = String(codeB || '').trim();
      if (!strA || !strB) return false;
      if (strA.toLowerCase() === strB.toLowerCase()) return true;
      
      const numA = parseInt(strA, 10);
      const numB = parseInt(strB, 10);
      if (!isNaN(numA) && !isNaN(numB) && numA === numB) {
        return true;
      }
      return false;
    }
    
    // 1. Try to search in "협력업체 관리대장" sheet first
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
          const code = codeDisplayValues[i][0];
          if (isCodeMatched(code, vendorCode)) {
            matchedRowIdx = i + 5;
            break;
          }
        }
      }
    }
    
    // 2. If not found, try to search in "IPC Vendor" sheet as fallback
    if (matchedRowIdx === -1) {
      sheet = ss.getSheetByName("IPC Vendor");
      if (sheet) {
        const lastRow = sheet.getLastRow();
        lastCol = sheet.getLastColumn();
        if (lastRow >= 5) {
          const codeDisplayValues = sheet.getRange(5, 6, lastRow - 4, 1).getDisplayValues();
          for (let i = 0; i < codeDisplayValues.length; i++) {
            const code = codeDisplayValues[i][0];
            if (isCodeMatched(code, vendorCode)) {
              matchedRowIdx = i + 5;
              isIpc = true; // Mark as IPC Vendor!
              break;
            }
          }
        }
      }
    }
    
    if (matchedRowIdx === -1) {
      return { success: false, message: `협력업체 코드 '${vendorCode}'에 일치하는 정규협력사 또는 IPC Vendor를 찾을 수 없습니다.` };
    }
    
    // Read the exact display values of that specific row to preserve formats
    const displayRow = sheet.getRange(matchedRowIdx, 1, 1, lastCol).getDisplayValues()[0];
    
    const COLUMN_MAPPING = [
      3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 98, 26, 27, 28, 29, 30,
      31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48,
      58, 59, 60, 61, 62, 63, 64, 65, 66, 67,
      68, 69, 70, 71, 72, 73, 74, 75, 76, 77,
      78, 79, 80, 81, 82, 83, 84, 85, 86, 87,
      88, 89, 90, 91, 92, 93, 94, 95, 96, 97,
      99, 102, 105, 108,
      111, 112, 113, 114, 115, 116, 117
    ];
    
    const rowValues = COLUMN_MAPPING.map(colIdx => {
      return displayRow[colIdx - 1] || '';
    });
    
    // Construct the same item object as showPartnerDetail expects
    const item = {
      values: rowValues,
      isIpc: isIpc, // Add isIpc property!
      lat: displayRow[24 - 1] || '',           // Column X (24): 본사 위도
      lon: displayRow[25 - 1] || '',           // Column Y (25): 본사 경도
      factory1Lat: displayRow[100 - 1] || '',  // Column CV (100): 공장1 위도
      factory1Lon: displayRow[101 - 1] || '',  // Column CW (101): 공장1 경도
      factory2Lat: displayRow[103 - 1] || '',  // Column CY (103): 공장2 위도
      factory2Lon: displayRow[104 - 1] || '',  // Column CZ (104): 공장2 경도
      factory3Lat: displayRow[106 - 1] || '',  // Column DB (106): 공장3 위도
      factory3Lon: displayRow[107 - 1] || '',  // Column DC (107): 공장3 경도
      factory4Lat: displayRow[109 - 1] || '',  // Column DE (109): 공장4 위도
      factory4Lon: displayRow[110 - 1] || ''   // Column DF (110): 공장4 경도
    };
    
    return { success: true, item: item };
    
  } catch (e) {
    return { success: false, message: "상세 정보 조회 중 오류가 발생했습니다: " + e.message };
  }
}