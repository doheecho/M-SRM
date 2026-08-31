const SHEET_NAME = '협력업체 관리대장';
const HEADER_ROW = 4;
const DATA_START_ROW = 5;

const COLUMN_MAPPING = [
  3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 98, 26, 27, 28, 29, 30, // 25 columns (W열의 23(Map)을 제거하고 CT열의 98(주거래사업부)을 해당 위치로 앞당김!)
  31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, // 18 columns (2013년~2030년 당사구매액)
  58, 59, 60, 61, 62, 63, 64, 65, 66, 67, // 10 columns (매출액 21년~30년)
  68, 69, 70, 71, 72, 73, 74, 75, 76, 77, // 10 columns (영업이익 21년~30년)
  78, 79, 80, 81, 82, 83, 84, 85, 86, 87, // 10 columns (당기순이익 21년~30년)
  88, 89, 90, 91, 92, 93, 94, 95, 96, 97, // 10 columns (이익잉여금 21년~30년)
  99, 102, 105, 108, // 4 columns (주소 공장1, 주소 공장2, 주소 공장3, 주소 공장4 - 한칸씩 왼쪽 자동 시프트)
  111, 112, 113, 114, 115, 116, 117 // 7 columns (Etc1~Etc7 - 한칸씩 왼쪽 자동 시프트)
];

const STATUS_COL = 3;       
const PLANT_COL = 8; // [수정] J열(10)에서 2개 칸 왼쪽 시프트된 H열(8)       
const TYPE_COL = 9; // [수정] K열(11)에서 2개 칸 왼쪽 시프트된 I열(9)         
const BIZ_TYPE_COL = 11; // [수정] M열(13)에서 2개 칸 왼쪽 시프트된 K열(11)     
const SUBCONTRACT_COL = 15; // [수정] Q열(17)에서 2개 칸 왼쪽 시프트된 O열(15)  
const LISTING_STATUS_COL = 17; // [수정] S열(19)에서 2개 칸 왼쪽 시프트된 Q열(17) 
const REVENUE_COL = 57; // [유지] BE열(57)      
const LATITUDE_COL = 24;     
const LONGITUDE_COL = 25;    

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  const params = {
    status: (e && e.parameter && e.parameter.status) || '정규',
    plant: (e && e.parameter && e.parameter.plant) || '전체',
    type: (e && e.parameter && e.parameter.type) || '전체',
    bizType: (e && e.parameter && e.parameter.bizType) || '전체',
    subcontract: (e && e.parameter && e.parameter.subcontract) || '전체',
    listingStatus: (e && e.parameter && e.parameter.listingStatus) || '전체',
    revenue: (e && e.parameter && e.parameter.revenue) || '전체',
    partnerType: (e && e.parameter && e.parameter.partnerType) || '당사 거래선'
  };
  
  template.initialDataString = JSON.stringify(getData(params));
  
  return template.evaluate()
    .setTitle('구매 협력업체 관리대장')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const SECURITY_CONFIG = {
  maskedColumnNumbers: [
    20, 21, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
  ]
};

// 강력한 구글 세션 이메일 추출 헬퍼 함수
function getActiveUserEmail() {
  let email = "";
  try {
    email = Session.getActiveUser().getEmail();
  } catch(e) {}
  
  if (!email || email.trim() === "" || email === "Unknown") {
    try {
      email = Session.getEffectiveUser().getEmail();
    } catch(e) {}
  }
  
  if (!email || email.trim() === "") {
    email = "dohee.cho@ai.samsunghealthcare.com"; // 기본 폴백 이메일
  }
  return email;
}

function getData(params) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // 협력사 구분 선택지에 따라 워크시트 선택 (당사 거래선 -> '협력업체 관리대장', IPC 거래선 -> 'IPC Vendor')
    const selectedSheetName = (params && params.partnerType === 'IPC 거래선') ? 'IPC Vendor' : SHEET_NAME;
    const sheet = spreadsheet ? spreadsheet.getSheetByName(selectedSheetName) : null;
    if (!sheet) return { success: false, message: `'${selectedSheetName}' 시트를 찾을 수 없습니다.` };
    
    const currentUserEmail = getActiveUserEmail();
    const hasFullAccess = checkUserFullAccess(currentUserEmail);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < DATA_START_ROW) {
      const headersArray = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
      return { success: true, headers: getHeaders(headersArray), data: [], currentFilters: params, filterOptions: {}, dashboardMeta: [] };
    }
    
    const allSheetData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headersArray = allSheetData[HEADER_ROW - 1];
    const headers = getHeaders(headersArray);
    
    const fullRawRows = allSheetData.slice(DATA_START_ROW - 1);
    
    const plantOptions = new Set(), typeOptions = new Set(), bizTypeOptions = new Set(), revenueOptions = new Set(), listingStatusOptions = new Set();
    
    const allData = fullRawRows.map(row => {
      const statusVal = String(row[STATUS_COL - 1] || '').trim();
      const plantVal = String(row[PLANT_COL - 1] || '').trim();
      const typeVal = String(row[TYPE_COL - 1] || '').trim();
      const bizTypeVal = String(row[BIZ_TYPE_COL - 1] || '').trim();
      const subcontractVal = String(row[SUBCONTRACT_COL - 1] || '').trim();
      const listingStatusVal = String(row[LISTING_STATUS_COL - 1] || '').trim();
      const revenueVal = String(row[REVENUE_COL - 1] || '').trim();
      
      if(plantVal) plantOptions.add(plantVal);
      if(typeVal) typeOptions.add(typeVal);
      if(bizTypeVal) bizTypeOptions.add(bizTypeVal);
      if(listingStatusVal) listingStatusOptions.add(listingStatusVal);
      if(revenueVal) revenueOptions.add(revenueVal);
      
      const rowValues = COLUMN_MAPPING.map(colIndex => {
        let val = row[colIndex - 1] || '';
        
        if (!hasFullAccess) {
          const isColNumberMatched = SECURITY_CONFIG.maskedColumnNumbers.indexOf(colIndex) !== -1;
          
          if (isColNumberMatched) {
            val = 'MASKED_SECURE_DATA';
          }
        }
        return val;
      });
      
      return {
        values: rowValues,
        lat: row[LATITUDE_COL - 1] || '',
        lon: row[LONGITUDE_COL - 1] || '',
        factory1Lat: row[100 - 1] || '', // Column CV (100): 공장1 위도
        factory1Lon: row[101 - 1] || '', // Column CW (101): 공장1 경도
        factory2Lat: row[103 - 1] || '', // Column CY (103): 공장2 위도
        factory2Lon: row[104 - 1] || '', // Column CZ (104): 공장2 경도
        factory3Lat: row[106 - 1] || '', // Column DB (106): 공장3 위도
        factory3Lon: row[107 - 1] || '', // Column DC (107): 공장3 경도
        factory4Lat: row[109 - 1] || '', // Column DE (109): 공장4 위도
        factory4Lon: row[110 - 1] || '', // Column DF (110): 공장4 경도
        filters: {
          status: statusVal,
          plant: plantVal,
          type: typeVal,
          bizType: bizTypeVal,
          subcontract: subcontractVal,
          listingStatus: listingStatusVal,
          revenue: revenueVal
        }
      };
    });
    
    const revenueSortOrder = ['100억 미만', '100~500억', '500~1,000억', '1,000억 이상', '미확인'];
    const sortedRevenues = Array.from(revenueOptions).sort((a, b) => {
        const indexA = revenueSortOrder.indexOf(a);
        const indexB = revenueSortOrder.indexOf(b);
        if (indexA === -1) return 1;
        if (indexB === -1) return -1;
        return indexA - indexB;
    });
    
    const listingStatusSortOrder = ['코스피', '코스닥', '코넥스', '비상장', '개인사업자', '외자'];
    const sortedListingStatuses = Array.from(listingStatusOptions).sort((a,b) => {
        const indexA = listingStatusSortOrder.indexOf(a);
        const indexB = listingStatusSortOrder.indexOf(b);
        if (indexA === -1) return 1;
        if (indexB === -1) return -1;
        return indexA - indexB;
    });
    
    const dashboardMeta = getDashboardMeta(spreadsheet);
    
    return {
      success: true,
      headers: headers,
      allData: allData,
      currentFilters: params,
      filterOptions: {
        plants: ['전체', ...Array.from(plantOptions).sort()],
        types: ['전체', ...Array.from(typeOptions).sort()],
        bizTypes: ['전체', ...Array.from(bizTypeOptions).sort()],
        subcontracts: ['전체', '대기업', '중견', '중소', '비대상'],
        listingStatuses: ['전체', ...sortedListingStatuses],
        revenues: ['전체', ...sortedRevenues],
      },
      dashboardMeta: dashboardMeta,
      userEmail: currentUserEmail,
      hasFullAccess: hasFullAccess
    };
  } catch (e) {
    return { success: false, message: `데이터 로딩 중 오류 발생: ${e.message}` };
  }
}

function getDashboardMeta(spreadsheet) {
    try {
        const dashboardSheet = spreadsheet.getSheetByName('대시보드');
        if (!dashboardSheet) return [];
        const dbData = dashboardSheet.getDataRange().getDisplayValues();
        
        function extractRange(rangeStr) {
            const match = rangeStr.match(/([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)/);
            if (!match) return [];
            let cStart = 0; for(let i=0;i<match[1].length;i++) cStart = cStart*26 + match[1].charCodeAt(i)-64; cStart--;
            const rStart = parseInt(match[2], 10) - 1;
            let cEnd = 0; for(let i=0;i<match[3].length;i++) cEnd = cEnd*26 + match[3].charCodeAt(i)-64; cEnd--;
            const rEnd = parseInt(match[4], 10) - 1;
            
            const res = [];
            for(let r=rStart; r<=rEnd; r++) {
                if (dbData[r]) res.push(dbData[r].slice(cStart, cEnd + 1));
            }
            return res;
        }
        const chartConfigs = {
            'status':        { range: 'B4:C6',   isSubcontract: false },
            'plant':         { range: 'E4:F7',   isSubcontract: false },
            'type':          { range: 'H4:I6',   isSubcontract: false },
            'bizType':       { range: 'K4:L10',  isSubcontract: false },
            'subcontract':   { range: 'N4:O10',  isSubcontract: true },
            'listingStatus': { range: 'Q4:R10',  isSubcontract: false }                   
        };
        
        return Object.keys(chartConfigs).map(key => {
            const config = chartConfigs[key];
            const values = extractRange(config.range);
            if (!values || values.length === 0) return null;
            
            const title = values[0][0];
            const labels = [];
            for (let i = 1; i < values.length; i++) {
                if (values[i][0]) labels.push(values[i][0].trim());
            }
            return { key, title, labels, isSubcontract: config.isSubcontract };
        }).filter(Boolean);
    } catch(e) {
        return [];
    }
}

function getHeaders(fullHeadersArray) {
    try {
        const mainHeaders = COLUMN_MAPPING.map(colIndex => ({ name: fullHeadersArray[colIndex - 1] || '', originalColumn: colIndex }));
        const summaryHeaders = new Array(mainHeaders.length).fill(null);
        
        const summaryTitles = { C: '거래정보', F: '기본정보', Z: '계약정보', AE: '당사 구매액 (백만원)', BF: '매출액 (백만원)', BP: '영업이익 (백만원)', BZ: '당기순이익 (백만원)', CJ: '이익잉여금 (백만원)', Etc1: 'Etc.' };
        const SUMMARY_HEADER_COLS = { C: 3, F: 6, Z: 26, AE: 31, BF: 58, BP: 68, BZ: 78, CJ: 88, Etc1: 111 };
        
        const summaryKeys = Object.keys(SUMMARY_HEADER_COLS);
        for (let i = 0; i < summaryKeys.length; i++) {
            const currentKey = summaryKeys[i];
            const currentCol = SUMMARY_HEADER_COLS[currentKey];
            const currentIndex = mainHeaders.findIndex(h => h.originalColumn === currentCol);
            if (currentIndex === -1) continue;
            
            const nextKey = (i + 1 < summaryKeys.length) ? summaryKeys[i+1] : null;
            const nextCol = nextKey ? SUMMARY_HEADER_COLS[nextKey] : -1;
            let nextIndex = nextCol !== -1 ? mainHeaders.findIndex(h => h.originalColumn === nextCol) : -1;
            if (nextIndex === -1) nextIndex = mainHeaders.length;
            
            const span = nextIndex - currentIndex;
            if (span > 0) summaryHeaders[currentIndex] = { title: summaryTitles[currentKey], span: span };
        }
        return { summary: summaryHeaders, main: mainHeaders };
    } catch(e) {
        return { summary: [], main: [] };
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
    
    const userEmail = getActiveUserEmail();
    const userId = userEmail.split('@')[0];
    
    logSheet.appendRow([new Date(), userId, moduleTitle]);
  } catch (e) {
  }
}

function fetchUserDeptAndName() {
  try {
    const email = getActiveUserEmail();
    const userId = email.split('@')[0];
    
    // 로컬 폴백용 사용자 정보 맵 (네트워크 단절 및 멀티 구글 로그인 충댈 대비)
    const userFallbackMap = {
      "dohee.cho": { name: "조도희", dept: "구매그룹" },
      "hm2521.kwon": { name: "권혁민", dept: "구매그룹" },
      "ky3247.kim": { name: "김경율", dept: "구매그룹" },
      "nt.kim": { name: "김나단", dept: "구매그룹" },
      "mh501.kim": { name: "김미현", dept: "구매그룹" },
      "kmjun.kim": { name: "김민준", dept: "구매그룹" },
      "yumi91.kim": { name: "김유미", dept: "구매그룹" },
      "yubeom.kim": { name: "김유범", dept: "구매그룹" },
      "es0901.kim": { name: "김은선", dept: "구매그룹" },
      "jung.bae.kim": { name: "김정배", dept: "구매그룹" },
      "jinchul2.kim": { name: "김진철", dept: "구매그룹" },
      "hr1206.kim": { name: "김효령", dept: "구매그룹" },
      "seongoh.noh": { name: "노성오", dept: "구매그룹" },
      "sehyun4.park": { name: "박세현", dept: "구매그룹" },
      "jy3124.park": { name: "박재용", dept: "구매그룹" },
      "jps.baek": { name: "백정필", stroke: "구매그룹" },
      "jungjin.seo": { name: "서정진", dept: "구매그룹" },
      "juwon.seo": { name: "서주원", dept: "구매그룹" },
      "mjnew.wang": { name: "왕민정", dept: "구매그룹" },
      "sangduek.lee": { name: "이상득", dept: "구매그룹" },
      "eunho3.lee": { name: "이은호", dept: "구매그룹" },
      "je0408.lee": { name: "이정은", dept: "구매그룹" },
      "sujin.jeong": { name: "정수진", dept: "구매그룹" },
      "jinsang.jung": { name: "정진상", dept: "구매그룹" },
      "hyoseok.cho": { name: "조효석", dept: "구매그룹" },
      "jisoon8.park": { name: "박지순", dept: "구매그룹" },
      "tu.kang": { name: "강태욱", dept: "구매그룹" },
      "syn.joo": { name: "주승연", dept: "구매그룹" }
    };
    
    let displayName = "";
    let loadedFromSheet = false;
    try {
      const ss = SpreadsheetApp.openById("1-qBeuf94mvboL0zZkfVgEAtgKsZiAAD6A6d9XjPlwb8");
      if (ss) {
        const sheet = ss.getSheetByName("사용자정보");
        if (sheet) {
          const lastRow = sheet.getLastRow();
          if (lastRow > 0) {
            const data = sheet.getDataRange().getDisplayValues();
            for (let i = 0; i < data.length; i++) {
              if (data[i][0] && String(data[i][0]).toLowerCase().trim() === userId.toLowerCase().trim()) {
                const name = data[i][1] || '';
                const dept = data[i][2] || '';
                displayName = dept ? `${dept} ${name}` : name;
                loadedFromSheet = true;
                break;
              }
            }
          }
        }
      }
    } catch(e) {}
    
    if (!loadedFromSheet || !displayName) {
      const fallbackUser = userFallbackMap[userId];
      if (fallbackUser) {
        displayName = `${fallbackUser.dept} ${fallbackUser.name}`;
      } else {
        displayName = `구매그룹 ${userId ? userId.toUpperCase() : "USER"}`;
      }
    }
    return displayName;
  } catch (e) {
    return "구매그룹 조도희";
  }
}

function checkUserFullAccess(email) {
  try {
    if (!email) {
      email = getActiveUserEmail();
    }
    const userId = email.split('@')[0];
    const ss = SpreadsheetApp.openById("1-qBeuf94mvboL0zZkfVgEAtgKsZiAAD6A6d9XjPlwb8"); 
    if (!ss) return false;
    const sheet = ss.getSheetByName("사용자정보");
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).toLowerCase().trim() === userId.toLowerCase().trim()) {
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
    if (!ss) return [];
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