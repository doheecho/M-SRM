function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  
  template.randomImageFile = 'image_data1';
  
  var greetings = ["행복 가득한 하루 되세요!", "오늘도 힘내세요!", "성공적인 하루를 응원합니다.", "오늘은 부디 칼퇴하세요!", "내일도 부디 칼퇴하세요!", "머리아픈 일은 다음에 하세요!", "오늘은 이쁜말만 하세요", "졸리면 어디가서 자고오세요", "주말에는 가족과 함께!"];
  template.randomGreeting = greetings[Math.floor(Math.random() * greetings.length)];
  template.archivedNews = JSON.stringify(getArchivedNews());
  
  return template.evaluate()
      .setTitle('메-SRM')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getUserDetails() {
  let email = 'Unknown';
  try { 
      email = Session.getActiveUser().getEmail(); 
  } catch(e) { 
      Logger.log("getActiveUser() 에러: " + e.toString());
  }
  
  // 구글 보안 정책 및 멀티 로그인 이슈로 getEmail()이 빈 값("")을 반환할 수 있으므로 강력한 폴백 구성
  if (!email || email.trim() === "" || email === "Unknown") {
    try {
      email = Session.getEffectiveUser().getEmail();
    } catch(e) {
      email = 'dohee.cho@ai.samsunghealthcare.com';
    }
  }
  
  const userId = email.split('@')[0];
  let userName = "";
  let deptName = "";
  
  // 로컬 폴백용 사용자 정보 맵 (스프레드시트가 연동 해제되거나 시트 에러가 발생해도 완벽 작동하도록 구성)
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
    "jps.baek": { name: "백정필", dept: "구매그룹" },
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

  let loadedFromSheet = false;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      const sheet = ss.getSheetByName("사용자정보");
      if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
          for (let i = 0; i < data.length; i++) {
            if (data[i][0] && String(data[i][0]).trim() === userId) {
              if (data[i][1]) userName = String(data[i][1]).trim();
              if (data[i][2]) deptName = String(data[i][2]).trim();
              loadedFromSheet = true;
              break;
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log("getUserDetails 스프레드시트 로드 에러: " + e.toString());
  }
  
  // 스프레드시트 로드가 실패했거나 사용자가 없는 경우 로컬 폴백 맵 적용
  if (!loadedFromSheet || !userName) {
    const fallbackUser = userFallbackMap[userId];
    if (fallbackUser) {
      userName = fallbackUser.name;
      deptName = fallbackUser.dept;
    } else {
      userName = userId ? userId.toUpperCase() : "사용자";
      deptName = "구매그룹";
    }
  }
  
  return {
    email: email,
    id: userId,
    name: userName,
    dept: deptName
  };
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
    let email = '';
    try {
      email = Session.getActiveUser().getEmail();
    } catch(e) {}
    
    if (!email || email.trim() === "" || email === "Unknown") {
      try {
        email = Session.getEffectiveUser().getEmail();
      } catch(e) {
        email = 'dohee.cho@ai.samsunghealthcare.com';
      }
    }
    
    const userId = email.split('@')[0];
    
    // 로컬 폴백용 사용자 정보 맵
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
      "jps.baek": { name: "백정필", dept: "구매그룹" },
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
    
    let loaded = false;
    let displayName = "";
    try {
      const ss = SpreadsheetApp.openById("1-qBeuf94mvboL0zZkfVgEAtgKsZiAAD6A6d9XjPlwb8");
      const sheet = ss.getSheetByName("사용자정보");
      if (sheet) {
        const data = sheet.getDataRange().getDisplayValues();
        for (let i = 0; i < data.length; i++) {
          if (data[i][0] === userId) {
            const name = data[i][1] || '';
            const dept = data[i][2] || '';
            displayName = dept ? `${dept} ${name}` : name;
            loaded = true;
            break;
          }
        }
      }
    } catch(e) {
    }
    
    if (!loaded || !displayName) {
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
      try {
        email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'dohee.cho@ai.samsunghealthcare.com';
      } catch(e) {
        email = 'dohee.cho@ai.samsunghealthcare.com';
      }
    }
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

function buildRssRequest(query, isDomestic) {
  let encodedQuery = encodeURIComponent(query); 
  let url = isDomestic
    ? `https://news.google.com/rss/search?q=${encodedQuery}&hl=ko&gl=KR&ceid=KR:ko`
    : `https://news.google.com/rss/search?q=${encodedQuery}&hl=en-US&gl=US&ceid=US:en`;
  
  return {
    url: url,
    method: "get",
    muteHttpExceptions: true
  };
}

function parseRssResponse(xml, query) {
  let newsList = [];
  if (!xml) return newsList;
  
  const oneWeekAgo = new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
  
  try {
    let itemMatches = xml.match(/<item>([\s\S]*?)<\/item>/g);
    if (itemMatches) {
      for (let i = 0; i < itemMatches.length; i++) {
        let itemXml = itemMatches[i];
        
        let pubDateStr = (itemXml.match(/<pubDate>([\s\S]*?)<\/pubDate>/) || ["", ""])[1];
        let rawDate = new Date(pubDateStr);
        
        if (rawDate >= oneWeekAgo) {
          let title = (itemXml.match(/<title>([\s\S]*?)<\/title>/) || ["", ""])[1]
                        .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, "$1")
                        .replace(/&lt;!\[CDATA\[([\s\S]*?)\]\]&gt;/g, "$1")
                        .replace(/<\/?[^>]+(>|$)/g, ""); // HTML 태그 제거
          
          let link = (itemXml.match(/<link>([\s\S]*?)<\/link>/) || ["", "#"])[1];
          
          newsList.push({ supplier: query, title: title, link: link, pubDate: pubDateStr, rawDate: rawDate });
        }
      }
    }
  } catch (error) {
    Logger.log("정규식 파싱 오류, XmlService로 폴백 시도: " + error.toString());
    try {
      let document = XmlService.parse(xml);
      let root = document.getRootElement();
      let channel = root.getChild('channel');
      if (channel) {
        let items = channel.getChildren('item');
        for (let i = 0; i < items.length; i++) {
          let pubDateStr = items[i].getChildText('pubDate') || new Date().toUTCString();
          let rawDate = new Date(pubDateStr);
          if (rawDate >= oneWeekAgo) {
            let title = items[i].getChildText('title') || '';
            let link = items[i].getChildText('link') || '#';
            newsList.push({ supplier: query, title: title, link: link, pubDate: pubDateStr, rawDate: rawDate });
          }
        }
      }
    } catch(xmlErr) {
      Logger.log("XmlService 폴백 실패: " + xmlErr.toString());
    }
  }
  
  newsList.sort((a, b) => b.rawDate - a.rawDate);
  return newsList.slice(0, 10);
}

function saveNewsToArchive(domestic, overseas, updateTimeStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("뉴스보관");
    
    if (!sheet) {
      sheet = ss.insertSheet("뉴스보관");
      sheet.appendRow(["구분", "협력사", "기사제목", "링크", "발행일", "업데이트일시"]);
      sheet.getRange("A1:F1").setBackground("#D3D3D3").setFontWeight("bold");
    }
    
    let dataArr = [];
    domestic.forEach(item => {
      dataArr.push(["국내", item.supplier, item.title, item.link, item.pubDate, updateTimeStr]);
    });
    overseas.forEach(item => {
      dataArr.push(["해외", item.supplier, item.title, item.link, item.pubDate, updateTimeStr]);
    });
    
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();
    }
    
    if (dataArr.length > 0) {
      sheet.getRange(2, 1, dataArr.length, 6).setValues(dataArr);
    }
  } catch(e) {
    Logger.log("[뉴스보관 오류] " + e.toString());
  }
}

function getArchivedNews() {
  let result = { domestic: [], overseas: [], lastUpdated: "" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("뉴스보관");
    if (!sheet) return result;
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return result;
    
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getDisplayValues(); 
    
    data.forEach(row => {
      const [type, supplier, title, link, pubDate, updateDate] = row;
      
      if (!result.lastUpdated && updateDate) result.lastUpdated = updateDate; 
      
      const item = { supplier: supplier, title: title, link: link, pubDate: pubDate };
      if (type === "국내") result.domestic.push(item);
      else if (type === "해외") result.overseas.push(item);
    });
  } catch(e) {
    Logger.log("[아카이브 로드 오류] " + e.toString());
  }
  return result;
}

function getSupplierNews(forceUpdate) { 
  try {
    const cache = CacheService.getScriptCache();
    if (!forceUpdate) {
      const cachedData = cache.get("supplier_news_data");
      if (cachedData) {
        return JSON.parse(cachedData);
      }
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("뉴스");
    if (!sheet) throw new Error("'뉴스' 시트를 찾을 수 없습니다.");
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { domestic: [], overseas: [], lastUpdated: "" };
    
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    let requests = [];
    let requestMeta = []; 
    
    for (let i = 0; i < data.length; i++) {
      let supplier = String(data[i][0]).trim();
      let regionType = String(data[i][1]).trim();
      if (!supplier) continue;
      
      let isDomestic = (regionType.indexOf('국내') !== -1);
      let isOverseas = (regionType.indexOf('해외') !== -1 || regionType.indexOf('국외') !== -1);
      
      if (isDomestic) {
        requests.push(buildRssRequest(supplier, true));
        requestMeta.push({ supplier: supplier, isDomestic: true });
      } else if (isOverseas) {
        requests.push(buildRssRequest(supplier, false));
        requestMeta.push({ supplier: supplier, isDomestic: false });
      }
    }
    
    if (requests.length === 0) return { domestic: [], overseas: [], lastUpdated: "" };
    
    let responses = UrlFetchApp.fetchAll(requests);
    let domestic = [];
    let overseas = [];
    
    for (let i = 0; i < responses.length; i++) {
      let res = responses[i];
      let meta = requestMeta[i];
      if (res.getResponseCode() === 200) {
        let xml = res.getContentText();
        let parsedNews = parseRssResponse(xml, meta.supplier);
        
        if (meta.isDomestic) {
          domestic = domestic.concat(parsedNews);
        } else {
          overseas = overseas.concat(parsedNews);
        }
      }
    }
    const sortByDate = (a, b) => new Date(b.pubDate) - new Date(a.pubDate);
    domestic.sort(sortByDate);
    overseas.sort(sortByDate);
    
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const nowStr = `${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
    
    const cleanDomestic = domestic.map(item => ({ supplier: item.supplier, title: item.title, link: item.link, pubDate: item.pubDate }));
    const cleanOverseas = overseas.map(item => ({ supplier: item.supplier, title: item.title, link: item.link, pubDate: item.pubDate }));
    
    let result = { domestic: cleanDomestic, overseas: cleanOverseas, lastUpdated: nowStr };
    let jsonString = JSON.stringify(result);
    
    saveNewsToArchive(domestic, overseas, nowStr);
    
    try {
      if (jsonString.length < 90000) {
        cache.put("supplier_news_data", jsonString, 1800); 
      }
    } catch(cacheError) {
      Logger.log("[뉴스] 캐시 에러 무시: " + cacheError.toString());
    }
    
    return result;
  } catch (e) {
    Logger.log("[뉴스 치명적 오류] " + e.toString());
    throw new Error(e.message || e.toString());
  }
}

function getExchangeRateData() {
  try {
    const ssId = '15mDVNS3jFIX4mNdu0OEHMTaHgDbwvNDSmBp7OxHsH9k';
    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName('CurrRaw');
    
    if (!sheet) {
      return { success: false, message: "'CurrRaw' 시트를 찾을 수 없습니다." };
    }
    const table1 = sheet.getRange('C5:G8').getDisplayValues();
    const table2 = sheet.getRange('I5:M8').getDisplayValues();
    return {
      success: true,
      table1: table1,
      table2: table2
    };
  } catch (e) {
    Logger.log("환율 데이터 로드 오류: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getRandomGreeting() {
  const greetings = [
    "행복 가득한 하루 되세요!",
    "오늘도 힘내세요!",
    "성공적인 하루를 응원합니다.",
    "오늘은 부디 칼퇴하세요!",
    "내일도 부디 칼퇴하세요!",
    "머리아픈 일은 다음에 하세요!",
    "오늘은 이쁜말만 하세요",
    "졸리면 어디가서 자고오세요",
    "주말에는 가족과 함께!"
  ];
  const randomIndex = Math.floor(Math.random() * greetings.length);
  return greetings[randomIndex];
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
      const inactiveFlag = data[i][4] ? String(data[i][4]).trim().toUpperCase() : "";
      const isInactiveFromSheet = (inactiveFlag === 'Y');
      
      if (!title || !url) continue;
      if (currentGroupNum !== groupNum) {
        if (currentGroup.length > 0) {
          menuGroups.push(currentGroup);
        }
        currentGroup = [];
        currentGroupNum = groupNum;
      }
      currentGroup.push({ title: title, url: url, icon: icon, isInactiveFromSheet: isInactiveFromSheet });
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
