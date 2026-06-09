const SHEETS = {
  settings: "活動設定",
  itinerary: "行程表",
  pickups: "上車地點",
  buses: "車次資料",
  registrations: "總表"
};

const DEFAULT_SETTINGS = [
  ["活動名稱", "2026 屏東西岐城、西方道堂之旅", "前台主標題"],
  ["活動日期", "2026.11.8 星期日", "前台顯示日期"],
  ["報名狀態", "開放", "填「開放」才可以送出；其他文字會關閉報名"],
  ["每台車保留名額", "2", "預設保留名額；當「車次資料」某車的保留名額欄空白時用此值（各車可在車次資料各自設定）"],
  ["前台說明", "請先選擇車次，再填寫姓名。可以一次報名多人，一行一個姓名。", "前台副標說明"],
  ["遊覽車大人費用", "1100", "依去年 PDF 金額先輸入"],
  ["遊覽車小孩費用", "900", "國一以下孩童"],
  ["自行開車大人費用", "800", "屏東地區自行開車同修"],
  ["自行開車小孩費用", "600", "屏東地區自行開車小孩"],
  ["費用包含", "費用包含午餐、水2瓶、下午茶、香火錢、車資。", "會同步到前台"],
  ["注意事項", "請著輕便服裝、運動鞋方便靈動參拜。", "會同步到前台"],
  ["西岐城地址", "屏東縣長治鄉新興路99巷73號", "資料備註"],
  ["西方道堂地址", "屏東縣萬巒鄉新厝村富山路1號", "資料備註"],
  ["午餐地址", "新和喜宴會館－屏東縣長治鄉繁昌村振興路16-11號（08-7623889）", "資料備註"]
];

const DEFAULT_ITINERARY = [
  [1, "08:00", "高雄道場＋高雄文化中心＋台南永康出發"],
  [2, "09:00", "屏東西岐城"],
  [3, "12:30", "午餐：新和喜宴會館"],
  [4, "14:00", "西方道堂"],
  [5, "17:30", "快樂返家"]
];

const DEFAULT_PICKUPS = [
  [1, "高雄文化中心", "早上8:00", "1.2.6.7.8.9車", "五福路大門口", 6, 43, "否", "是"],
  [2, "高雄道場三多", "早上8:00", "3.12車", "三多路全家便利商店前", 2, 43, "否", "是"],
  [3, "高雄道場85大樓", "早上8:00", "5.10車", "三多路85大樓門口", 2, 43, "否", "是"],
  [4, "台南永康", "早上8:00", "11車", "合眾汽車台南營業所（億峰家俱對面）", 1, 43, "否", "是"],
  [5, "屏東自行開車", "自行前往", "自行開車", "不占遊覽車座位", 0, 0, "是", "是"]
];

// 這張表的「列順序」就是車次顯示順序，可自行調整／增刪（自行開車也排在這裡）
// 欄位：車次, 車長, 副車長, 車號, 司機, 司機電話, 保留名額
const DEFAULT_BUSES = [
  ["1車", "", "", "", "", "", 2],
  ["2車", "", "", "", "", "", 2],
  ["3車", "", "", "", "", "", 5],
  ["自行開車", "", "", "", "", "", 0],
  ["5車", "", "", "", "", "", 2],
  ["6車", "", "", "", "", "", 2],
  ["7車", "", "", "", "", "", 2],
  ["8車", "", "", "", "", "", 2],
  ["9車", "", "", "", "", "", 2],
  ["10車", "", "", "", "", "", 2],
  ["11車", "", "", "", "", "", 2]
];

const REGISTRATION_HEADERS = [
  "時間",
  "姓名",
  "車次",
  "上車地點",
  "出發時間",
  "身分"
];

function doGet(e) {
  const params = e.parameter || {};
  const action = params.action || "status";
  const callback = sanitizeCallback_(params.callback);

  let result;

  try {
    setupSheets_();

    if (action === "submit") {
      result = submitRegistration_(params);
    } else if (action === "query") {
      result = queryByName_(params);
    } else if (action === "roster") {
      result = getRoster_();
    } else {
      result = getStatus_();
    }
  } catch (error) {
    result = {
      status: "error",
      message: error.message || "系統錯誤"
    };
  }

  return output_(result, callback);
}

function doPost(e) {
  let params = {};

  try {
    params = JSON.parse(e.postData.contents || "{}");
  } catch (error) {
    params = {};
  }

  setupSheets_();
  const result = submitRegistration_(params);

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 第一次可手動執行這個，會自動建立後台工作表。
function setupSheet() {
  setupSheets_();
}

// 開啟試算表時加上自訂選單
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("報名系統")
    .addItem("產生各車次名單", "generateBusListSheet")
    .addItem("產生統計表", "generateStatsSheet")
    .addItem("產生車次總覽", "generateBusSummarySheet")
    .addItem("重建後台工作表", "setupSheet")
    .addToUi();
}

// 依「總表」報名，套上車次資料，產生「各車次」工作表（名單 4 人一排）
function generateBusListSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets_();

  const regSheet = getRegistrationSheet_();
  const lastRow = regSheet.getLastRow();
  const regs = lastRow > 1
    ? regSheet.getRange(2, 1, lastRow - 1, REGISTRATION_HEADERS.length).getValues()
    : [];

  const busMap = {};
  const busOrderList = [];
  getBusInfo_().forEach(function (b) {
    busMap[b.busNo] = b;
    busOrderList.push(b.busNo);
  });

  const order = [];
  const groups = {};

  regs.forEach(function (row) {
    const name = String(row[1] || "").trim();
    if (!name) return;
    const busNo = String(row[2] || "").trim() || "未分配車次";
    if (!groups[busNo]) {
      groups[busNo] = [];
      order.push(busNo);
    }
    groups[busNo].push({
      name: name,
      pickup: String(row[3] || "").trim(),
      departureTime: String(row[4] || "").trim(),
      identity: String(row[5] || "").trim()
    });
  });

  order.sort(function (a, b) {
    const ia = busOrderList.indexOf(a);
    const ib = busOrderList.indexOf(b);
    const ka = ia < 0 ? 1000 + busSortKey_(a) : ia;
    const kb = ib < 0 ? 1000 + busSortKey_(b) : ib;
    return ka - kb;
  });

  const out = [];
  const titleRows = [];
  const headerRows = [];
  out.push(["各車次名單（系統自動產生，每次點選單會重新產生）", "", "", "", ""]);
  out.push(["", "", "", "", ""]);

  order.forEach(function (busNo) {
    const people = groups[busNo];
    const isSelf = busNo.indexOf("自行") >= 0;
    const info = busMap[busNo] || {};
    const p0 = people[0];

    titleRows.push(out.length);
    if (isSelf) {
      out.push([busNo, "", "", "", ""]);
      out.push(["出發時間與地點：自行前往", "", "", "", ""]);
    } else {
      out.push([busNo + "　車長：" + (info.captain || "") + "　副車長：" + (info.viceCaptain || ""), "", "", "", ""]);
      out.push(["出發時間與地點：" + (p0.departureTime || "") + "　" + p0.pickup, "", "", "", ""]);
      out.push(["車號：" + (info.plate || "") + "　" + (info.driver || "") + "　" + (info.driverPhone || ""), "", "", "", ""]);
    }

    headerRows.push(out.length);
    out.push(["", "姓名", "姓名", "姓名", "姓名"]);

    for (var i = 0; i < people.length; i += 4) {
      const row = [(i / 4) + 1, "", "", "", ""];
      for (var j = 0; j < 4; j++) {
        const p = people[i + j];
        if (p) row[j + 1] = p.name + (p.identity === "小孩" ? "（小孩）" : "");
      }
      out.push(row);
    }

    out.push(["", "", "", "", ""]);
  });

  let sheet = ss.getSheetByName("各車次");
  if (!sheet) {
    sheet = ss.insertSheet("各車次");
  }
  sheet.clear();

  if (out.length) {
    sheet.getRange(1, 1, out.length, 5).setValues(out);
    sheet.getRange(1, 1, out.length, 5).setBorder(true, true, true, true, true, true, "#c9d4da", null);
  }

  sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setFontSize(14);
  titleRows.forEach(function (rowIndex) {
    sheet.getRange(rowIndex + 1, 1, 1, 5)
      .setFontWeight("bold")
      .setBackground("#0f80ad")
      .setFontColor("#ffffff");
  });
  headerRows.forEach(function (rowIndex) {
    sheet.getRange(rowIndex + 1, 1, 1, 5)
      .setFontWeight("bold")
      .setBackground("#eef3f6")
      .setHorizontalAlignment("center");
  });
  sheet.setColumnWidth(1, 48);
  sheet.setColumnWidths(2, 4, 150);
  sheet.setFrozenRows(1);

  generateStatsSheet();
  generateBusSummarySheet();

  SpreadsheetApp.getActive().toast("各車次名單、統計表、車次總覽已產生／更新", "完成", 5);
}

// 產生／更新「車次總覽」：車長副車長＋搭車人數(公式)＋實到人數(手填)＋用餐桌次(每車4桌)＋備註(手填)＋合計
function generateBusSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regName = SHEETS.registrations;
  const buses = getBusInfo_();
  const SHEET_NAME = "車次總覽";

  let sheet = ss.getSheetByName(SHEET_NAME);

  // 保留舊的手填欄（實到人數 C、備註 E），以隱藏的 busNo（H 欄）當鍵
  const prev = {};
  if (sheet) {
    const lr = sheet.getLastRow();
    if (lr >= 1) {
      sheet.getRange(1, 1, lr, 8).getValues().forEach(function (row) {
        const key = String(row[7] || "").trim();
        if (key) prev[key] = { actual: row[2], note: row[4] };
      });
    }
    sheet.clear();
  } else {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const out = [];
  out.push(["車長／副車長", "搭車人數", "實到人數", "用餐桌次", "備註（多出來的排法請自行填寫）", "", "", ""]);

  let tableNo = 1;
  buses.forEach(function (b) {
    const isSelf = b.busNo.indexOf("自行") >= 0;
    const label = isSelf
      ? b.busNo + (b.captain ? "　" + b.captain : "")
      : b.busNo + "　車長：" + (b.captain || "") + "　副車長：" + (b.viceCaptain || "");

    const tables = [];
    for (var t = 0; t < 4; t++) { tables.push(tableNo); tableNo++; }

    const restored = prev[b.busNo] || {};

    out.push([
      label,
      "=COUNTIF('" + regName + "'!$C:$C,\"" + b.busNo + "\")",
      restored.actual != null ? restored.actual : "",
      tables.join("."),
      restored.note != null ? restored.note : "",
      "", "",
      b.busNo
    ]);
  });

  out.push(["", "", "", "", "", "", "", ""]);
  out.push(["司機（請自行填寫）", "", "", "", "", "", "", ""]);

  const sumEnd = buses.length + 3; // 含上方資料列、空列、司機列
  out.push([
    "合計",
    "=SUM(B2:B" + sumEnd + ")",
    "=SUM(C2:C" + sumEnd + ")",
    "", "", "", "", ""
  ]);

  sheet.getRange(1, 1, out.length, 8).setValues(out);

  const totalRow = out.length;
  sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#0f80ad").setFontColor("#ffffff");
  sheet.getRange(totalRow, 1, 1, 3).setFontWeight("bold").setBackground("#fff4c2");
  if (buses.length > 0) {
    sheet.getRange(2, 3, buses.length, 1).setBackground("#fff3cd"); // 實到人數手填
    sheet.getRange(2, 5, buses.length, 1).setBackground("#fff3cd"); // 備註手填
  }
  sheet.getRange(1, 1, out.length, 5).setBorder(true, true, true, true, true, true, "#c9d4da", null);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidths(2, 3, 90);
  sheet.setColumnWidth(5, 260);
  sheet.hideColumns(8);
  sheet.setFrozenRows(1);
}

// 產生／更新「統計」工作表：各車人數＋區域小計＋台中手填＋總人數（皆即時公式）
function generateStatsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regName = SHEETS.registrations;
  const buses = getBusInfo_().map(function (b) { return b.busNo; });

  let sheet = ss.getSheetByName("統計");
  let prevTaichung1 = "";
  let prevTaichung2 = "";

  if (sheet) {
    prevTaichung1 = sheet.getRange("E5").getValue();
    prevTaichung2 = sheet.getRange("E6").getValue();
    sheet.clear();
  } else {
    sheet = ss.insertSheet("統計");
  }

  // 左區塊：各車人數（即時公式）
  const left = [["各車人數", "人數"]];
  buses.forEach(function (bus) {
    const rowNum = left.length + 1;
    left.push([bus, "=COUNTIF('" + regName + "'!$C:$C,$A" + rowNum + ")"]);
  });
  sheet.getRange(1, 1, left.length, 2).setValues(left);

  // 右區塊：區域統計（固定位置，台中手填於 E5/E6）
  const right = [
    ["區域統計", "人數"],
    ["高雄", "=COUNTIF('" + regName + "'!$D:$D,\"高雄*\")"],
    ["台南", "=COUNTIF('" + regName + "'!$D:$D,\"台南*\")"],
    ["屏東", "=COUNTIF('" + regName + "'!$D:$D,\"屏東*\")"],
    ["台中（第一台）", prevTaichung1],
    ["台中（第二台）", prevTaichung2],
    ["", ""],
    ["總人數", "=E2+E3+E4+E5+E6"]
  ];
  sheet.getRange(1, 4, right.length, 2).setValues(right);

  // 樣式
  sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#0f80ad").setFontColor("#ffffff");
  sheet.getRange("D1:E1").setFontWeight("bold").setBackground("#0f80ad").setFontColor("#ffffff");
  sheet.getRange("D5:E6").setBackground("#fff3cd"); // 台中手填區
  sheet.getRange("D8:E8").setFontWeight("bold").setBackground("#fff4c2");
  sheet.setColumnWidths(1, 5, 130);
}

function busSortKey_(busNo) {
  const text = String(busNo || "");
  if (text.indexOf("自行") >= 0) return 9999;
  const match = text.match(/\d+/);
  return match ? Number(match[0]) : 9998;
}

function submitRegistration_(params) {
  const settings = getSettings_();

  if (String(settings.registrationStatus || "").trim() !== "開放") {
    return {
      status: "closed",
      message: "目前報名尚未開放或已關閉"
    };
  }

  const people = parsePeople_(params);

  if (people.length === 0) {
    return {
      status: "error",
      message: "請至少輸入一個姓名"
    };
  }

  const busNo = String(params.bus || "").trim();
  const option = getBusOptions_().find(function (o) {
    return o.busNo === busNo;
  });

  if (!busNo || !option) {
    return {
      status: "error",
      message: "請選擇正確的車次（可能該車尚未指定車長或已關閉）"
    };
  }

  const reserve = Number(option.reserve || 0);
  const names = people.map(function (p) { return p.name; });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getRegistrationSheet_();
    const used = countByBus_(sheet, busNo);
    const requested = people.length;
    const duplicates = findDuplicateNames_(sheet, names);

    if (!option.selfDrive) {
      const openCap = Math.max(option.capacity - reserve, 0);
      const remaining = openCap - used;

      if (requested > remaining) {
        return {
          status: "full",
          busNo: busNo,
          pickup: option.pickup,
          used: used,
          capacity: openCap,
          remaining: Math.max(remaining, 0),
          requested: requested
        };
      }
    }

    const now = new Date();
    const rows = people.map(function (p) {
      return [now, p.name, busNo, option.pickup, option.departureTime, p.identity];
    });
    const records = people.map(function (p) {
      return {
        name: p.name,
        identity: p.identity,
        pickup: option.pickup,
        departureTime: option.departureTime,
        location: option.location,
        busNo: busNo,
        selfDrive: option.selfDrive
      };
    });

    sheet
      .getRange(sheet.getLastRow() + 1, 1, rows.length, REGISTRATION_HEADERS.length)
      .setValues(rows);

    return {
      status: "success",
      eventName: settings.title,
      pickup: option.pickup,
      busNo: busNo,
      count: people.length,
      records: records,
      duplicates: duplicates,
      used: used + people.length,
      remaining: option.selfDrive ? "不限" : Math.max(option.capacity - reserve - used - people.length, 0)
    };
  } finally {
    lock.releaseLock();
  }
}

// 解析前台送來的報名者（含大人/小孩身分）
function parsePeople_(params) {
  const people = [];

  if (params.people) {
    try {
      JSON.parse(params.people).forEach(function (p) {
        const name = String((p && p.name) || "").trim();
        if (name) {
          people.push({ name: name, identity: String(p.identity || "").trim() === "小孩" ? "小孩" : "大人" });
        }
      });
    } catch (e) {
      // 解析失敗就走下面相容模式
    }
  }

  if (people.length === 0) {
    parseNames_(params.adults || params.names || params.name || "").forEach(function (n) {
      people.push({ name: n, identity: "大人" });
    });
    parseNames_(params.children || "").forEach(function (n) {
      people.push({ name: n, identity: "小孩" });
    });
  }

  return people;
}

// 計算某車次已報名人數
function countByBus_(sheet, busNo) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return 0;
  }
  return sheet.getRange(2, 1, lastRow - 1, REGISTRATION_HEADERS.length).getValues().filter(function (row) {
    return String(row[2] || "").trim() === busNo;
  }).length;
}

// 計算所有車次的人數，回傳 {busNo: 人數}
function countAllByBus_(sheet) {
  const map = {};
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return map;
  }
  sheet.getRange(2, 1, lastRow - 1, REGISTRATION_HEADERS.length).getValues().forEach(function (row) {
    const bus = String(row[2] || "").trim();
    if (bus) map[bus] = (map[bus] || 0) + 1;
  });
  return map;
}

// 產生登記頁的車次選項：每台車 = 上車地點 + 車號 + 車長，只列「已填車長」的車
function getBusOptions_() {
  const pickups = getPickupConfigs_();
  const busInfo = {};
  const busOrderList = [];
  getBusInfo_().forEach(function (b) {
    busInfo[b.busNo] = b;
    busOrderList.push(b.busNo);
  });

  const settings = getSettings_();
  const defaultReserve = Number(settings.reservePerBus || 2);
  const counts = countAllByBus_(getRegistrationSheet_());

  const options = [];

  pickups.forEach(function (p) {
    if (p.selfDrive) {
      options.push({
        busNo: "自行開車",
        captain: "",
        viceCaptain: "",
        pickup: p.name,
        departureTime: p.departureTime,
        location: p.location,
        selfDrive: true,
        used: counts["自行開車"] || 0,
        remaining: "不限"
      });
      return;
    }

    parseBusNumbers_(p.busLabel).forEach(function (n) {
      const busNo = n + "車";
      const info = busInfo[busNo] || {};
      if (!info.captain) return; // 只列已指定車長的車次

      const used = counts[busNo] || 0;
      const reserve = (info.reserve != null && !isNaN(info.reserve)) ? info.reserve : defaultReserve;
      const openCap = Math.max(p.seatsPerBus - reserve, 0);

      options.push({
        busNo: busNo,
        captain: info.captain,
        viceCaptain: info.viceCaptain || "",
        pickup: p.name,
        departureTime: p.departureTime,
        location: p.location,
        selfDrive: false,
        capacity: p.seatsPerBus,
        reserve: reserve,
        used: used,
        remaining: Math.max(openCap - used, 0)
      });
    });
  });

  options.sort(function (a, b) {
    const ia = busOrderList.indexOf(a.busNo);
    const ib = busOrderList.indexOf(b.busNo);
    const ka = ia < 0 ? 1000 + busSortKey_(a.busNo) : ia;
    const kb = ib < 0 ? 1000 + busSortKey_(b.busNo) : ib;
    return ka - kb;
  });

  return options;
}

function getStatus_() {
  const settings = getSettings_();
  const itinerary = getItinerary_();
  const pickups = getPickupConfigs_();
  const sheet = getRegistrationSheet_();

  const pickupStatus = pickups.map(function (pickup) {
    const used = countActiveByPickup_(sheet, pickup.name);

    return {
      name: pickup.name,
      departureTime: pickup.departureTime,
      busLabel: pickup.busLabel,
      location: pickup.location,
      buses: pickup.buses,
      seatsPerBus: pickup.seatsPerBus,
      used: used,
      remaining: pickup.selfDrive ? "不限" : Math.max(pickup.capacity - used, 0),
      capacity: pickup.selfDrive ? "不限" : pickup.capacity,
      selfDrive: pickup.selfDrive
    };
  });

  return {
    status: "success",
    activity: settings,
    itinerary: itinerary,
    pickups: pickupStatus,
    busOptions: getBusOptions_()
  };
}

// 依姓名查詢個人安排（車次、上車地點、同行名單）
function queryByName_(params) {
  const name = String(params.name || "").trim();

  if (!name) {
    return { status: "error", message: "請輸入姓名" };
  }

  const sheet = getRegistrationSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return { status: "success", name: name, matches: [] };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, REGISTRATION_HEADERS.length).getValues();

  const matches = values
    .filter(function (row) {
      return String(row[1] || "").trim() === name;
    })
    .map(function (row) {
      return {
        name: String(row[1] || "").trim(),
        busNo: String(row[2] || "").trim(),
        pickup: String(row[3] || "").trim(),
        departureTime: String(row[4] || "").trim(),
        identity: String(row[5] || "").trim()
      };
    });

  return { status: "success", name: name, matches: matches };
}

// 全部名單（給前台依車次排序顯示）
function getRoster_() {
  const sheet = getRegistrationSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return { status: "success", roster: [] };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, REGISTRATION_HEADERS.length).getValues();

  const roster = values
    .filter(function (row) {
      return String(row[1] || "").trim();
    })
    .map(function (row) {
      return {
        name: String(row[1] || "").trim(),
        busNo: String(row[2] || "").trim(),
        pickup: String(row[3] || "").trim(),
        identity: String(row[5] || "").trim()
      };
    });

  return { status: "success", roster: roster, buses: getBusInfo_() };
}

function setupSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  setupSettingsSheet_(ss);
  setupItinerarySheet_(ss);
  setupPickupsSheet_(ss);
  setupBusesSheet_(ss);
  getRegistrationSheet_();
}

function setupSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEETS.settings);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.settings);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([["設定項目", "設定內容", "備註"]]);
    sheet.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    formatHeader_(sheet, 3);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 3);
  }
}

function setupItinerarySheet_(ss) {
  let sheet = ss.getSheetByName(SHEETS.itinerary);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.itinerary);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([["排序", "時間", "行程內容"]]);
    sheet.getRange(2, 1, DEFAULT_ITINERARY.length, 3).setValues(DEFAULT_ITINERARY);
    formatHeader_(sheet, 3);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 3);
  }
}

function setupPickupsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEETS.pickups);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.pickups);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 9).setValues([[
      "排序",
      "上車地點",
      "出發時間",
      "車次",
      "集合地點",
      "台數",
      "每台座位",
      "是否自行開車",
      "是否啟用"
    ]]);
    sheet.getRange(2, 1, DEFAULT_PICKUPS.length, 9).setValues(DEFAULT_PICKUPS);
    formatHeader_(sheet, 9);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 9);
  }
}

function setupBusesSheet_(ss) {
  let sheet = ss.getSheetByName(SHEETS.buses);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.buses);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 7).setValues([["車次", "車長", "副車長", "車號", "司機", "司機電話", "保留名額"]]);
    sheet.getRange(2, 1, DEFAULT_BUSES.length, 7).setValues(DEFAULT_BUSES);
    formatHeader_(sheet, 7);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 7);
  }
}

function getBusInfo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.buses);

  if (!sheet) {
    return [];
  }

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  return values
    .filter(function (row) {
      return String(row[0] || "").trim();
    })
    .map(function (row) {
      const reserveRaw = row[6];
      return {
        busNo: String(row[0] || "").trim(),
        captain: String(row[1] || "").trim(),
        viceCaptain: String(row[2] || "").trim(),
        plate: String(row[3] || "").trim(),
        driver: String(row[4] || "").trim(),
        driverPhone: String(row[5] || "").trim(),
        reserve: (reserveRaw === "" || reserveRaw === null || reserveRaw === undefined) ? null : Number(reserveRaw)
      };
    });
}

function getRegistrationSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.registrations);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.registrations);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, REGISTRATION_HEADERS.length).setValues([REGISTRATION_HEADERS]);
    formatHeader_(sheet, REGISTRATION_HEADERS.length);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, REGISTRATION_HEADERS.length);
    applyDuplicateHighlight_(sheet);
  }

  return sheet;
}

// 「姓名」欄位若出現重複，整格上色（色塊區分），會隨資料自動更新
function applyDuplicateHighlight_(sheet) {
  const range = sheet.getRange("B2:B1000");
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B2<>"",COUNTIF($B$2:$B$1000,$B2)>1)')
    .setBackground("#ffd5da")
    .setFontColor("#9b2c2c")
    .setRanges([range])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

// 找出本次報名中、姓名已存在於總表的（重複登記）
function findDuplicateNames_(sheet, names) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const existing = {};
  sheet.getRange(2, 1, lastRow - 1, REGISTRATION_HEADERS.length).getValues().forEach(function (row) {
    const name = String(row[1] || "").trim();
    if (name) existing[name] = true;
  });

  return names.filter(function (name) {
    return existing[name];
  });
}

function formatHeader_(sheet, columnCount) {
  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight("bold")
    .setBackground("#0f80ad")
    .setFontColor("#ffffff");
}

function getSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.settings);
  const values = sheet.getDataRange().getValues();
  const map = {};

  values.slice(1).forEach(function (row) {
    const key = String(row[0] || "").trim();
    const value = String(row[1] || "").trim();
    if (key) map[key] = value;
  });

  return {
    title: map["活動名稱"] || "2026 屏東西岐城、西方道堂之旅",
    date: map["活動日期"] || "2026.11.8 星期日",
    subtitle: map["前台說明"] || "請先選擇車次，再填寫姓名。可以一次報名多人，一行一個姓名。",
    registrationStatus: map["報名狀態"] || "開放",
    reservePerBus: map["每台車保留名額"] || "5",
    busAdultFee: map["遊覽車大人費用"] || "1100",
    busChildFee: map["遊覽車小孩費用"] || "900",
    selfDriveAdultFee: map["自行開車大人費用"] || "800",
    selfDriveChildFee: map["自行開車小孩費用"] || "600",
    feeIncludes: map["費用包含"] || "費用包含午餐、水2瓶、下午茶、香火錢、車資。",
    notes: map["注意事項"] || "請著輕便服裝、運動鞋方便靈動參拜。",
    xiqichengAddress: map["西岐城地址"] || "",
    xifangAddress: map["西方道堂地址"] || "",
    lunchAddress: map["午餐地址"] || ""
  };
}

function getItinerary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.itinerary);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  return values
    .filter(function (row) {
      return String(row[1] || "").trim() || String(row[2] || "").trim();
    })
    .sort(function (a, b) {
      return Number(a[0] || 999) - Number(b[0] || 999);
    })
    .map(function (row) {
      return {
        time: String(row[1] || "").trim(),
        content: String(row[2] || "").trim()
      };
    });
}

function getPickupConfigs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.pickups);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  return values
    .filter(function (row) {
      return isYes_(row[8]) && String(row[1] || "").trim();
    })
    .sort(function (a, b) {
      return Number(a[0] || 999) - Number(b[0] || 999);
    })
    .map(function (row) {
      const selfDrive = isYes_(row[7]);
      const buses = Number(row[5] || 0);
      const seatsPerBus = Number(row[6] || 0);
      const capacity = selfDrive ? null : buses * seatsPerBus;

      return {
        name: String(row[1] || "").trim(),
        departureTime: String(row[2] || "").trim(),
        busLabel: String(row[3] || "").trim(),
        location: String(row[4] || "").trim(),
        buses: buses,
        seatsPerBus: seatsPerBus,
        selfDrive: selfDrive,
        capacity: capacity
      };
    });
}

function countActiveByPickup_(sheet, pickup) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return 0;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, REGISTRATION_HEADERS.length).getValues();

  return values.filter(function (row) {
    return String(row[3] || "").trim() === pickup;
  }).length;
}

function parseNames_(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map(function (name) {
      return name.trim();
    })
    .filter(function (name) {
      return name.length > 0;
    });
}

function parseBusNumbers_(busLabel) {
  const matches = String(busLabel || "").match(/\d+/g);
  return matches || [];
}

function formatBusNo_(busNumber, fallbackIndex) {
  if (busNumber) {
    return busNumber + "車";
  }

  return "第" + fallbackIndex + "台車";
}

function makeRegistrationNo_(pickup, number) {
  const codeMap = {
    "高雄文化中心": "WENHUA",
    "高雄道場三多": "SANDUO",
    "高雄道場85大樓": "85",
    "台南永康": "TAINAN",
    "屏東自行開車": "PINGTUNG"
  };

  const code = codeMap[pickup] || "SIGNUP";
  return code + "-" + String(number).padStart(3, "0");
}

function makeBatchNo_() {
  return "BATCH-" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 1000);
}

function isYes_(value) {
  const text = String(value || "").trim().toLowerCase();
  return ["是", "yes", "y", "true", "1", "啟用"].indexOf(text) >= 0;
}

function output_(result, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(result) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function sanitizeCallback_(callback) {
  const text = String(callback || "").trim();

  if (/^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(text)) {
    return text;
  }

  return "";
}
