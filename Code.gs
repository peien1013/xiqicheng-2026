const SHEETS = {
  settings: "活動設定",
  itinerary: "行程表",
  pickups: "上車地點",
  buses: "車次資料",
  registrations: "報名資料"
};

const DEFAULT_SETTINGS = [
  ["活動名稱", "2026 屏東西岐城、西方道堂之旅", "前台主標題"],
  ["活動日期", "2026.11.8 星期日", "前台顯示日期"],
  ["報名狀態", "開放", "填「開放」才可以送出；其他文字會關閉報名"],
  ["前台說明", "請輸入姓名並選擇上車地點。可以一次報名多人，一行一個姓名。", "前台副標說明"],
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

const DEFAULT_BUSES = [
  ["1車", "", "", "", "", ""],
  ["2車", "", "", "", "", ""],
  ["3車", "", "", "", "", ""],
  ["4車", "", "", "", "", ""],
  ["5車", "", "", "", "", ""],
  ["6車", "", "", "", "", ""],
  ["7車", "", "", "", "", ""],
  ["8車", "", "", "", "", ""],
  ["9車", "", "", "", "", ""],
  ["10車", "", "", "", "", ""],
  ["11車", "", "", "", "", ""],
  ["12車", "", "", "", "", ""],
  ["13車", "", "", "", "", ""]
];

const REGISTRATION_HEADERS = [
  "時間",
  "批次編號",
  "活動名稱",
  "姓名",
  "上車地點",
  "出發時間",
  "集合地點",
  "車次",
  "座位序號",
  "報名編號",
  "狀態"
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

function submitRegistration_(params) {
  const settings = getSettings_();

  if (String(settings.registrationStatus || "").trim() !== "開放") {
    return {
      status: "closed",
      message: "目前報名尚未開放或已關閉"
    };
  }

  const names = parseNames_(params.names || params.name || "");
  const pickupName = String(params.pickup || "").trim();
  const pickupConfigs = getPickupConfigs_();
  const pickup = pickupConfigs.find(function (item) {
    return item.name === pickupName;
  });

  if (names.length === 0) {
    return {
      status: "error",
      message: "請至少輸入一個姓名"
    };
  }

  if (!pickup) {
    return {
      status: "error",
      message: "請選擇正確的上車地點"
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getRegistrationSheet_();
    const used = countActiveByPickup_(sheet, pickup.name);
    const requested = names.length;

    if (!pickup.selfDrive) {
      const remaining = pickup.capacity - used;

      if (requested > remaining) {
        return {
          status: "full",
          pickup: pickup.name,
          used: used,
          capacity: pickup.capacity,
          remaining: Math.max(remaining, 0),
          requested: requested
        };
      }
    }

    const batchNo = makeBatchNo_();
    const now = new Date();
    const rows = [];
    const records = [];
    const busNumbers = parseBusNumbers_(pickup.busLabel);

    names.forEach(function (name, index) {
      const nextNumber = used + index + 1;
      const busIndex = pickup.selfDrive ? -1 : Math.floor((nextNumber - 1) / pickup.seatsPerBus);
      const seatNo = pickup.selfDrive ? "自行開車" : ((nextNumber - 1) % pickup.seatsPerBus) + 1;
      const busNo = pickup.selfDrive ? "自行開車" : formatBusNo_(busNumbers[busIndex], busIndex + 1);
      const registrationNo = makeRegistrationNo_(pickup.name, nextNumber);

      rows.push([
        now,
        batchNo,
        settings.title,
        name,
        pickup.name,
        pickup.departureTime,
        pickup.location,
        busNo,
        seatNo,
        registrationNo,
        "有效"
      ]);

      records.push({
        name: name,
        pickup: pickup.name,
        departureTime: pickup.departureTime,
        location: pickup.location,
        busNo: busNo,
        seatNo: seatNo,
        registrationNo: registrationNo,
        selfDrive: pickup.selfDrive
      });
    });

    sheet
      .getRange(sheet.getLastRow() + 1, 1, rows.length, REGISTRATION_HEADERS.length)
      .setValues(rows);

    return {
      status: "success",
      eventName: settings.title,
      pickup: pickup.name,
      count: names.length,
      batchNo: batchNo,
      records: records,
      used: used + names.length,
      capacity: pickup.selfDrive ? "不限" : pickup.capacity,
      remaining: pickup.selfDrive ? "不限" : pickup.capacity - used - names.length
    };
  } finally {
    lock.releaseLock();
  }
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
    pickups: pickupStatus
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
  const active = values.filter(function (row) {
    return String(row[10] || "有效").trim() !== "取消" && String(row[3] || "").trim();
  });

  const matches = active
    .filter(function (row) {
      return String(row[3] || "").trim() === name;
    })
    .map(function (row) {
      const batchNo = String(row[1] || "").trim();
      const companions = active
        .filter(function (r) {
          return String(r[1] || "").trim() === batchNo && String(r[3] || "").trim() !== name;
        })
        .map(function (r) {
          return String(r[3] || "").trim();
        });

      return {
        name: String(row[3] || "").trim(),
        pickup: String(row[4] || "").trim(),
        departureTime: String(row[5] || "").trim(),
        location: String(row[6] || "").trim(),
        busNo: String(row[7] || "").trim(),
        seatNo: row[8],
        registrationNo: String(row[9] || "").trim(),
        companions: companions
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
      return String(row[10] || "有效").trim() !== "取消" && String(row[3] || "").trim();
    })
    .map(function (row) {
      return {
        name: String(row[3] || "").trim(),
        pickup: String(row[4] || "").trim(),
        busNo: String(row[7] || "").trim(),
        seatNo: row[8]
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
    sheet.getRange(1, 1, 1, 6).setValues([["車次", "車長", "副車長", "車號", "司機", "司機電話"]]);
    sheet.getRange(2, 1, DEFAULT_BUSES.length, 6).setValues(DEFAULT_BUSES);
    formatHeader_(sheet, 6);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 6);
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

  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  return values
    .filter(function (row) {
      return String(row[0] || "").trim();
    })
    .map(function (row) {
      return {
        busNo: String(row[0] || "").trim(),
        captain: String(row[1] || "").trim(),
        viceCaptain: String(row[2] || "").trim(),
        plate: String(row[3] || "").trim(),
        driver: String(row[4] || "").trim(),
        driverPhone: String(row[5] || "").trim()
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
  }

  return sheet;
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
    subtitle: map["前台說明"] || "請輸入姓名並選擇上車地點。可以一次報名多人，一行一個姓名。",
    registrationStatus: map["報名狀態"] || "開放",
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
    const rowPickup = String(row[4] || "").trim();
    const status = String(row[10] || "有效").trim();
    return rowPickup === pickup && status !== "取消";
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
