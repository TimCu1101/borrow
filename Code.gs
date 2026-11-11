/****************************************************
 * 設備 QR Code 借還管理系統
 * Google Apps Script 後端
 *
 * Equipment_List：
 * A 設備編號 | B 設備名稱 | C 設備狀態
 * D QR Code | E 廠牌
 *
 * Borrow_Records：
 * A 設備編號 | B 設備名稱 | C 廠牌
 * D 借用人 | E 借用時間 | F 歸還人 | G 歸還時間
 ****************************************************/

const CONFIG = {
  EQUIPMENT_SHEET: "Equipment_List",
  RECORD_SHEET: "Borrow_Records",
  AVAILABLE: "可借用",
  BORROWED: "借出中",
  TIMEZONE: "Asia/Taipei",
  LOCK_TIMEOUT: 15000,
  HISTORY_LIMIT: 500
};

// GAS 網頁入口
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("設備 QR Code 借還管理系統")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// 取得目前綁定的 Google 試算表
function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      "無法取得試算表，請從 Google 試算表的「擴充功能 → Apps Script」建立專案。"
    );
  }

  return ss;
}

// 建立或取得設備工作表
function getEquipmentSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.EQUIPMENT_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.EQUIPMENT_SHEET);

    sheet.getRange(1, 1, 1, 5).setValues([[
      "設備編號",
      "設備名稱",
      "設備狀態",
      "QR Code",
      "廠牌"
    ]]);

    formatHeader_(sheet, 5);
  }

  return sheet;
}

// 建立或取得借還紀錄工作表
function getRecordSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.RECORD_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.RECORD_SHEET);
  }

  const headers = [
    "設備編號",
    "設備名稱",
    "廠牌",
    "借用人",
    "借用時間",
    "歸還人",
    "歸還時間"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]);

    formatHeader_(sheet, headers.length);
  }

  sheet.getRange("E:E")
    .setNumberFormat("yyyy/MM/dd HH:mm:ss");

  sheet.getRange("G:G")
    .setNumberFormat("yyyy/MM/dd HH:mm:ss");

  return sheet;
}

// 標題格式
function formatHeader_(sheet, columns) {
  sheet.getRange(1, 1, 1, columns)
    .setFontWeight("bold")
    .setBackground("#dbeafe")
    .setFontColor("#1e40af");

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, columns);
}

// 初始化系統：第一次執行
function setupSystem() {
  const ss = getSpreadsheet_();

  const equipmentSheet = getEquipmentSheet_(ss);
  getRecordSheet_(ss);

  equipmentSheet.getRange("A:A").setNumberFormat("@");
  equipmentSheet.getRange("D:D").setNumberFormat("@");

  // 沒有設備資料時建立範例
  if (equipmentSheet.getLastRow() < 2) {
    equipmentSheet.getRange(2, 1, 1, 5).setValues([[
      "EQ001",
      "範例設備",
      CONFIG.AVAILABLE,
      "EQ001",
      "範例廠牌"
    ]]);
  }

  SpreadsheetApp.flush();

  return "初始化完成，請至 Equipment_List 修改設備資料。";
}

// 輸入驗證
function validateInput_(id, name) {
  const cleanId = String(id || "").trim();
  const cleanName = String(name || "").trim();

  if (!cleanId) {
    throw new Error("請輸入或掃描設備編號。");
  }

  if (!cleanName) {
    throw new Error("請輸入操作人姓名。");
  }

  if (cleanName.length > 50) {
    throw new Error("姓名不可超過 50 個字。");
  }

  if (cleanId.length > 100) {
    throw new Error("設備編號長度不正確。");
  }

  return {
    id: cleanId,
    name: cleanName
  };
}

// 查找設備
function findEquipment_(sheet, id) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("設備清單目前沒有資料。");
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 5)
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === id) {
      return {
        row: i + 2,
        id: String(values[i][0]).trim(),
        equipmentName: values[i][1],
        status: values[i][2],
        qrCode: values[i][3],
        brand: values[i][4]
      };
    }
  }

  throw new Error("找不到設備編號：" + id);
}

// 前端初始資料
function getInitialData() {
  try {
    const ss = getSpreadsheet_();

    const equipmentSheet = getEquipmentSheet_(ss);
    const recordSheet = getRecordSheet_(ss);

    const equipment = [];

    if (equipmentSheet.getLastRow() >= 2) {
      const values = equipmentSheet
        .getRange(
          2,
          1,
          equipmentSheet.getLastRow() - 1,
          5
        )
        .getDisplayValues();

      values.forEach(row => {
        if (String(row[0]).trim()) {
          equipment.push({
            id: String(row[0]).trim(),
            equipmentName: row[1],
            status: row[2],
            qrCode: row[3],
            brand: row[4]
          });
        }
      });
    }

    return {
      success: true,
      equipment: equipment,
      records: readHistory_(recordSheet)
    };

  } catch (err) {
    return {
      success: false,
      message: err.message || "載入資料失敗。"
    };
  }
}

// 借用設備
function apiBorrowEquipment(id, name) {
  return processEquipment_("borrow", id, name);
}

// 歸還設備
function apiReturnEquipment(id, name) {
  return processEquipment_("return", id, name);
}

// 借用與歸還共用處理程序
function processEquipment_(action, id, name) {
  const input = validateInput_(id, name);
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT);

    const ss = getSpreadsheet_();

    const equipmentSheet = getEquipmentSheet_(ss);
    const recordSheet = getRecordSheet_(ss);

    const equipment = findEquipment_(
      equipmentSheet,
      input.id
    );

    const now = new Date();

    // 借用
    if (action === "borrow") {

      if (equipment.status !== CONFIG.AVAILABLE) {
        throw new Error(
          "此設備目前為「" +
          equipment.status +
          "」，無法借用。"
        );
      }

      let appendedRow = 0;

      try {
        recordSheet.appendRow([
          equipment.id,
          equipment.equipmentName,
          equipment.brand,
          input.name,
          now,
          "",
          ""
        ]);

        appendedRow = recordSheet.getLastRow();

        equipmentSheet
          .getRange(equipment.row, 3)
          .setValue(CONFIG.BORROWED);

      } catch (err) {

        if (appendedRow) {
          recordSheet.deleteRow(appendedRow);
        }

        throw err;
      }

    }

    // 歸還
    else if (action === "return") {

      if (equipment.status !== CONFIG.BORROWED) {
        throw new Error(
          "此設備目前不是借出狀態，無法歸還。"
        );
      }

      const lastRow = recordSheet.getLastRow();

      if (lastRow < 2) {
        throw new Error("找不到尚未歸還的借用紀錄。");
      }

      const rows = recordSheet
        .getRange(2, 1, lastRow - 1, 7)
        .getValues();

      let targetRow = -1;

      // 從最新紀錄往前尋找
      for (let i = rows.length - 1; i >= 0; i--) {
        const rowId = String(rows[i][0]).trim();
        const returnTime = rows[i][6];

        if (
          rowId === equipment.id &&
          (returnTime === "" || returnTime === null)
        ) {
          targetRow = i + 2;
          break;
        }
      }

      if (targetRow < 0) {
        throw new Error("找不到此設備尚未歸還的借用紀錄。");
      }

      const oldValues = recordSheet
        .getRange(targetRow, 6, 1, 2)
        .getValues();

      try {
        recordSheet
          .getRange(targetRow, 6, 1, 2)
          .setValues([[input.name, now]]);

        equipmentSheet
          .getRange(equipment.row, 3)
          .setValue(CONFIG.AVAILABLE);

      } catch (err) {

        recordSheet
          .getRange(targetRow, 6, 1, 2)
          .setValues(oldValues);

        throw err;
      }

    } else {
      throw new Error("不支援的操作類型。");
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      action: action,
      id: equipment.id,
      equipmentName: equipment.equipmentName,
      brand: equipment.brand,
      person: input.name,
      time: formatTime_(now),
      status:
        action === "borrow"
          ? CONFIG.BORROWED
          : CONFIG.AVAILABLE,
      message:
        action === "borrow"
          ? "設備借用成功！"
          : "設備歸還成功！"
    };

  } catch (err) {

    return {
      success: false,
      message: err.message || "操作失敗。"
    };

  } finally {

    if (lock.hasLock()) {
      lock.releaseLock();
    }

  }
}

// 查詢借還紀錄
function getBorrowHistory(filter) {
  try {
    const sheet = getRecordSheet_(getSpreadsheet_());

    let records = readHistory_(sheet);

    if (filter === "outstanding") {
      records = records.filter(
        record => !record.returnTime
      );
    }

    return {
      success: true,
      records: records
    };

  } catch (err) {

    return {
      success: false,
      message: err.message || "讀取紀錄失敗。",
      records: []
    };

  }
}

// 讀取最近紀錄，最新資料排前面
function readHistory_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const startRow = Math.max(
    2,
    lastRow - CONFIG.HISTORY_LIMIT + 1
  );

  const values = sheet
    .getRange(
      startRow,
      1,
      lastRow - startRow + 1,
      7
    )
    .getValues();

  return values.reverse().map(row => ({
    id: String(row[0] || ""),
    equipmentName: row[1] || "",
    brand: row[2] || "",
    borrower: row[3] || "",
    borrowTime: formatTime_(row[4]),
    returner: row[5] || "",
    returnTime: formatTime_(row[6])
  }));
}

// 台北時區格式化
function formatTime_(value) {
  if (!value) {
    return "";
  }

  return Utilities.formatDate(
    new Date(value),
    CONFIG.TIMEZONE,
    "yyyy/MM/dd HH:mm:ss"
  );
}

/**
 * 匯入 52 台設備並自動建立 QR Code
 * 執行一次即可；不會刪除既有借還紀錄。
 */
function importEquipmentList() {
  const ss = getSpreadsheet_();
  const sheet = getEquipmentSheet_(ss);
  getRecordSheet_(ss);

  const headers = ["設備編號", "設備名稱", "設備狀態", "QR Code", "廠牌"];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([headers]);
    formatHeader_(sheet, 5);
  }

  const devices = [];

  // 1001～1013：ASUS M640C
  for (let id = 1001; id <= 1013; id++) {
    devices.push([String(id), "M640C", "ASUS"]);
  }

  // 2023～2033：APPLE 平板電腦
  for (let id = 2023; id <= 2033; id++) {
    devices.push([String(id), "平板電腦", "APPLE"]);
  }

  // 3007～3012：ACER 筆記型電腦
  for (let id = 3007; id <= 3012; id++) {
    devices.push([String(id), "筆記型電腦", "ACER"]);
  }

  // 3013～3017：Apple 桌上型電腦
  for (let id = 3013; id <= 3017; id++) {
    devices.push([String(id), "桌上型電腦", "Apple"]);
  }

  // 2051～2067：mini
  for (let id = 2051; id <= 2067; id++) {
    devices.push([String(id), "mini", "mini"]);
  }

  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("D:D").setNumberFormat("@");

  // 建立既有設備編號對照表
  const existing = {};
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const oldData = sheet.getRange(2, 1, lastRow - 1, 5)
      .getDisplayValues();

    oldData.forEach((row, index) => {
      const id = String(row[0]).trim();
      if (id && !existing[id]) {
        existing[id] = index + 2;
      }
    });
  }

  let added = 0;
  let updated = 0;

  devices.forEach(device => {
    const id = device[0];
    const name = device[1];
    const brand = device[2];

    let row = existing[id];

    if (row) {
      // 更新名稱及廠牌，不覆蓋既有設備狀態
      sheet.getRange(row, 2).setValue(name);
      sheet.getRange(row, 5).setValue(brand);

      if (!sheet.getRange(row, 3).getDisplayValue()) {
        sheet.getRange(row, 3).setValue(CONFIG.AVAILABLE);
      }

      updated++;
    } else {
      // 新增設備
      sheet.appendRow([
        id,
        name,
        CONFIG.AVAILABLE,
        "",
        brand
      ]);

      row = sheet.getLastRow();
      added++;
      existing[id] = row;
    }

    // D 欄自動產生 QR Code 圖片
    sheet.getRange(row, 4).setFormula(
      '=IMAGE("https://quickchart.io/qr?text="&ENCODEURL(A' +
      row +
      ')&"&size=250&margin=2")'
    );
  });

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 110);

  SpreadsheetApp.flush();

  return "設備匯入完成！共處理 " + devices.length +
    " 台設備；新增 " + added + " 台，更新 " + updated + " 台。";
}
