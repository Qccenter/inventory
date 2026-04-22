var SHEET_MASTER = "Chemical_Master";
var SHEET_STOCK = "Stock_Inventory";
var SHEET_LOG = "Transaction_Log";
var SHEET_USERS = "User_Master";
var SHEET_CONFIG = "System_Config";
var SHEET_AUDIT = "Audit_Log";
var SESSION_TTL_SECONDS = 21600;

function doGet(e) {
  var ss = getSS_();
  initializeSchema_(ss);

  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest_(e);
  }

  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta charset="utf-8"><title>Chemical API</title></head><body>' +
    '<h2>Chemical Control API พร้อมใช้งาน</h2>' +
    '<p>สคริปต์นี้ใช้ Google Sheet ที่ bind อยู่กับโปรเจกต์นี้อัตโนมัติ</p>' +
    '<p>ให้นำ URL ของ Web App นี้ไปใส่ในหน้า Settings ของเว็บบน GitHub Pages</p>' +
    '<p>ตัวอย่าง: <code>?action=system-status</code></p>' +
    '</body></html>'
  ).setTitle("Chemical Control API");
}

function doPost(e) {
  var ss = getSS_();
  initializeSchema_(ss);
  return handleApiRequest_(e);
}

function handleApiRequest_(e) {
  try {
    var params = getParams_(e);
    var action = (params.action || "").toLowerCase();
    var response;

    switch (action) {
      case "login":
        response = apiLogin_(params);
        break;
      case "register":
        response = apiRegisterUser_(params);
        break;
      case "session":
        response = withSession_(params, function(session) {
          return success_({ user: session.user });
        });
        break;
      case "dashboard":
        response = withSession_(params, function() {
          return success_({ items: getDashboardItems_() });
        });
        break;
      case "expiring":
        response = withSession_(params, function() {
          return success_({ items: getExpiringSoon_() });
        });
        break;
      case "logs":
        response = withSession_(params, function() {
          return success_({ rows: getLogRows_() });
        });
        break;
      case "analytics":
        response = withSession_(params, function() {
          return success_(getAnalytics_());
        });
        break;
      case "stock-details":
        response = withSession_(params, function() {
          return success_({ items: getStockDetails_(params.name || "") });
        });
        break;
      case "resolve-scan":
        response = withSession_(params, function() {
          return success_(resolveScanCode_(params.code || ""));
        });
        break;
      case "withdraw":
        response = withSession_(params, function(session) {
          return apiWithdrawStock_(params, session.user);
        });
        break;
      case "receive":
        response = withSession_(params, function(session) {
          return apiReceiveStock_(params, session.user);
        });
        break;
      case "dispose":
        response = withSession_(params, function(session) {
          return apiDisposeStock_(params, session.user);
        });
        break;
      case "adjust-stock":
        response = withSession_(params, function(session) {
          requireAdmin_(session);
          return apiAdjustStock_(params, session.user);
        });
        break;
      case "master-list":
        response = withSession_(params, function() {
          return success_({ rows: getMasterRows_() });
        });
        break;
      case "update-master":
        response = withSession_(params, function(session) {
          requireAdmin_(session);
          return apiUpdateMasterItem_(params, session.user);
        });
        break;
      case "delete-master":
        response = withSession_(params, function(session) {
          requireAdmin_(session);
          return apiDeleteMasterItem_(params.name || "", session.user);
        });
        break;
      case "users":
        response = withSession_(params, function(session) {
          requireAdmin_(session);
          return success_({ rows: getUserRows_() });
        });
        break;
      case "update-user":
        response = withSession_(params, function(session) {
          requireAdmin_(session);
          return apiUpdateUser_(params, session.user);
        });
        break;
      case "delete-user":
        response = withSession_(params, function(session) {
          requireAdmin_(session);
          return apiDeleteUser_(params.username || "", session.user);
        });
        break;
      case "config":
        response = withSession_(params, function(session) {
          requireAdmin_(session);
          return success_({ config: getConfigMap_() });
        });
        break;
      case "update-config":
        response = withSession_(params, function(session) {
          requireAdmin_(session);
          return apiUpdateConfig_(params.key || "", params.value || "", session.user);
        });
        break;
      case "monthly-summary":
        response = withSession_(params, function() {
          return success_({ rows: getMonthlySummary_(params.year, params.month) });
        });
        break;
      case "system-status":
        var ss = getSS_();
        response = success_({
          ok: true,
          spreadsheetId: ss.getId(),
          spreadsheetName: ss.getName(),
          sheets: ss.getSheets().map(function(sheet) { return sheet.getName(); }),
          usersSheetColumns: getSheetHeaders_(getSheet_(SHEET_USERS))
        });
        break;
      default:
        response = failure_("UNKNOWN_ACTION", "Unknown action: " + action);
    }

    return buildOutput_(response, params);
  } catch (error) {
    return buildOutput_(failure_("SERVER_ERROR", error.message || String(error)), getParams_(e));
  }
}

function getParams_(e) {
  var params = {};
  if (!e) return params;
  if (e.parameter) {
    Object.keys(e.parameter).forEach(function(key) {
      params[key] = e.parameter[key];
    });
  }
  if (e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      Object.keys(body).forEach(function(key) {
        params[key] = body[key];
      });
    } catch (error) {}
  }
  return params;
}

function buildOutput_(payload, params) {
  var callback = (params.callback || params.prefix || "").trim();
  var json = JSON.stringify(payload);

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function success_(data) {
  var payload = data || {};
  payload.success = true;
  return payload;
}

function failure_(code, message) {
  return {
    success: false,
    errorCode: code || "ERROR",
    message: message || "Unknown error"
  };
}

function requireAdmin_(session) {
  if ((session.user.role || "").toUpperCase() !== "ADMIN") {
    throw new Error("ADMIN_REQUIRED");
  }
}

function withSession_(params, callback) {
  var token = params.token || "";
  if (!token) {
    return failure_("UNAUTHORIZED", "Missing session token");
  }

  var cache = CacheService.getScriptCache();
  var raw = cache.get("session:" + token);
  if (!raw) {
    return failure_("SESSION_EXPIRED", "Session expired or not found");
  }

  var session = JSON.parse(raw);
  cache.put("session:" + token, raw, SESSION_TTL_SECONDS);
  return callback(session);
}

function apiLogin_(params) {
  var username = normalizeKey_(params.username);
  var passwordHash = String(params.passwordHash || "");
  if (!username || !passwordHash) {
    return failure_("INVALID_LOGIN", "Username and password are required");
  }

  var sheet = getSheet_(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = mapUserRow_(data[i]);
    if (!row.username) continue;
    if (normalizeKey_(row.username) !== username) continue;
    if (String(row.status || "").toUpperCase() !== "ACTIVE") {
      return failure_("USER_DISABLED", "This user is inactive");
    }
    if (String(row.passwordHash || "") !== passwordHash) {
      return failure_("INVALID_LOGIN", "Invalid username or password");
    }

    var token = Utilities.getUuid().replace(/-/g, "") + String(new Date().getTime());
    var session = {
      token: token,
      loginAt: new Date().toISOString(),
      user: {
        username: row.username,
        fullName: row.fullName,
        department: row.department,
        role: row.role,
        status: row.status
      }
    };

    CacheService.getScriptCache().put("session:" + token, JSON.stringify(session), SESSION_TTL_SECONDS);
    sheet.getRange(i + 1, 7).setValue(new Date());
    appendAuditLog_("LOGIN", row.fullName || row.username, "Successful login");
    return success_({ token: token, user: session.user });
  }

  return failure_("INVALID_LOGIN", "Invalid username or password");
}

function apiRegisterUser_(params) {
  var username = normalizeKey_(params.username);
  var passwordHash = String(params.passwordHash || "");
  var fullName = String(params.fullName || "").trim();
  var department = String(params.department || "").trim();

  if (!username || !passwordHash || !fullName || !department) {
    return failure_("INVALID_REGISTER", "Username, password, full name and department are required");
  }

  if (!/^[a-z0-9._-]{3,40}$/i.test(username)) {
    return failure_("INVALID_REGISTER", "Username must be 3-40 characters and use only letters, numbers, dot, dash or underscore");
  }

  var sheet = getSheet_(SHEET_USERS);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var existing = mapUserRow_(data[i]);
      if (normalizeKey_(existing.username) === username) {
        return failure_("DUPLICATE_USER", "Username already exists");
      }
    }

    sheet.appendRow([username, passwordHash, fullName, department, "USER", "ACTIVE", ""]);
    appendAuditLog_("REGISTER", fullName, username + " / " + department);
    return success_({
      message: "สมัครสมาชิกสำเร็จ",
      user: {
        username: username,
        fullName: fullName,
        department: department,
        role: "USER",
        status: "ACTIVE"
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function getSS_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("No active spreadsheet");
    }
    return ss;
  } catch (error) {
    throw new Error("Spreadsheet not found. Please open this script from Extensions > Apps Script inside the target Google Sheet.");
  }
}

function getSheet_(name) {
  var sheet = getSS_().getSheetByName(name);
  if (!sheet) {
    throw new Error("Sheet not found: " + name);
  }
  return sheet;
}

function getSheetHeaders_(sheet) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
}

function initializeSchema_(ss) {
  ensureSheet_(ss, SHEET_MASTER, ["Chemical Name", "Min Stock", "Hazard Type", "SDS Link", "Location", "Supplier", "Unit Price"]);
  ensureSheet_(ss, SHEET_STOCK, ["Lot ID", "Chemical Name", "Lot No", "Receive Date", "MFG Date", "EXP Date", "Received Qty", "Balance", "Status", "Note", "Update Date", "Update By"]);
  ensureSheet_(ss, SHEET_LOG, ["Timestamp", "Type", "Chemical Name", "Qty", "User", "Ref ID", "Note"]);
  ensureSheet_(ss, SHEET_USERS, ["Username", "Password Hash", "Full Name", "Department", "Role", "Status", "Last Login"]);
  ensureSheet_(ss, SHEET_CONFIG, ["Key", "Value"]);
  ensureSheet_(ss, SHEET_AUDIT, ["Timestamp", "Action", "Actor", "Detail"]);

  migrateUsersSheet_();
  seedDefaults_();
}

function initializeSchema() {
  return resetAndSeedDemoData();
}

function setupDemoData() {
  return resetAndSeedDemoData();
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  var currentHeaders = getSheetHeaders_(sheet);
  if (String(currentHeaders[0] || "") !== String(headers[0] || "")) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  if (sheet.getLastColumn() < headers.length) {
    sheet.insertColumnsAfter(sheet.getLastColumn(), headers.length - sheet.getLastColumn());
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f1f5f9");
}

function migrateUsersSheet_() {
  var sheet = getSheet_(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var headers = data[0];
  if (headers[0] === "Username" && headers[1] === "Password Hash") return;

  var migrated = [["Username", "Password Hash", "Full Name", "Department", "Role", "Status", "Last Login"]];
  for (var i = 1; i < data.length; i++) {
    var oldRow = data[i];
    if (!oldRow[0]) continue;
    migrated.push([
      oldRow[0],
      hashString_("1234"),
      oldRow[0],
      oldRow[1] || "",
      oldRow[2] || "USER",
      oldRow[3] || "ACTIVE",
      ""
    ]);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, migrated.length, migrated[0].length).setValues(migrated);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, migrated[0].length).setFontWeight("bold").setBackground("#f1f5f9");
}

function seedDefaults_() {
  seedConfigIfEmpty_();
  seedUsersIfEmpty_();
  seedMasterIfEmpty_();
  seedStockIfEmpty_();
}

function seedConfigIfEmpty_() {
  var sheet = getSheet_(SHEET_CONFIG);
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, 4, 2).setValues([
    ["APP_NAME", "Chemical Requisition System"],
    ["LINE_TOKEN", ""],
    ["SCAN_HINT", "สแกน QR หรือ Barcode จากกล้องได้"],
    ["DEFAULT_LOCATION", "Main Chemical Room"]
  ]);
}

function seedUsersIfEmpty_() {
  var sheet = getSheet_(SHEET_USERS);
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, 3, 7).setValues([
    ["admin", hashString_("admin1234"), "System Admin", "Central Lab", "ADMIN", "ACTIVE", ""],
    ["lab001", hashString_("lab1234"), "Lab Operator", "Quality Control", "USER", "ACTIVE", ""],
    ["store001", hashString_("store1234"), "Store Keeper", "Warehouse", "USER", "ACTIVE", ""]
  ]);
}

function seedMasterIfEmpty_() {
  var sheet = getSheet_(SHEET_MASTER);
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, 4, 7).setValues([
    ["Buffer pH 7.00", 3, "General", "https://example.com/sds/ph7", "Shelf A-01", "Merck", 320],
    ["Hydrochloric Acid 37%", 2, "Acid", "https://example.com/sds/hcl", "Cabinet C-02", "RCI Labscan", 890],
    ["Methanol", 4, "Flame", "https://example.com/sds/meoh", "Cabinet F-03", "Sigma-Aldrich", 560],
    ["Sodium Hydroxide", 2, "General", "https://example.com/sds/naoh", "Shelf B-04", "Ajax Finechem", 410]
  ]);
}

function seedStockIfEmpty_() {
  var sheet = getSheet_(SHEET_STOCK);
  if (sheet.getLastRow() > 1) return;

  var now = new Date();
  var rows = [
    buildStockRow_("LOT-1001", "Buffer pH 7.00", "PH700-A1", daysFrom_(now, -20), daysFrom_(now, -40), daysFrom_(now, 180), 1, 1, "ACTIVE", "Demo stock", now, "seed"),
    buildStockRow_("LOT-1002", "Buffer pH 7.00", "PH700-A2", daysFrom_(now, -10), daysFrom_(now, -30), daysFrom_(now, 120), 1, 1, "ACTIVE", "Demo stock", now, "seed"),
    buildStockRow_("LOT-2001", "Hydrochloric Acid 37%", "HCL-2401", daysFrom_(now, -12), daysFrom_(now, -45), daysFrom_(now, 45), 1, 1, "ACTIVE", "Demo stock", now, "seed"),
    buildStockRow_("LOT-3001", "Methanol", "MEOH-2402", daysFrom_(now, -7), daysFrom_(now, -30), daysFrom_(now, 90), 1, 1, "ACTIVE", "Demo stock", now, "seed"),
    buildStockRow_("LOT-4001", "Sodium Hydroxide", "NAOH-2401", daysFrom_(now, -15), daysFrom_(now, -60), daysFrom_(now, 365), 1, 1, "ACTIVE", "Demo stock", now, "seed")
  ];

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  var logSheet = getSheet_(SHEET_LOG);
  if (logSheet.getLastRow() <= 1) {
    logSheet.getRange(2, 1, 2, 7).setValues([
      [daysFrom_(now, -10), "IN", "Buffer pH 7.00", 2, "seed", "LOT-1001, LOT-1002", "Initial demo data"],
      [daysFrom_(now, -7), "IN", "Methanol", 1, "seed", "LOT-3001", "Initial demo data"]
    ]);
  }
}

function buildStockRow_(lotId, chemName, lotNo, receiveDate, mfgDate, expDate, receivedQty, balance, status, note, updateDate, updateBy) {
  return [lotId, chemName, lotNo, receiveDate, mfgDate, expDate, receivedQty, balance, status, note, updateDate, updateBy];
}

function seedDummyData() {
  initializeSchema_(getSS_());
}

function resetAndSeedDemoData() {
  var ss = getSS_();
  resetWorkbookForDemo_(ss);
  initializeSchema_(ss);
  return {
    success: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheets: ss.getSheets().map(function(sheet) { return sheet.getName(); }),
    message: "Workbook was reset and seeded with demo data"
  };
}

function resetWorkbookForDemo_(ss) {
  var sheets = ss.getSheets();
  if (!sheets.length) {
    ss.insertSheet("Temp_Init");
    sheets = ss.getSheets();
  }

  var primarySheet = sheets[0];
  for (var i = sheets.length - 1; i >= 1; i--) {
    ss.deleteSheet(sheets[i]);
  }

  primarySheet.clear();
  primarySheet.clearFormats();
  if (primarySheet.getName() !== SHEET_MASTER) {
    primarySheet.setName(SHEET_MASTER);
  }
}

function getDashboardItems_() {
  var masterRows = getMasterRows_();
  var stockRows = getStockRows_();
  var masterMap = {};
  masterRows.forEach(function(row) {
    masterMap[row.name] = row;
  });

  var summary = {};
  stockRows.forEach(function(row) {
    if (row.status !== "ACTIVE" || Number(row.balance || 0) <= 0) return;
    if (!summary[row.name]) {
      var master = masterMap[row.name] || {};
      summary[row.name] = {
        name: row.name,
        hazard: master.hazard || "General",
        sds: master.sds || "",
        location: master.location || "",
        supplier: master.supplier || "",
        price: Number(master.price || 0),
        minStock: Number(master.minStock || 0),
        stats: {
          total: 0,
          available: 0,
          nextExp: ""
        }
      };
    }

    summary[row.name].stats.total += Number(row.receivedQty || 0);
    summary[row.name].stats.available += Number(row.balance || 0);
    if (row.exp && (!summary[row.name].stats.nextExp || new Date(row.exp) < new Date(summary[row.name].stats.nextExp))) {
      summary[row.name].stats.nextExp = row.exp;
    }
  });

  return Object.keys(summary).sort().map(function(key) {
    return summary[key];
  });
}

function getStockDetails_(chemName) {
  return getStockRows_()
    .filter(function(row) {
      return row.name === chemName && row.status === "ACTIVE" && Number(row.balance || 0) > 0;
    })
    .sort(function(a, b) {
      return safeDate_(a.exp).getTime() - safeDate_(b.exp).getTime();
    });
}

function getExpiringSoon_() {
  var today = stripTime_(new Date());
  var next30Days = daysFrom_(today, 30);
  return getStockRows_()
    .filter(function(row) {
      var expDate = safeDate_(row.exp);
      return row.status === "ACTIVE" &&
        Number(row.balance || 0) > 0 &&
        expDate >= today &&
        expDate <= next30Days;
    })
    .sort(function(a, b) {
      return safeDate_(a.exp).getTime() - safeDate_(b.exp).getTime();
    });
}

function getAnalytics_() {
  var rows = getLogRows_();
  var stats = {
    IN: 0,
    OUT: 0,
    DISPOSAL: 0,
    ADJUST: 0,
    monthly: {}
  };

  rows.forEach(function(row) {
    var qty = Number(row.qty || 0);
    if (stats[row.type] !== undefined) {
      stats[row.type] += qty;
    }

    var date = safeDate_(row.timestamp);
    var monthKey = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM");
    if (!stats.monthly[monthKey]) {
      stats.monthly[monthKey] = { IN: 0, OUT: 0, DISPOSAL: 0, ADJUST: 0 };
    }
    if (stats.monthly[monthKey][row.type] !== undefined) {
      stats.monthly[monthKey][row.type] += qty;
    }
  });

  return stats;
}

function getMonthlySummary_(year, month) {
  var targetYear = Number(year || 0);
  var targetMonth = Number(month || 0);
  var rows = getLogRows_().filter(function(row) {
    var date = safeDate_(row.timestamp);
    return date.getFullYear() === targetYear && (date.getMonth() + 1) === targetMonth;
  });

  var result = [["Timestamp", "Type", "Chemical Name", "Qty", "User", "Ref ID", "Note"]];
  rows.forEach(function(row) {
    result.push([row.timestamp, row.type, row.name, row.qty, row.user, row.refId, row.note]);
  });
  return result;
}

function getLogRows_() {
  var sheet = getSheet_(SHEET_LOG);
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({
      timestamp: data[i][0],
      type: data[i][1],
      name: data[i][2],
      qty: Number(data[i][3] || 0),
      user: data[i][4],
      refId: data[i][5],
      note: data[i][6]
    });
  }
  return rows;
}

function getMasterRows_() {
  var sheet = getSheet_(SHEET_MASTER);
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({
      name: data[i][0],
      minStock: Number(data[i][1] || 0),
      hazard: data[i][2] || "General",
      sds: data[i][3] || "",
      location: data[i][4] || "",
      supplier: data[i][5] || "",
      price: Number(data[i][6] || 0)
    });
  }
  return rows;
}

function getUserRows_() {
  var sheet = getSheet_(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = mapUserRow_(data[i]);
    if (!row.username) continue;
    rows.push(row);
  }
  return rows;
}

function getStockRows_() {
  var sheet = getSheet_(SHEET_STOCK);
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({
      id: data[i][0],
      name: data[i][1],
      lotNo: data[i][2],
      receiveDate: data[i][3],
      mfg: data[i][4],
      exp: data[i][5],
      receivedQty: Number(data[i][6] || 0),
      balance: Number(data[i][7] || 0),
      status: data[i][8],
      note: data[i][9] || "",
      updateDate: data[i][10] || "",
      updateBy: data[i][11] || ""
    });
  }
  return rows;
}

function mapUserRow_(row) {
  return {
    username: row[0] || "",
    passwordHash: row[1] || "",
    fullName: row[2] || row[0] || "",
    department: row[3] || "",
    role: (row[4] || "USER").toUpperCase(),
    status: (row[5] || "ACTIVE").toUpperCase(),
    lastLogin: row[6] || ""
  };
}

function apiWithdrawStock_(params, actor) {
  var chemicalName = String(params.name || "");
  var qty = Math.max(1, Number(params.qty || 1));
  var note = String(params.note || "Withdraw via web");
  var requestedBy = actor.fullName || actor.username;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_STOCK);
    var data = sheet.getDataRange().getValues();
    var candidates = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === chemicalName && data[i][8] === "ACTIVE" && Number(data[i][7] || 0) > 0) {
        candidates.push({ rowIndex: i + 1, lotId: data[i][0], exp: safeDate_(data[i][5]) });
      }
    }

    candidates.sort(function(a, b) { return a.exp.getTime() - b.exp.getTime(); });
    if (candidates.length < qty) {
      return failure_("INSUFFICIENT_STOCK", "Stock ไม่พอสำหรับการเบิก");
    }

    var usedLotIds = [];
    for (var j = 0; j < qty; j++) {
      var item = candidates[j];
      sheet.getRange(item.rowIndex, 8).setValue(0);
      sheet.getRange(item.rowIndex, 9).setValue("USED");
      sheet.getRange(item.rowIndex, 11, 1, 2).setValues([[new Date(), requestedBy]]);
      usedLotIds.push(item.lotId);
    }

    appendTransaction_("OUT", chemicalName, qty, requestedBy, usedLotIds.join(", "), note);
    checkStockAlert_(chemicalName);
    appendAuditLog_("WITHDRAW", requestedBy, chemicalName + " x" + qty);
    return success_({ message: "เบิกสารเคมีเรียบร้อย", usedLotIds: usedLotIds });
  } finally {
    lock.releaseLock();
  }
}

function apiReceiveStock_(params, actor) {
  var chemicalName = String(params.chemicalName || "");
  var qty = Math.max(1, Number(params.qty || 1));
  var lotNo = String(params.lotNo || "");
  var mfgDate = params.mfgDate ? safeDate_(params.mfgDate) : "";
  var expDate = params.expDate ? safeDate_(params.expDate) : "";
  var note = String(params.note || "Receive via web");
  var requestedBy = actor.fullName || actor.username;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_STOCK);
    var timestamp = new Date();
    var lotIds = [];

    for (var i = 0; i < qty; i++) {
      var lotId = "LOT-" + timestamp.getTime() + "-" + (i + 1);
      sheet.appendRow([lotId, chemicalName, lotNo, timestamp, mfgDate, expDate, 1, 1, "ACTIVE", note, timestamp, requestedBy]);
      lotIds.push(lotId);
    }

    appendTransaction_("IN", chemicalName, qty, requestedBy, lotIds.join(", "), note);
    appendAuditLog_("RECEIVE", requestedBy, chemicalName + " x" + qty);
    return success_({ message: "รับเข้าสต็อกเรียบร้อย", lotIds: lotIds });
  } finally {
    lock.releaseLock();
  }
}

function apiDisposeStock_(params, actor) {
  var chemicalName = String(params.name || "");
  var qty = Math.max(1, Number(params.qty || 1));
  var note = String(params.note || "Dispose via web");
  var requestedBy = actor.fullName || actor.username;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_STOCK);
    var data = sheet.getDataRange().getValues();
    var candidates = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === chemicalName && data[i][8] === "ACTIVE" && Number(data[i][7] || 0) > 0) {
        candidates.push({ rowIndex: i + 1, lotId: data[i][0], exp: safeDate_(data[i][5]) });
      }
    }

    candidates.sort(function(a, b) { return a.exp.getTime() - b.exp.getTime(); });
    var items = candidates.slice(0, qty);
    if (items.length === 0) {
      return failure_("NOT_FOUND", "ไม่พบสต็อกที่พร้อมทำลาย");
    }

    items.forEach(function(item) {
      sheet.getRange(item.rowIndex, 8).setValue(0);
      sheet.getRange(item.rowIndex, 9).setValue("DISPOSED");
      sheet.getRange(item.rowIndex, 11, 1, 2).setValues([[new Date(), requestedBy]]);
    });

    appendTransaction_("DISPOSAL", chemicalName, items.length, requestedBy, items.map(function(item) { return item.lotId; }).join(", "), note);
    appendAuditLog_("DISPOSE", requestedBy, chemicalName + " x" + items.length);
    return success_({ message: "ทำลายสารเคมีเรียบร้อย", lotIds: items.map(function(item) { return item.lotId; }) });
  } finally {
    lock.releaseLock();
  }
}

function apiAdjustStock_(params, actor) {
  var lotId = String(params.lotId || "");
  var newBalance = Number(params.newBalance || 0);
  var newLotNo = String(params.newLotNo || "");
  var newExp = params.newExp ? safeDate_(params.newExp) : "";
  var note = String(params.note || "Manual adjustment");
  var requestedBy = actor.fullName || actor.username;

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_STOCK);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== lotId) continue;

      var oldBalance = Number(data[i][7] || 0);
      sheet.getRange(i + 1, 3).setValue(newLotNo || data[i][2]);
      sheet.getRange(i + 1, 6).setValue(newExp || data[i][5]);
      sheet.getRange(i + 1, 8).setValue(newBalance);
      sheet.getRange(i + 1, 9).setValue(newBalance > 0 ? "ACTIVE" : "USED");
      sheet.getRange(i + 1, 10).setValue(note);
      sheet.getRange(i + 1, 11, 1, 2).setValues([[new Date(), requestedBy]]);

      appendTransaction_("ADJUST", data[i][1], newBalance - oldBalance, requestedBy, lotId, note);
      appendAuditLog_("ADJUST", requestedBy, lotId + " => " + newBalance);
      return success_({ message: "ปรับยอดเรียบร้อย" });
    }

    return failure_("NOT_FOUND", "ไม่พบ lot ที่ระบุ");
  } finally {
    lock.releaseLock();
  }
}

function apiUpdateMasterItem_(params, actor) {
  var oldName = String(params.oldName || "");
  var newName = String(params.name || "");
  if (!newName) {
    return failure_("INVALID_DATA", "Chemical name is required");
  }

  var minStock = Number(params.minStock || 0);
  var hazard = String(params.hazard || "General");
  var sds = String(params.sds || "");
  var location = String(params.location || "");
  var supplier = String(params.supplier || "");
  var price = Number(params.price || 0);

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_MASTER);
    var data = sheet.getDataRange().getValues();
    var foundRow = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === oldName || data[i][0] === newName) {
        foundRow = i + 1;
        break;
      }
    }

    var values = [newName, minStock, hazard, sds, location, supplier, price];
    if (foundRow) {
      sheet.getRange(foundRow, 1, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }

    if (oldName && oldName !== newName) {
      syncChemicalName_(oldName, newName);
    }

    appendAuditLog_("MASTER_SAVE", actor.fullName || actor.username, newName);
    return success_({ message: "บันทึกข้อมูลสารเคมีเรียบร้อย" });
  } finally {
    lock.releaseLock();
  }
}

function syncChemicalName_(oldName, newName) {
  var stockSheet = getSheet_(SHEET_STOCK);
  var logSheet = getSheet_(SHEET_LOG);

  var stockData = stockSheet.getDataRange().getValues();
  for (var i = 1; i < stockData.length; i++) {
    if (stockData[i][1] === oldName) {
      stockSheet.getRange(i + 1, 2).setValue(newName);
    }
  }

  var logData = logSheet.getDataRange().getValues();
  for (var j = 1; j < logData.length; j++) {
    if (logData[j][2] === oldName) {
      logSheet.getRange(j + 1, 3).setValue(newName);
    }
  }
}

function apiDeleteMasterItem_(name, actor) {
  if (!name) {
    return failure_("INVALID_DATA", "Name is required");
  }
  var sheet = getSheet_(SHEET_MASTER);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      sheet.deleteRow(i + 1);
      appendAuditLog_("MASTER_DELETE", actor.fullName || actor.username || actor, name);
      return success_({ message: "ลบข้อมูลสารเคมีแล้ว" });
    }
  }
  return failure_("NOT_FOUND", "Chemical not found");
}

function apiUpdateUser_(params, actor) {
  var oldUsername = String(params.oldUsername || "");
  var username = String(params.username || "");
  var fullName = String(params.fullName || username);
  var department = String(params.department || "");
  var role = String(params.role || "USER").toUpperCase();
  var status = String(params.status || "ACTIVE").toUpperCase();
  var passwordHash = String(params.passwordHash || "");

  if (!username) {
    return failure_("INVALID_DATA", "Username is required");
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_USERS);
    var data = sheet.getDataRange().getValues();
    var targetRow = 0;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === oldUsername || data[i][0] === username) {
        targetRow = i + 1;
        break;
      }
    }

    if (!passwordHash && targetRow) {
      passwordHash = String(sheet.getRange(targetRow, 2).getValue() || "");
    }
    if (!passwordHash) {
      passwordHash = hashString_("1234");
    }

    var values = [username, passwordHash, fullName, department, role, status, targetRow ? sheet.getRange(targetRow, 7).getValue() : ""];
    if (targetRow) {
      sheet.getRange(targetRow, 1, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }

    appendAuditLog_("USER_SAVE", actor.fullName || actor.username, username);
    return success_({ message: "บันทึกผู้ใช้งานเรียบร้อย" });
  } finally {
    lock.releaseLock();
  }
}

function apiDeleteUser_(username, actor) {
  if (!username) {
    return failure_("INVALID_DATA", "Username is required");
  }
  var sheet = getSheet_(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      sheet.deleteRow(i + 1);
      appendAuditLog_("USER_DELETE", actor.fullName || actor.username || actor, username);
      return success_({ message: "ลบผู้ใช้งานแล้ว" });
    }
  }
  return failure_("NOT_FOUND", "User not found");
}

function getConfigMap_() {
  var sheet = getSheet_(SHEET_CONFIG);
  var data = sheet.getDataRange().getValues();
  var config = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    config[data[i][0]] = data[i][1];
  }
  return config;
}

function apiUpdateConfig_(key, value, actor) {
  if (!key) {
    return failure_("INVALID_DATA", "Key is required");
  }
  var sheet = getSheet_(SHEET_CONFIG);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      appendAuditLog_("CONFIG_UPDATE", actor.fullName || actor.username || actor, key);
      return success_({ message: "อัปเดตการตั้งค่าแล้ว" });
    }
  }
  sheet.appendRow([key, value]);
  appendAuditLog_("CONFIG_CREATE", actor.fullName || actor.username || actor, key);
  return success_({ message: "บันทึกการตั้งค่าแล้ว" });
}

function resolveScanCode_(code) {
  var rawCode = String(code || "").trim();
  if (!rawCode) {
    return { found: false, message: "Empty scan result" };
  }

  var normalized = rawCode.replace(/^LOT:/i, "").trim();
  var rows = getStockRows_();

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row.id === normalized || row.lotNo === normalized || row.name === rawCode || row.name === normalized) {
      return {
        found: true,
        rawCode: rawCode,
        chemicalName: row.name,
        lotId: row.id,
        lotNo: row.lotNo,
        balance: row.balance,
        status: row.status
      };
    }
  }

  return {
    found: false,
    rawCode: rawCode,
    message: "ไม่พบข้อมูลที่ตรงกับ QR/Barcode นี้"
  };
}

function appendTransaction_(type, name, qty, user, refId, note) {
  getSheet_(SHEET_LOG).appendRow([new Date(), type, name, qty, user, refId, note]);
}

function appendAuditLog_(action, actor, detail) {
  getSheet_(SHEET_AUDIT).appendRow([new Date(), action, actor, detail]);
}

function checkStockAlert_(chemicalName) {
  var items = getDashboardItems_();
  var current = null;
  for (var i = 0; i < items.length; i++) {
    if (items[i].name === chemicalName) {
      current = items[i];
      break;
    }
  }
  if (!current) return;
  if (Number(current.stats.available || 0) > Number(current.minStock || 0)) return;

  var lineToken = getConfigMap_().LINE_TOKEN || "";
  if (!lineToken) return;

  var message =
    "\nแจ้งเตือนสารเคมีใกล้หมด" +
    "\nสาร: " + chemicalName +
    "\nคงเหลือ: " + current.stats.available +
    "\nขั้นต่ำ: " + current.minStock;
  sendLineNotify_(lineToken, message);
}

function sendLineNotify_(token, message) {
  try {
    UrlFetchApp.fetch("https://notify-api.line.me/api/notify", {
      method: "post",
      headers: { Authorization: "Bearer " + token },
      payload: { message: message },
      muteHttpExceptions: true
    });
  } catch (error) {}
}

function hashString_(text) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return digest.map(function(byte) {
    var value = (byte < 0 ? byte + 256 : byte).toString(16);
    return value.length === 1 ? "0" + value : value;
  }).join("");
}

function normalizeKey_(value) {
  return String(value || "").trim().toLowerCase();
}

function safeDate_(value) {
  if (!value) return new Date(0);
  return value instanceof Date ? value : new Date(value);
}

function daysFrom_(date, days) {
  var next = new Date(date.getTime());
  next.setDate(next.getDate() + days);
  return next;
}

function stripTime_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}
