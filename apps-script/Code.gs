/*************************************************
 * VOID X VPN™ - COMPLETE GOOGLE APPS SCRIPT
 * Public Website: Vercel (Index.html)
 * Admin Dashboard: Admin.html
 * Database: Google Sheets
 * + REST API for Vercel Frontend
 *************************************************/

const SHEET_PACKAGES = "Packages";
const SHEET_CONFIGS = "Configs";
const SHEET_EHI_FILES = "EhiFiles";
const EHI_FOLDER_PROPERTY = "EHI_DRIVE_FOLDER_ID";
const SESSION_PREFIX = "VOID_ADMIN_";
const SESSION_SECONDS = 21600;

const PACKAGE_HEADERS = [
  "id", "network", "accessType", "name",
  "description", "active", "createdAt",
  "price", "dataAmount"
];

const CONFIG_HEADERS = [
  "id", "packageId", "type", "name",
  "content", "active", "createdAt"
];

const EHI_HEADERS = [
  "id", "packageId", "name", "fileId",
  "fileName", "mimeType", "active", "createdAt"
];

/* ==============================================
   CORS HELPERS (Vercel frontend සඳහා)
============================================== */

function corsHeaders_() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Max-Age": "86400"
  };
}

function jsonResponse_(obj) {
  var out = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  var h = corsHeaders_();
  for (var k in h) out.setHeader(k, h[k]);
  return out;
}

/* ==============================================
   1. WEBSITE ROUTING + REST API
============================================== */

function doGet(e) {
  var page = e && e.parameter ? e.parameter.page : "";
  var action = e && e.parameter ? e.parameter.action : "";

  // ---- REST API (Vercel frontend එකෙන් GET) ----
  if (action) {
    try {
      if (action === "publicData") {
        return jsonResponse_({ status: "success", data: getPublicData() });
      }
      return jsonResponse_({ status: "error", message: "Unknown action: " + action });
    } catch (err) {
      return jsonResponse_({ status: "error", message: err.toString() });
    }
  }

  // ---- Admin page (GAS ඇතුළෙම) ----
  if (page === "admin") {
    return HtmlService.createHtmlOutputFromFile("Admin")
      .setTitle("VOID X VPN™ | Admin Dashboard")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ---- Index page (ඔබ Vercel එකේ host කරන නිසා මේක backup එකක් විතරයි) ----
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("VOID X VPN™ | KULITH")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---- CORS preflight (OPTIONS) ----
function doOptions(e) {
  var out = ContentService.createTextOutput("");
  var h = corsHeaders_();
  for (var k in h) out.setHeader(k, h[k]);
  return out;
}

// ---- REST API (Vercel frontend එකෙන් POST) ----
function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (x) { body = {}; }
    }
    if (!body.action && e && e.parameter) {
      body = Object.assign({}, e.parameter);
    }

    var action = String(body.action || "").trim();
    var token = String(body.token || "").trim();
    var payload = body.data || {};

    switch (action) {
      // ---------- Public ----------
      case "publicData":
        return jsonResponse_({ status: "success", data: getPublicData() });

      // ---------- Admin Auth ----------
      case "adminLogin":
        return jsonResponse_({
          status: "success",
          data: adminLogin(body.username, body.password)
        });

      case "adminLogout":
        return jsonResponse_({ status: "success", data: adminLogout(token) });

      case "adminData":
        return jsonResponse_({ status: "success", data: getAdminData(token) });

      // ---------- Packages ----------
      case "addPackage":
        return jsonResponse_({ status: "success", data: addPackage(token, payload) });
      case "updatePackage":
        return jsonResponse_({ status: "success", data: updatePackage(token, payload) });
      case "deletePackage":
        return jsonResponse_({ status: "success", data: deletePackage(token, payload.id) });

      // ---------- Configs ----------
      case "addConfig":
        return jsonResponse_({ status: "success", data: addConfig(token, payload) });
      case "updateConfig":
        return jsonResponse_({ status: "success", data: updateConfig(token, payload) });
      case "deleteConfig":
        return jsonResponse_({ status: "success", data: deleteConfig(token, payload.id) });

      // ---------- EHI files ----------
      case "addEhiFile":
        return jsonResponse_({
          status: "success",
          data: addEhiFile(
            token,
            payload.packageId,
            payload.name,
            payload.fileName,
            payload.dataUrl
          )
        });
      case "deleteEhiFile":
        return jsonResponse_({ status: "success", data: deleteEhiFile(token, payload.id) });

      default:
        return jsonResponse_({ status: "error", message: "Unknown action: " + action });
    }
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }
}

/* ==============================================
   2. DATABASE SETUP + SAFE MIGRATION
============================================== */

function setupDatabase() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const savedId = props.getProperty("SPREADSHEET_ID");

  if (savedId) {
    ss = SpreadsheetApp.openById(savedId);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  if (!ss) {
    throw new Error(
      "Google Sheet not connected. Open Apps Script from your Google Sheet."
    );
  }

  createSheet_(ss, SHEET_PACKAGES, PACKAGE_HEADERS);
  createSheet_(ss, SHEET_CONFIGS, CONFIG_HEADERS);
  createSheet_(ss, SHEET_EHI_FILES, EHI_HEADERS);

  props.setProperty("SPREADSHEET_ID", ss.getId());

  return "VOID X VPN database is ready! Existing data has been preserved.";
}

function createSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const existingHeaders = sheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0]
      .map(value => String(value).trim());

    headers.forEach(header => {
      if (!existingHeaders.includes(header)) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
        existingHeaders.push(header);
      }
    });
  }
  sheet.setFrozenRows(1);
}

/* ==============================================
   3. DATABASE CONNECTION
============================================== */

function getDatabase_() {
  const id = PropertiesService.getScriptProperties()
    .getProperty("SPREADSHEET_ID");
  if (id) return SpreadsheetApp.openById(id);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Database not connected. Run setupDatabase() first.");

  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", ss.getId());
  return ss;
}

function getSheet_(name) {
  const sheet = getDatabase_().getSheetByName(name);
  if (!sheet) {
    throw new Error("Database sheet missing: " + name + ". Run setupDatabase() first.");
  }
  return sheet;
}

function getHeaders_(sheet) {
  if (sheet.getLastColumn() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(value => String(value).trim());
}

function appendObjectRow_(sheet, data) {
  const headers = getHeaders_(sheet);
  const row = headers.map(header =>
    Object.prototype.hasOwnProperty.call(data, header) ? data[header] : ""
  );
  sheet.appendRow(row);
}

function updateObjectRow_(sheet, rowNumber, data) {
  const headers = getHeaders_(sheet);
  headers.forEach((header, index) => {
    if (Object.prototype.hasOwnProperty.call(data, header)) {
      sheet.getRange(rowNumber, index + 1).setValue(data[header]);
    }
  });
}

/* ==============================================
   4. ADMIN LOGIN
============================================== */

function adminLogin(username, password) {
  const props = PropertiesService.getScriptProperties();
  const savedUsername = props.getProperty("ADMIN_USERNAME");
  const savedPassword = props.getProperty("ADMIN_PASSWORD");

  if (!savedUsername || !savedPassword) {
    throw new Error("Admin credentials are not configured in Script Properties.");
  }

  if (String(username) !== savedUsername || String(password) !== savedPassword) {
    return { success: false, message: "Invalid username or password." };
  }

  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put(
    SESSION_PREFIX + token,
    "authorized",
    SESSION_SECONDS
  );

  return { success: true, token: token, message: "Login successful." };
}

function verifyAdmin_(token) {
  if (!token) throw new Error("Admin login required.");
  const session = CacheService.getScriptCache().get(SESSION_PREFIX + token);
  if (session !== "authorized") {
    throw new Error("Session expired. Please login again.");
  }
  return true;
}

function adminLogout(token) {
  if (token) CacheService.getScriptCache().remove(SESSION_PREFIX + token);
  return { success: true };
}

/* ==============================================
   5. PUBLIC WEBSITE DATA
============================================== */

function getPublicData() {
  const allPackages = readRows_(SHEET_PACKAGES);

  const packages = allPackages
    .filter(row => isActive_(row.active))
    .map(row => ({
      id: String(row.id || ""),
      network: String(row.network || ""),
      accessType: String(row.accessType || "").toUpperCase(),
      name: String(row.name || ""),
      description: String(row.description || ""),
      dataAmount: String(row.dataAmount || ""),
      price: Number(row.price || 0)
    }));

  const freePackageIds = new Set(
    allPackages
      .filter(row =>
        isActive_(row.active) &&
        String(row.accessType || "").toUpperCase() === "FREE"
      )
      .map(row => String(row.id || ""))
  );

  const configs = readRows_(SHEET_CONFIGS)
    .filter(row =>
      isActive_(row.active) &&
      freePackageIds.has(String(row.packageId || ""))
    )
    .map(row => ({
      id: String(row.id || ""),
      packageId: String(row.packageId || ""),
      type: String(row.type || "").toUpperCase(),
      name: String(row.name || ""),
      content: String(row.content || "")
    }));

  const ehiFiles = readRows_(SHEET_EHI_FILES)
    .filter(row =>
      isActive_(row.active) &&
      freePackageIds.has(String(row.packageId || ""))
    )
    .map(row => ({
      id: String(row.id || ""),
      packageId: String(row.packageId || ""),
      name: String(row.name || ""),
      fileName: String(row.fileName || ""),
      downloadUrl: getEhiDownloadUrl_(String(row.fileId || ""))
    }));

  return { packages: packages, configs: configs, ehiFiles: ehiFiles };
}

/* ==============================================
   6. ADMIN DASHBOARD DATA
============================================== */

function getAdminData(token) {
  verifyAdmin_(token);

  const packages = readRows_(SHEET_PACKAGES).map(row => ({
    id: String(row.id || ""),
    network: String(row.network || ""),
    accessType: String(row.accessType || "").toUpperCase(),
    name: String(row.name || ""),
    description: String(row.description || ""),
    dataAmount: String(row.dataAmount || ""),
    price: Number(row.price || 0),
    active: isActive_(row.active),
    createdAt: safeDate_(row.createdAt)
  }));

  const configs = readRows_(SHEET_CONFIGS).map(row => ({
    id: String(row.id || ""),
    packageId: String(row.packageId || ""),
    type: String(row.type || "").toUpperCase(),
    name: String(row.name || ""),
    content: String(row.content || ""),
    active: isActive_(row.active),
    createdAt: safeDate_(row.createdAt)
  }));

  const ehiFiles = readRows_(SHEET_EHI_FILES).map(row => ({
    id: String(row.id || ""),
    packageId: String(row.packageId || ""),
    name: String(row.name || ""),
    fileName: String(row.fileName || ""),
    fileId: String(row.fileId || ""),
    active: isActive_(row.active),
    createdAt: safeDate_(row.createdAt)
  }));

  return { packages: packages, configs: configs, ehiFiles: ehiFiles };
}

function safeDate_(value) {
  if (value instanceof Date) return value.toISOString();
  return String(value || "");
}

/* ==============================================
   7. ADD PACKAGE
============================================== */

function addPackage(token, data) {
  verifyAdmin_(token);
  validatePackage_(data);

  const id = Utilities.getUuid();

  appendObjectRow_(getSheet_(SHEET_PACKAGES), {
    id: id,
    network: clean_(data.network).toUpperCase(),
    accessType: clean_(data.accessType).toUpperCase(),
    name: clean_(data.name),
    description: clean_(data.description || ""),
    active: true,
    createdAt: new Date(),
    price: clean_(data.accessType).toUpperCase() === "PAID" ? Number(data.price) : "",
    dataAmount: clean_(data.accessType).toUpperCase() === "PAID" ? clean_(data.dataAmount) : ""
  });

  return { success: true, id: id };
}

/* ==============================================
   8. UPDATE PACKAGE
============================================== */

function updatePackage(token, data) {
  verifyAdmin_(token);
  validatePackage_(data);

  const sheet = getSheet_(SHEET_PACKAGES);
  const row = findRow_(sheet, data.id);
  if (!row) throw new Error("Package not found.");

  const accessType = clean_(data.accessType).toUpperCase();

  updateObjectRow_(sheet, row, {
    network: clean_(data.network).toUpperCase(),
    accessType: accessType,
    name: clean_(data.name),
    description: clean_(data.description || ""),
    active: data.active !== false,
    price: accessType === "PAID" ? Number(data.price) : "",
    dataAmount: accessType === "PAID" ? clean_(data.dataAmount) : ""
  });

  return { success: true };
}

/* ==============================================
   9. DELETE PACKAGE AND ITS CONFIGS
============================================== */

function deletePackage(token, id) {
  verifyAdmin_(token);

  const packageSheet = getSheet_(SHEET_PACKAGES);
  const configSheet = getSheet_(SHEET_CONFIGS);
  const packageRow = findRow_(packageSheet, id);
  if (!packageRow) throw new Error("Package not found.");

  const configHeaders = getHeaders_(configSheet);
  const packageIdColumn = configHeaders.indexOf("packageId") + 1;

  if (packageIdColumn > 0 && configSheet.getLastRow() >= 2) {
    const configIds = configSheet.getRange(
      2, packageIdColumn, configSheet.getLastRow() - 1, 1
    ).getValues();

    for (let i = configIds.length - 1; i >= 0; i--) {
      if (String(configIds[i][0]) === String(id)) configSheet.deleteRow(i + 2);
    }
  }

  const ehiSheet = getSheet_(SHEET_EHI_FILES);
  const ehiHeaders = getHeaders_(ehiSheet);
  const ehiPackageIdColumn = ehiHeaders.indexOf("packageId") + 1;
  const ehiFileIdColumn = ehiHeaders.indexOf("fileId") + 1;

  if (ehiPackageIdColumn > 0 && ehiSheet.getLastRow() >= 2) {
    const ehiRows = ehiSheet.getRange(
      2, 1, ehiSheet.getLastRow() - 1, ehiHeaders.length
    ).getValues();

    for (let i = ehiRows.length - 1; i >= 0; i--) {
      if (String(ehiRows[i][ehiPackageIdColumn - 1]) === String(id)) {
        const fileId = ehiFileIdColumn > 0
          ? String(ehiRows[i][ehiFileIdColumn - 1] || "")
          : "";
        if (fileId) {
          try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
        }
        ehiSheet.deleteRow(i + 2);
      }
    }
  }

  packageSheet.deleteRow(packageRow);
  return { success: true };
}

/* ==============================================
   10. ADD CONFIG
============================================== */

function addConfig(token, data) {
  verifyAdmin_(token);
  validateConfig_(data);

  const packageRow = findRow_(getSheet_(SHEET_PACKAGES), data.packageId);
  if (!packageRow) throw new Error("Selected package does not exist.");

  const id = Utilities.getUuid();

  appendObjectRow_(getSheet_(SHEET_CONFIGS), {
    id: id,
    packageId: clean_(data.packageId),
    type: clean_(data.type).toUpperCase(),
    name: clean_(data.name),
    content: String(data.content),
    active: true,
    createdAt: new Date()
  });

  return { success: true, id: id };
}

/* ==============================================
   11. UPDATE CONFIG
============================================== */

function updateConfig(token, data) {
  verifyAdmin_(token);
  validateConfig_(data);

  const sheet = getSheet_(SHEET_CONFIGS);
  const row = findRow_(sheet, data.id);
  if (!row) throw new Error("Config not found.");

  const packageRow = findRow_(getSheet_(SHEET_PACKAGES), data.packageId);
  if (!packageRow) throw new Error("Selected package does not exist.");

  updateObjectRow_(sheet, row, {
    packageId: clean_(data.packageId),
    type: clean_(data.type).toUpperCase(),
    name: clean_(data.name),
    content: String(data.content),
    active: data.active !== false
  });

  return { success: true };
}

/* ==============================================
   12. DELETE CONFIG
============================================== */

function deleteConfig(token, id) {
  verifyAdmin_(token);

  const sheet = getSheet_(SHEET_CONFIGS);
  const row = findRow_(sheet, id);
  if (!row) throw new Error("Config not found.");

  sheet.deleteRow(row);
  return { success: true };
}

/* ==============================================
   13. EHI FILE MANAGEMENT
============================================== */

function addEhiFile(token, packageId, name, fileName, dataUrl) {
  verifyAdmin_(token);

  packageId = clean_(packageId);
  name = clean_(name);
  fileName = clean_(fileName);
  dataUrl = String(dataUrl || "");

  if (!packageId) throw new Error("Please select a package.");
  if (!name) throw new Error("Please enter an EHI file name.");
  if (!fileName) throw new Error("Please select an EHI file.");
  if (!dataUrl || dataUrl.indexOf("base64,") === -1) {
    throw new Error("Invalid EHI file data.");
  }

  const packageRow = findRow_(getSheet_(SHEET_PACKAGES), packageId);
  if (!packageRow) throw new Error("Selected package does not exist.");

  if (!/\.ehi$/i.test(fileName)) throw new Error("Only .ehi files are allowed.");

  const match = dataUrl.match(/^data:([^;]+);base64,(.*)$/s);
  if (!match) throw new Error("Invalid file format.");

  const mimeType = match[1] || "application/octet-stream";
  const bytes = Utilities.base64Decode(match[2]);

  if (bytes.length > 15 * 1024 * 1024) {
    throw new Error("EHI file is too large. Maximum size is 15 MB.");
  }

  const folder = getEhiFolder_();
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const id = Utilities.getUuid();
  appendObjectRow_(getSheet_(SHEET_EHI_FILES), {
    id: id,
    packageId: packageId,
    name: name,
    fileId: file.getId(),
    fileName: fileName,
    mimeType: mimeType,
    active: true,
    createdAt: new Date()
  });

  return { success: true, id: id, fileName: fileName };
}

function deleteEhiFile(token, id) {
  verifyAdmin_(token);

  const sheet = getSheet_(SHEET_EHI_FILES);
  const row = findRow_(sheet, id);
  if (!row) throw new Error("EHI file not found.");

  const headers = getHeaders_(sheet);
  const fileIdColumn = headers.indexOf("fileId") + 1;
  const fileId = fileIdColumn > 0
    ? String(sheet.getRange(row, fileIdColumn).getValue() || "")
    : "";

  if (fileId) {
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
  }

  sheet.deleteRow(row);
  return { success: true };
}

function getEhiFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(EHI_FOLDER_PROPERTY);

  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) {}
  }

  const folders = DriveApp.getFoldersByName("VOID X VPN EHI FILES");
  const folder = folders.hasNext()
    ? folders.next()
    : DriveApp.createFolder("VOID X VPN EHI FILES");

  props.setProperty(EHI_FOLDER_PROPERTY, folder.getId());
  return folder;
}

function getEhiDownloadUrl_(fileId) {
  if (!fileId) return "";
  return "https://drive.google.com/uc?export=download&id=" + encodeURIComponent(fileId);
}

/* ==============================================
   14. DATABASE HELPERS
============================================== */

function readRows_(name) {
  const sheet = getSheet_(name);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(value => String(value).trim());

  return values.slice(1)
    .filter(row => row.some(value => value !== ""))
    .map(row => {
      const item = {};
      headers.forEach((header, index) => { item[header] = row[index]; });
      return item;
    });
}

function findRow_(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return null;

  const headers = getHeaders_(sheet);
  const idColumn = headers.indexOf("id") + 1;
  if (idColumn < 1) return null;

  const ids = sheet.getRange(2, idColumn, sheet.getLastRow() - 1, 1).getValues();
  const index = ids.findIndex(row => String(row[0]) === String(id));
  return index === -1 ? null : index + 2;
}

function isActive_(value) {
  return value === true || String(value).toLowerCase() === "true";
}

function clean_(value) {
  return String(value == null ? "" : value).trim();
}

/* ==============================================
   15. VALIDATION
============================================== */

function validatePackage_(data) {
  if (!data) throw new Error("Package information is missing.");

  const networks = ["DIALOG", "MOBITEL", "HUTCH", "AIRTEL"];
  const accessTypes = ["FREE", "PAID"];

  const network = clean_(data.network).toUpperCase();
  const accessType = clean_(data.accessType).toUpperCase();

  if (!networks.includes(network)) throw new Error("Invalid network.");
  if (!accessTypes.includes(accessType)) throw new Error("Invalid package category.");
  if (!clean_(data.name)) throw new Error("Please enter a package name.");

  if (accessType === "PAID") {
    const amount = clean_(data.dataAmount);
    const price = Number(data.price);

    if (!amount) throw new Error("Please enter Data Amount or Unlimited.");
    if (!Number.isFinite(price) || price <= 0) {
      throw new Error("Please enter a valid package price.");
    }
  }
}

function validateConfig_(data) {
  if (!data) throw new Error("Config information is missing.");
  if (!clean_(data.packageId)) throw new Error("Please select a package.");
  if (!["CODE", "EHI"].includes(clean_(data.type).toUpperCase())) {
    throw new Error("Config type must be CODE or EHI.");
  }
  if (!clean_(data.name) || !String(data.content || "").trim()) {
    throw new Error("Config name and content are required.");
  }
}
