/*************************************************
 * VOID X VPN™ - GOOGLE APPS SCRIPT BACKEND (API)
 * Public site: Vercel (public/index.html) | Admin: this script (?page=admin)
 * Database: Google Sheets
 *
 * Deploy as Web app:  Execute as: Me | Who has access: Anyone
 * Admin credentials go in Project Settings > Script Properties:
 *   ADMIN_USERNAME, ADMIN_PASSWORD
 *************************************************/

const SHEET_PACKAGES = "Packages";
const SHEET_CONFIGS = "Configs";
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

/* ==============================================
   1. API ROUTING (used by the Vercel website)
============================================== */

// Only these functions can be called from the website.
const API_METHODS = {
  getPublicData, adminLogin, adminLogout, getAdminData,
  addPackage, updatePackage, deletePackage,
  addConfig, updateConfig, deleteConfig
};

function doGet(e) {
  // Admin dashboard is served by Apps Script itself: .../exec?page=admin
  if (e && e.parameter && e.parameter.page === "admin") {
    return HtmlService.createHtmlOutputFromFile("Admin")
      .setTitle("VOID X VPN™ | Admin Dashboard")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  return json_({ ok: true, service: "VOID X VPN API", status: "online" });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const fn = API_METHODS[body.method];
    if (!fn) throw new Error("Unknown method.");
    return json_({ ok: true, result: fn.apply(null, body.args || []) });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ==============================================
   2. DATABASE SETUP + SAFE MIGRATION
   (run setupDatabase once from the editor)
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

  props.setProperty("SPREADSHEET_ID", ss.getId());

  return "VOID X VPN database is ready! Existing data has been preserved.";
}

function createSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

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

  if (id) {
    return SpreadsheetApp.openById(id);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error("Database not connected. Run setupDatabase() first.");
  }

  PropertiesService.getScriptProperties()
    .setProperty("SPREADSHEET_ID", ss.getId());

  return ss;
}

function getSheet_(name) {
  const sheet = getDatabase_().getSheetByName(name);

  if (!sheet) {
    throw new Error(
      "Database sheet missing: " + name + ". Run setupDatabase() first."
    );
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
    Object.prototype.hasOwnProperty.call(data, header)
      ? data[header]
      : ""
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
   4. ADMIN LOGIN (with basic brute-force protection)
============================================== */

const MAX_FAILED_LOGINS = 5;
const LOCKOUT_SECONDS = 900; // 15 minutes

function adminLogin(username, password) {
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();

  const failures = Number(cache.get("VOID_LOGIN_FAILS") || 0);

  if (failures >= MAX_FAILED_LOGINS) {
    return {
      success: false,
      message: "Too many failed attempts. Try again in 15 minutes."
    };
  }

  const savedUsername = props.getProperty("ADMIN_USERNAME");
  const savedPassword = props.getProperty("ADMIN_PASSWORD");

  if (!savedUsername || !savedPassword) {
    throw new Error(
      "Admin credentials are not configured in Script Properties."
    );
  }

  if (
    String(username) !== savedUsername ||
    String(password) !== savedPassword
  ) {
    cache.put("VOID_LOGIN_FAILS", String(failures + 1), LOCKOUT_SECONDS);

    return {
      success: false,
      message: "Invalid username or password."
    };
  }

  cache.remove("VOID_LOGIN_FAILS");

  const token = Utilities.getUuid() + Utilities.getUuid();

  cache.put(SESSION_PREFIX + token, "authorized", SESSION_SECONDS);

  return {
    success: true,
    token: token,
    message: "Login successful."
  };
}

function verifyAdmin_(token) {
  if (!token) {
    throw new Error("Admin login required.");
  }

  const session = CacheService.getScriptCache()
    .get(SESSION_PREFIX + token);

  if (session !== "authorized") {
    throw new Error("Session expired. Please login again.");
  }

  return true;
}

function adminLogout(token) {
  if (token) {
    CacheService.getScriptCache()
      .remove(SESSION_PREFIX + token);
  }

  return { success: true };
}

/* ==============================================
   5. PUBLIC WEBSITE DATA
   PAID CONFIG CONTENT IS NEVER PUBLIC
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

  // Only active FREE packages can expose configs publicly.
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

  return {
    packages: packages,
    configs: configs
  };
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

  return {
    packages: packages,
    configs: configs
  };
}

function safeDate_(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }

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
    price: clean_(data.accessType).toUpperCase() === "PAID"
      ? Number(data.price)
      : "",
    dataAmount: clean_(data.accessType).toUpperCase() === "PAID"
      ? clean_(data.dataAmount)
      : ""
  });

  return {
    success: true,
    id: id
  };
}

/* ==============================================
   8. UPDATE PACKAGE
============================================== */

function updatePackage(token, data) {
  verifyAdmin_(token);
  validatePackage_(data);

  const sheet = getSheet_(SHEET_PACKAGES);
  const row = findRow_(sheet, data.id);

  if (!row) {
    throw new Error("Package not found.");
  }

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

  if (!packageRow) {
    throw new Error("Package not found.");
  }

  const configHeaders = getHeaders_(configSheet);
  const packageIdColumn = configHeaders.indexOf("packageId") + 1;

  if (packageIdColumn > 0 && configSheet.getLastRow() >= 2) {
    const configIds = configSheet.getRange(
      2,
      packageIdColumn,
      configSheet.getLastRow() - 1,
      1
    ).getValues();

    for (let i = configIds.length - 1; i >= 0; i--) {
      if (String(configIds[i][0]) === String(id)) {
        configSheet.deleteRow(i + 2);
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

  const packageSheet = getSheet_(SHEET_PACKAGES);
  const packageRow = findRow_(packageSheet, data.packageId);

  if (!packageRow) {
    throw new Error("Selected package does not exist.");
  }

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

  return {
    success: true,
    id: id
  };
}

/* ==============================================
   11. UPDATE CONFIG
============================================== */

function updateConfig(token, data) {
  verifyAdmin_(token);
  validateConfig_(data);

  const sheet = getSheet_(SHEET_CONFIGS);
  const row = findRow_(sheet, data.id);

  if (!row) {
    throw new Error("Config not found.");
  }

  const packageRow = findRow_(
    getSheet_(SHEET_PACKAGES),
    data.packageId
  );

  if (!packageRow) {
    throw new Error("Selected package does not exist.");
  }

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

  if (!row) {
    throw new Error("Config not found.");
  }

  sheet.deleteRow(row);

  return { success: true };
}

/* ==============================================
   13. DATABASE HELPERS
============================================== */

function readRows_(name) {
  const sheet = getSheet_(name);
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(value => String(value).trim());

  return values.slice(1)
    .filter(row => row.some(value => value !== ""))
    .map(row => {
      const item = {};

      headers.forEach((header, index) => {
        item[header] = row[index];
      });

      return item;
    });
}

function findRow_(sheet, id) {
  if (!id || sheet.getLastRow() < 2) {
    return null;
  }

  const headers = getHeaders_(sheet);
  const idColumn = headers.indexOf("id") + 1;

  if (idColumn < 1) return null;

  const ids = sheet.getRange(
    2,
    idColumn,
    sheet.getLastRow() - 1,
    1
  ).getValues();

  const index = ids.findIndex(row =>
    String(row[0]) === String(id)
  );

  return index === -1 ? null : index + 2;
}

function isActive_(value) {
  return value === true ||
    String(value).toLowerCase() === "true";
}

function clean_(value) {
  return String(value == null ? "" : value).trim();
}

/* ==============================================
   14. VALIDATION
============================================== */

function validatePackage_(data) {
  if (!data) {
    throw new Error("Package information is missing.");
  }

  const networks = ["DIALOG", "MOBITEL", "HUTCH", "AIRTEL"];
  const accessTypes = ["FREE", "PAID"];

  const network = clean_(data.network).toUpperCase();
  const accessType = clean_(data.accessType).toUpperCase();

  if (!networks.includes(network)) {
    throw new Error("Invalid network.");
  }

  if (!accessTypes.includes(accessType)) {
    throw new Error("Invalid package category.");
  }

  if (!clean_(data.name)) {
    throw new Error("Please enter a package name.");
  }

  if (accessType === "PAID") {
    const amount = clean_(data.dataAmount);
    const price = Number(data.price);

    if (!amount) {
      throw new Error("Please enter Data Amount or Unlimited.");
    }

    if (!Number.isFinite(price) || price <= 0) {
      throw new Error("Please enter a valid package price.");
    }
  }
}

function validateConfig_(data) {
  if (!data) {
    throw new Error("Config information is missing.");
  }

  if (!clean_(data.packageId)) {
    throw new Error("Please select a package.");
  }

  if (!["CODE", "EHI"].includes(clean_(data.type).toUpperCase())) {
    throw new Error("Config type must be CODE or EHI.");
  }

  if (!clean_(data.name) || !String(data.content || "").trim()) {
    throw new Error("Config name and content are required.");
  }
}
