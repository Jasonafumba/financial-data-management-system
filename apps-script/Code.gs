/**
 * ============================================================================
 *  MSA FINANCE — DEMO VERSION (Portfolio showcase)
 *  Simplified data model. No file uploads. No real PDF/email generation.
 *  No real authentication — every visitor can explore every view.
 *  DO NOT reuse this file on the real production project.
 *  File: Code.gs
 * ============================================================================
 *  Only ONE id to fill in below.
 * ============================================================================
 */

const CONFIG = {
  SPREADSHEET_ID: "19eVYe8K-DkLz7Sa0DZUcaF6YeRObN0o8caQSPz3dArY",
  SHEETS: {
    ENTREES: "ENTREES",
    SORTIES: "SORTIES",
    PARAMETRES: "PARAMETRES",
    UTILISATEURS: "UTILISATEURS",
    MEMBRES: "MEMBRES"
  }
};

const ORG_NAME_FALLBACK = "Demo Student Association";

/* ============================================================================
 *  ENTREES — 14 columns
 * ========================================================================= */
const ENTREES_HEADERS = [
  "ID_PAYMENT", "DATE_RECORDED", "MEMBER_ID", "NAME", "UNIVERSITY",
  "EMAIL", "REASON", "AMOUNT", "CURRENCY", "TRANSACTION_NUMBER",
  "PAYMENT_DATE", "STATUS", "PROCESSED_BY", "PROCESSED_DATE"
];
const ENTREES_COL = {};
ENTREES_HEADERS.forEach(function (h, i) { ENTREES_COL[h] = i + 1; });

const SORTIES_HEADERS = [
  "ID_EXPENSE", "DATE_RECORDED", "EXPENSE_DATE", "CATEGORY",
  "AMOUNT", "CURRENCY", "TRANSACTION_NUMBER"
];
const SORTIES_COL = {};
SORTIES_HEADERS.forEach(function (h, i) { SORTIES_COL[h] = i + 1; });

const MEMBRES_HEADERS = ["MEMBER_ID", "NAME", "EMAIL", "UNIVERSITY", "DATE_CREATED", "STATUS"];
const MEMBRES_COL = {};
MEMBRES_HEADERS.forEach(function (h, i) { MEMBRES_COL[h] = i + 1; });

const UTILISATEURS_HEADERS = ["EMAIL", "ROLE", "STATUS", "NAME"]; // cosmetic only in demo mode

const PARAMETRES_HEADERS = [
  "REASONS", "UNIVERSITIES", "EXPENSE_CATEGORIES", "CURRENCIES",
  "RESERVED", "LOGO_URL", "ORGANIZATION"
];

/* ============================================================================
 *  SETUP — creates the 5 sheets with correct headers if missing
 * ========================================================================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("MSA Finance (Demo)")
    .addItem("Initialize sheets", "setupSheets")
    .addItem("Show web app URL", "showAppUrl_")
    .addToUi();
}

function showAppUrl_() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? ("Deployed app: " + url) : "Deploy the app first (Deploy > Web app).");
}

function setupSheets() {
  var ss = getSS_();
  createSheetIfMissing_(ss, CONFIG.SHEETS.ENTREES, ENTREES_HEADERS);
  createSheetIfMissing_(ss, CONFIG.SHEETS.SORTIES, SORTIES_HEADERS);
  createSheetIfMissing_(ss, CONFIG.SHEETS.MEMBRES, MEMBRES_HEADERS);
  createSheetIfMissing_(ss, CONFIG.SHEETS.UTILISATEURS, UTILISATEURS_HEADERS);
  createSheetIfMissing_(ss, CONFIG.SHEETS.PARAMETRES, PARAMETRES_HEADERS);
  SpreadsheetApp.getUi().alert("The 5 sheets have been created / verified.");
}

function createSheetIfMissing_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  var range = sheet.getRange(1, 1, 1, headers.length);
  if (range.getValues()[0].join("") === "") {
    range.setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0B3D91").setFontColor("#FFFFFF");
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

/* ============================================================================
 *  WEB SERVER
 * ========================================================================= */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile("Index");
  template.orgName = getOrgName_();
  return template
    .evaluate()
    .setTitle(template.orgName + " — Finance Demo")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================================
 *  ACCESS CONTROL — DEMO MODE: always allowed, no real authentication.
 *  NEVER reuse this bypass on the real production project.
 * ========================================================================= */
function requireRole_(allowedRoles) {
  return { email: "demo-visitor", role: "ADMIN" };
}

/* ============================================================================
 *  SHEET HELPERS
 * ========================================================================= */
function getSS_() {
  if (!CONFIG.SPREADSHEET_ID) throw new Error("CONFIG.SPREADSHEET_ID is not set.");
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet_(name) {
  var sheet = getSS_().getSheetByName(name);
  if (!sheet) throw new Error("Sheet not found: " + name);
  return sheet;
}

function sheetRowsToObjects_(sheet, headers) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var obj = {};
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = values[i][c];
    obj._row = i + 2;
    out.push(obj);
  }
  return out;
}

function formatDate_(date) {
  if (!date) return "";
  if (!(date instanceof Date)) date = new Date(date);
  if (isNaN(date.getTime())) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone() || "Etc/UTC", "MMM d, yyyy");
}

/* ============================================================================
 *  PARAMETERS (dropdowns) — public
 * ========================================================================= */
function getOrgName_() {
  try {
    var sheet = getSheet_(CONFIG.SHEETS.PARAMETRES);
    var idx = PARAMETRES_HEADERS.indexOf("ORGANIZATION");
    var v = sheet.getLastRow() >= 2 ? sheet.getRange(2, idx + 1).getValue() : "";
    return v ? String(v).trim() : ORG_NAME_FALLBACK;
  } catch (e) {
    return ORG_NAME_FALLBACK;
  }
}

function getParams() {
  var sheet = getSheet_(CONFIG.SHEETS.PARAMETRES);
  var lastRow = sheet.getLastRow();
  var result = {
    REASONS: [], UNIVERSITIES: [], EXPENSE_CATEGORIES: [], CURRENCIES: [],
    LOGO_URL: "", ORGANIZATION: ORG_NAME_FALLBACK
  };
  if (lastRow < 2) return result;

  var lastCol = PARAMETRES_HEADERS.length;
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var listCols = ["REASONS", "UNIVERSITIES", "EXPENSE_CATEGORIES", "CURRENCIES"];

  listCols.forEach(function (col) {
    var colIndex = PARAMETRES_HEADERS.indexOf(col);
    var set = [];
    for (var i = 0; i < values.length; i++) {
      var v = values[i][colIndex];
      if (v !== "" && v !== null && v !== undefined) set.push(String(v).trim());
    }
    result[col] = set;
  });

  var logoIdx = PARAMETRES_HEADERS.indexOf("LOGO_URL");
  var orgIdx = PARAMETRES_HEADERS.indexOf("ORGANIZATION");
  if (values[0][logoIdx]) result.LOGO_URL = toDisplayableImageUrl_(String(values[0][logoIdx]).trim());
  if (values[0][orgIdx]) result.ORGANIZATION = String(values[0][orgIdx]).trim();

  return result;
}

function extractDriveFileId_(url) {
  if (!url) return "";
  var patterns = [/\/d\/([-\w]{20,})/, /[?&]id=([-\w]{20,})/, /\/folders\/([-\w]{20,})/];
  for (var i = 0; i < patterns.length; i++) {
    var m = url.match(patterns[i]);
    if (m && m[1]) return m[1];
  }
  return "";
}

function toDisplayableImageUrl_(url) {
  if (!url) return "";
  url = String(url).trim();
  if (url.indexOf("drive.google.com") === -1) return url;
  var id = extractDriveFileId_(url);
  if (!id) return url;
  return "https://drive.google.com/thumbnail?id=" + id + "&sz=w300";
}

/* ============================================================================
 *  MEMBERS — automatic identification / creation
 * ========================================================================= */
function findOrCreateMember_(data) {
  if (!data.email) return "";

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_(CONFIG.SHEETS.MEMBRES);
    var rows = sheetRowsToObjects_(sheet, MEMBRES_HEADERS);
    var email = String(data.email).toLowerCase().trim();

    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].EMAIL || "").toLowerCase().trim() === email) return rows[i].MEMBER_ID;
    }

    var nextNum = rows.length + 1;
    var id = "M" + Utilities.formatString("%04d", nextNum);
    var existingIds = rows.map(function (r) { return r.MEMBER_ID; });
    while (existingIds.indexOf(id) !== -1) {
      nextNum++;
      id = "M" + Utilities.formatString("%04d", nextNum);
    }

    sheet.appendRow([id, data.name || "", data.email || "", data.university || "", new Date(), "ACTIVE"]);
    return id;
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================================
 *  PAYMENTS — member submission (no file upload; any field may be empty)
 * ========================================================================= */
function generatePaymentId_(sheet) {
  var lastRow = sheet.getLastRow();
  var nextNum = lastRow < 2 ? 1 : lastRow;
  var id = "PAY-" + Utilities.formatString("%08d", nextNum);
  var existing = lastRow < 2 ? [] : sheet.getRange(2, ENTREES_COL.ID_PAYMENT, lastRow - 1, 1).getValues().map(function (r) { return r[0]; });
  while (existing.indexOf(id) !== -1) {
    nextNum++;
    id = "PAY-" + Utilities.formatString("%08d", nextNum);
  }
  return id;
}

function submitPayment(data) {
  data = data || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var idPayment;
  try {
    var memberId = findOrCreateMember_(data);
    var sheet = getSheet_(CONFIG.SHEETS.ENTREES);
    idPayment = generatePaymentId_(sheet);

    var row = [];
    row[ENTREES_COL.ID_PAYMENT - 1] = idPayment;
    row[ENTREES_COL.DATE_RECORDED - 1] = new Date();
    row[ENTREES_COL.MEMBER_ID - 1] = memberId;
    row[ENTREES_COL.NAME - 1] = data.name || "";
    row[ENTREES_COL.UNIVERSITY - 1] = data.university || "";
    row[ENTREES_COL.EMAIL - 1] = data.email || "";
    row[ENTREES_COL.REASON - 1] = data.reason || "";
    row[ENTREES_COL.AMOUNT - 1] = Number(data.amount) || 0;
    row[ENTREES_COL.CURRENCY - 1] = data.currency || "";
    row[ENTREES_COL.TRANSACTION_NUMBER - 1] = data.transactionNumber || "";
    row[ENTREES_COL.PAYMENT_DATE - 1] = data.paymentDate ? new Date(data.paymentDate) : new Date();
    row[ENTREES_COL.STATUS - 1] = "PENDING";
    row[ENTREES_COL.PROCESSED_BY - 1] = "";
    row[ENTREES_COL.PROCESSED_DATE - 1] = "";

    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { success: true, idPayment: idPayment };
}

/* ============================================================================
 *  PENDING PAYMENTS — review
 * ========================================================================= */
function getPendingPayments() {
  requireRole_(["ADMIN"]);
  var sheet = getSheet_(CONFIG.SHEETS.ENTREES);
  var rows = sheetRowsToObjects_(sheet, ENTREES_HEADERS);
  return rows
    .filter(function (r) { return String(r.STATUS).toUpperCase().trim() === "PENDING"; })
    .map(function (r) {
      return {
        idPayment: r.ID_PAYMENT,
        name: r.NAME,
        university: r.UNIVERSITY,
        reason: r.REASON,
        amount: r.AMOUNT,
        currency: r.CURRENCY,
        transactionNumber: r.TRANSACTION_NUMBER,
        paymentDate: formatDate_(r.PAYMENT_DATE),
        email: r.EMAIL,
        memberId: r.MEMBER_ID
      };
    })
    .reverse();
}

function findEntreeRowById_(sheet, idPayment) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, ENTREES_COL.ID_PAYMENT, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === idPayment) return i + 2;
  }
  return -1;
}

/**
 * Approves a payment. No PDF, no email — returns the receipt data so the
 * client can render an on-screen receipt preview immediately.
 */
function approvePayment(idPayment) {
  requireRole_(["ADMIN"]);
  var sheet = getSheet_(CONFIG.SHEETS.ENTREES);
  var rowNum = findEntreeRowById_(sheet, idPayment);
  if (rowNum === -1) throw new Error("Payment not found: " + idPayment);

  var now = new Date();
  sheet.getRange(rowNum, ENTREES_COL.STATUS).setValue("APPROVED");
  sheet.getRange(rowNum, ENTREES_COL.PROCESSED_BY).setValue("Demo Admin");
  sheet.getRange(rowNum, ENTREES_COL.PROCESSED_DATE).setValue(now);

  var rowValues = sheet.getRange(rowNum, 1, 1, ENTREES_HEADERS.length).getValues()[0];
  var r = {};
  ENTREES_HEADERS.forEach(function (h, i) { r[h] = rowValues[i]; });

  return {
    success: true,
    idPayment: idPayment,
    receipt: {
      idPayment: r.ID_PAYMENT,
      name: r.NAME,
      university: r.UNIVERSITY,
      reason: r.REASON,
      amount: r.AMOUNT,
      currency: r.CURRENCY,
      transactionNumber: r.TRANSACTION_NUMBER,
      paymentDate: formatDate_(r.PAYMENT_DATE),
      processedBy: r.PROCESSED_BY,
      processedDate: formatDate_(r.PROCESSED_DATE)
    }
  };
}

function rejectPayment(idPayment) {
  requireRole_(["ADMIN"]);
  var sheet = getSheet_(CONFIG.SHEETS.ENTREES);
  var rowNum = findEntreeRowById_(sheet, idPayment);
  if (rowNum === -1) throw new Error("Payment not found: " + idPayment);

  sheet.getRange(rowNum, ENTREES_COL.STATUS).setValue("REJECTED");
  sheet.getRange(rowNum, ENTREES_COL.PROCESSED_BY).setValue("Demo Admin");
  sheet.getRange(rowNum, ENTREES_COL.PROCESSED_DATE).setValue(new Date());

  return { success: true, idPayment: idPayment };
}

/* ============================================================================
 *  EXPENSES
 * ========================================================================= */
function generateExpenseId_(sheet) {
  var lastRow = sheet.getLastRow();
  var nextNum = lastRow < 2 ? 1 : lastRow;
  var id = "EXP-" + Utilities.formatString("%08d", nextNum);
  var existing = lastRow < 2 ? [] : sheet.getRange(2, SORTIES_COL.ID_EXPENSE, lastRow - 1, 1).getValues().map(function (r) { return r[0]; });
  while (existing.indexOf(id) !== -1) {
    nextNum++;
    id = "EXP-" + Utilities.formatString("%08d", nextNum);
  }
  return id;
}

function recordExpense(data) {
  requireRole_(["ADMIN"]);
  data = data || {};

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var idExpense;
  try {
    var sheet = getSheet_(CONFIG.SHEETS.SORTIES);
    idExpense = generateExpenseId_(sheet);

    var row = [];
    row[SORTIES_COL.ID_EXPENSE - 1] = idExpense;
    row[SORTIES_COL.DATE_RECORDED - 1] = new Date();
    row[SORTIES_COL.EXPENSE_DATE - 1] = data.expenseDate ? new Date(data.expenseDate) : new Date();
    row[SORTIES_COL.CATEGORY - 1] = data.category || "";
    row[SORTIES_COL.AMOUNT - 1] = Number(data.amount) || 0;
    row[SORTIES_COL.CURRENCY - 1] = data.currency || "";
    row[SORTIES_COL.TRANSACTION_NUMBER - 1] = data.transactionNumber || "";

    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { success: true, idExpense: idExpense };
}

/* ============================================================================
 *  DASHBOARD
 * ========================================================================= */
function withinRange_(date, start, end) {
  if (!date) return !start && !end;
  var dt = date instanceof Date ? date : new Date(date);
  if (start && dt < new Date(start)) return false;
  if (end) {
    var endDate = new Date(end);
    endDate.setHours(23, 59, 59, 999);
    if (dt > endDate) return false;
  }
  return true;
}

function monthKey_(date) {
  var dt = date instanceof Date ? date : new Date(date);
  if (isNaN(dt.getTime())) return "N/A";
  return Utilities.formatDate(dt, Session.getScriptTimeZone() || "Etc/UTC", "yyyy-MM");
}

function getDashboardData(filters) {
  requireRole_(["ADMIN"]);
  filters = filters || {};

  var entriesSheet = getSheet_(CONFIG.SHEETS.ENTREES);
  var expensesSheet = getSheet_(CONFIG.SHEETS.SORTIES);

  var entries = sheetRowsToObjects_(entriesSheet, ENTREES_HEADERS).filter(function (r) {
    if (String(r.STATUS).toUpperCase().trim() !== "APPROVED") return false;
    if (filters.reason && r.REASON !== filters.reason) return false;
    if (filters.university && r.UNIVERSITY !== filters.university) return false;
    if (filters.currency && r.CURRENCY !== filters.currency) return false;
    if (!withinRange_(r.PAYMENT_DATE, filters.dateFrom, filters.dateTo)) return false;
    return true;
  });

  var expenses = sheetRowsToObjects_(expensesSheet, SORTIES_HEADERS).filter(function (r) {
    if (filters.currency && r.CURRENCY !== filters.currency) return false;
    if (!withinRange_(r.EXPENSE_DATE, filters.dateFrom, filters.dateTo)) return false;
    return true;
  });

  var totalIncome = entries.reduce(function (s, r) { return s + (Number(r.AMOUNT) || 0); }, 0);
  var totalExpenses = expenses.reduce(function (s, r) { return s + (Number(r.AMOUNT) || 0); }, 0);

  var monthlyIncome = {}, monthlyExpenses = {};
  entries.forEach(function (r) {
    var k = monthKey_(r.PAYMENT_DATE);
    monthlyIncome[k] = (monthlyIncome[k] || 0) + (Number(r.AMOUNT) || 0);
  });
  expenses.forEach(function (r) {
    var k = monthKey_(r.EXPENSE_DATE);
    monthlyExpenses[k] = (monthlyExpenses[k] || 0) + (Number(r.AMOUNT) || 0);
  });

  var reasonBreakdown = {};
  entries.forEach(function (r) {
    var m = r.REASON || "Other";
    reasonBreakdown[m] = (reasonBreakdown[m] || 0) + (Number(r.AMOUNT) || 0);
  });

  var months = {};
  Object.keys(monthlyIncome).forEach(function (k) { months[k] = true; });
  Object.keys(monthlyExpenses).forEach(function (k) { months[k] = true; });
  var sortedMonths = Object.keys(months).sort();

  return {
    kpi: {
      totalIncome: totalIncome,
      totalExpenses: totalExpenses,
      balance: totalIncome - totalExpenses,
      approvedCount: entries.length,
      expenseCount: expenses.length
    },
    monthly: {
      labels: sortedMonths,
      income: sortedMonths.map(function (m) { return monthlyIncome[m] || 0; }),
      expenses: sortedMonths.map(function (m) { return monthlyExpenses[m] || 0; })
    },
    reasonBreakdown: {
      labels: Object.keys(reasonBreakdown),
      values: Object.keys(reasonBreakdown).map(function (k) { return reasonBreakdown[k]; })
    },
    comparison: { totalIncome: totalIncome, totalExpenses: totalExpenses }
  };
}
