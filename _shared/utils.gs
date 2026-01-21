/**
 * SheetLink Recipes - Shared Utilities
 * Common functions used by all recipes
 *
 * This file is automatically injected by the installer when recipes are installed.
 * Community recipe developers should use these utilities to ensure compatibility.
 */

/**
 * Constants
 */
const TRANSACTIONS_SHEET_NAME = "Transactions";

/**
 * Get or create a sheet by name
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {string} sheetName - Name of sheet to get/create
 * @returns {Sheet} The requested sheet
 */
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // Move newly created sheet to the end (furthest right)
    sheet.activate();
    ss.moveActiveSheet(ss.getNumSheets());
    Logger.log(`Created new sheet: ${sheetName}`);
  }
  return sheet;
}

/**
 * Get transactions sheet
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Sheet|null} Transactions sheet or null if not found
 */
function getTransactionsSheet(ss) {
  return ss.getSheetByName(TRANSACTIONS_SHEET_NAME);
}

/**
 * Validate transactions sheet exists and has data
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Object} {valid: boolean, error: string|null}
 */
function validateTransactionsSheet(ss) {
  const sheet = getTransactionsSheet(ss);

  if (!sheet) {
    return {
      valid: false,
      error: `Sheet "${TRANSACTIONS_SHEET_NAME}" not found. Please sync your transactions first.`
    };
  }

  // Check for actual data by reading first data cell (A2)
  // This is more reliable than getLastRow() which can be inconsistent
  try {
    const firstDataCell = sheet.getRange(2, 1).getValue();
    // Convert to string and trim to handle whitespace and non-string types
    const cellValue = String(firstDataCell || '').trim();

    if (!cellValue) {
      return {
        valid: false,
        error: "No transaction data found. Please sync your transactions first."
      };
    }
  } catch (error) {
    return {
      valid: false,
      error: "Could not read transaction data. Please sync your transactions first."
    };
  }

  return { valid: true, error: null };
}

/**
 * Format currency
 * @param {number} value - Value to format
 * @returns {string} Formatted currency string
 */
function formatCurrency(value) {
  return '$' + value.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}

/**
 * Get column index by header name
 * @param {Sheet} sheet - Sheet to search
 * @param {string} headerName - Header name to find
 * @returns {number} Column index (1-based) or -1 if not found
 */
function getColumnIndexByHeader(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().toLowerCase() === headerName.toLowerCase()) {
      return i + 1;
    }
  }
  return -1;
}

/**
 * Parse date string to Date object
 * @param {string|Date} dateStr - Date string or Date object
 * @returns {Date} Parsed date
 */
function parseDate(dateStr) {
  if (dateStr instanceof Date) {
    return dateStr;
  }
  return new Date(dateStr);
}

/**
 * Format date as YYYY-MM-DD
 * @param {Date} date - Date to format
 * @returns {string} Formatted date string
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Get month key from date (YYYY-MM format)
 * @param {Date} date - Date object
 * @returns {string} Month key in YYYY-MM format
 */
function getMonthKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
}

/**
 * Clear sheet contents but preserve headers
 * @param {Sheet} sheet - Sheet to clear
 */
function clearSheetKeepHeaders(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
}
