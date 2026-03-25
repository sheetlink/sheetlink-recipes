/**
 * SheetLink Recipe: Financial Statements
 * Version: 3.0.0
 * Standalone Edition
 *
 * Description: Complete financial reporting suite with P&L, Balance Sheet, and Cash Flow
 *
 * Creates: Category Mapping, General Ledger, Financial Statements
 * Requires: date, amount, category_primary, pending, account_name, transaction_id columns
 */


// RECIPE LOGIC
// ========================================

function runFinancialsRecipe(ss) {
  try {
    // Ensure we have a spreadsheet object
    if (!ss) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }

    // Check for transactions data
    if (!checkTransactionsOrPrompt(ss)) {
      return;
    }

    logRecipe("Financials", "Starting Financial Statements Suite v3.0");
    showToast("Generating financial statements...", "Financial Statements", 3);

    // Get transactions sheet
    const transactionsSheet = getTransactionsSheet(ss);

    if (!transactionsSheet) {
      const errorMsg = "Transactions sheet not found. Please sync your transactions first.";
      showError(errorMsg);
      return { success: false, error: errorMsg };
    }

    const headerMap = getHeaderMap(transactionsSheet);

    // Verify required columns exist
    const requiredColumns = ['date', 'amount', 'merchant_name', 'category_primary', 'account_name', 'pending', 'transaction_id'];
    const missingColumns = [];

    for (const col of requiredColumns) {
      if (!getColumnIndex(headerMap, col)) {
        missingColumns.push(col);
      }
    }

    if (missingColumns.length > 0) {
      const errorMsg = `Missing required columns: ${missingColumns.join(', ')}. Please sync your transactions.`;
      showError(errorMsg);
      return { success: false, error: errorMsg };
    }

    // Format date and pending columns in Transactions sheet
    formatTransactionDateColumns(transactionsSheet, headerMap);
    formatTransactionPendingColumn(transactionsSheet, headerMap);

    // Rename legacy "Chart of Accounts" tab to "Category Mapping" if it exists
    const legacyCoa = ss.getSheetByName("Chart of Accounts");
    if (legacyCoa) {
      legacyCoa.setName("Category Mapping");
    }

    // Create Guide tab and position it directly left of Category Mapping
    const guideSheet = getOrCreateSheet(ss, "Guide");
    setupGuideSheet(guideSheet);
    const mappingSheetForOrder = ss.getSheetByName("Category Mapping");
    if (mappingSheetForOrder) {
      ss.setActiveSheet(guideSheet);
      ss.moveActiveSheet(mappingSheetForOrder.getIndex());
    }

    // Create or get Category Mapping sheet
    const mappingSheet = getOrCreateSheet(ss, "Category Mapping");
    setupCategoryMapping(mappingSheet, transactionsSheet, headerMap, ss);

    // Create General Ledger
    const ledgerSheet = getOrCreateSheet(ss, "General Ledger");
    setupGeneralLedgerV3(ledgerSheet, transactionsSheet, headerMap, mappingSheet, ss);

    // Create consolidated Financial Statements
    const statementsSheet = getOrCreateSheet(ss, "Financial Statements");
    setupFinancialStatementsV3(statementsSheet, ledgerSheet, transactionsSheet, headerMap, ss);

    showToast("Financial statements generated successfully!", "Complete", 5);
    logRecipe("Financials", "Recipe v3.0 completed successfully");
    return { success: true, error: null };

  } catch (error) {
    Logger.log(`Financials recipe error: ${error.message}`);
    Logger.log(error.stack);
    showError(`Error generating financial statements: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/**
 * Setup Category Mapping tab
 * Maps Plaid categories to user-defined custom categories, type, and statement.
 * Preserves existing user edits on re-run.
 */
function setupGuideSheet(sheet) {
  sheet.clear();
  sheet.clearFormats();

  const green = "#0b703a";
  const gray = "#6b7280";
  const bodyColor = "#374151";

  sheet.setColumnWidth(1, 25);
  sheet.setColumnWidth(2, 230);
  sheet.setColumnWidth(3, 700);
  sheet.setHiddenGridlines(true);

  // Build all content as value/format pairs — no nested functions
  // Each entry: [row, col, value, {bold, fg, bg, size, wrap}]
  const entries = [
    // Title
    [1, 2, "Financial Statements — Guide", {size: 20, bold: true, fg: green}],
    [2, 2, "A starting point for financial model building. Customize categories, override transactions, and build on top of what the recipe generates. Happy building!", {fg: gray}],

    // Section 1
    [4, 2, "1.  What This Recipe Creates", {bold: true, fg: "white", bg: green, size: 11}],
    [4, 3, "", {bg: green}],
    [6, 2, "Guide", {bold: true}],
    [6, 3, "This tab — overview and instructions.", {fg: gray}],
    [7, 2, "Category Mapping", {bold: true}],
    [7, 3, "Maps Plaid categories to your custom categories, type, and statement.", {fg: gray}],
    [8, 2, "General Ledger", {bold: true}],
    [8, 3, "Full transaction history in debit/credit accounting format.", {fg: gray}],
    [9, 2, "Financial Statements", {bold: true}],
    [9, 3, "Monthly P&L, Balance Sheet, and Cash Flow — auto-built from the GL.", {fg: gray}],

    // Section 2
    [11, 2, "2.  Required Transaction Columns", {bold: true, fg: "white", bg: green, size: 11}],
    [11, 3, "", {bg: green}],
    [13, 2, "Your Transactions sheet must contain: date, amount, category_primary, pending, account_name, transaction_id, merchant_name. These are synced automatically by SheetLink.", {fg: bodyColor, wrap: true}],

    // Section 3
    [15, 2, "3.  Category Mapping", {bold: true, fg: "white", bg: green, size: 11}],
    [15, 3, "", {bg: green}],
    [17, 2, "Plaid categories are auto-mapped to sensible defaults. Edit the Custom Category, Type, or Statement columns (yellow) to customize how transactions appear in your Financial Statements.", {fg: bodyColor, wrap: true}],
    [19, 2, "Statement options", {bold: true}],
    [19, 3, "P&L · Balance Sheet", {fg: gray}],
    [20, 2, "P&L Type options", {bold: true}],
    [20, 3, "Revenue or Expense. P&L transactions appear as line items in the income statement.", {fg: gray}],
    [21, 2, "Balance Sheet Type", {bold: true}],
    [21, 3, "Any value (or leave blank). Balance Sheet transactions net out into a running account balance — Type is not tracked, only the account's position in Assets or Liabilities (set in the GL config).", {fg: gray, wrap: true}],
    [23, 2, "To add a custom category not tied to a Plaid category, add a new row with col A blank and fill in Custom Category, Type, and Statement. Re-run to apply.", {fg: bodyColor, wrap: true}],

    // Section 4
    [25, 2, "4.  Manual Overrides in the General Ledger", {bold: true, fg: "white", bg: green, size: 11}],
    [25, 3, "", {bg: green}],
    [27, 2, "The Custom Category, Type, and Statement columns (yellow) in the General Ledger are formula-driven from Category Mapping. You can override any cell directly — overrides persist when re-run.", {fg: bodyColor, wrap: true}],
    [29, 2, "To reset a cell back to auto-mapping, clear it and re-run the recipe.", {fg: bodyColor}],

    // Section 5
    [31, 2, "5.  Re-Running the Recipe", {bold: true, fg: "white", bg: green, size: 11}],
    [31, 3, "", {bg: green}],
    [33, 2, "Run from SheetLink Recipes → Financial Statements. Safe to re-run — Category Mapping edits and GL manual overrides are preserved. The Financial Statements tab is fully rebuilt each run.", {fg: bodyColor, wrap: true}],

    // Section 6
    [35, 2, "6.  Tips", {bold: true, fg: "white", bg: green, size: 11}],
    [35, 3, "", {bg: green}],
    [37, 2, "Starting Balance", {bold: true}],
    [37, 3, "Set a Starting Balance and As of Date per account in the GL's Account Balance Configuration table for accurate Balance Sheet history.", {fg: gray, wrap: true}],
    [39, 2, "Account Type", {bold: true}],
    [39, 3, "Accounts are auto-detected as Asset or Liability. Review and correct in the GL config table if needed.", {fg: gray, wrap: true}],
    [41, 2, "Cash Flow", {bold: true}],
    [41, 3, "Uses the indirect method. Capital Expenditures and Loan Proceeds are manual input — edit directly in the Financial Statements tab.", {fg: gray, wrap: true}],

    // Section 7
    [43, 2, "7.  Troubleshooting", {bold: true, fg: "white", bg: green, size: 11}],
    [43, 3, "", {bg: green}],
    [45, 2, "$0 on Financial Statements", {bold: true}],
    [45, 3, "Check that GL rows have Custom Category, Type, and Statement populated. Re-run the recipe.", {fg: gray}],
    [46, 2, "Category missing from statements", {bold: true}],
    [46, 3, "Ensure the category exists in the GL (col E) with Type = Expense/Revenue and Statement = P&L. Re-run.", {fg: gray}],
    [47, 2, "Unexpected values after edits", {bold: true}],
    [47, 3, "Re-run the recipe to rebuild Financial Statements from the latest GL data.", {fg: gray}],
  ];

  entries.forEach(function(entry) {
    var row = entry[0], col = entry[1], value = entry[2], fmt = entry[3];
    var cell = sheet.getRange(row, col);
    if (value !== "") cell.setValue(value);
    if (fmt.bg) cell.setBackground(fmt.bg);
    if (fmt.fg) cell.setFontColor(fmt.fg);
    if (fmt.bold) cell.setFontWeight("bold");
    if (fmt.size) cell.setFontSize(fmt.size);
    if (fmt.wrap) cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  });

  // Merge body text cells across B and C
  [2, 13, 17, 23, 27, 33].forEach(function(r) {
    sheet.getRange(r, 2, 1, 2).merge();
  });

  // Vertical align top for tip rows and category mapping options
  [19, 20, 21, 37, 39, 41].forEach(function(r) {
    sheet.getRange(r, 2).setVerticalAlignment("top");
  });

  sheet.setFrozenRows(2);
}

function setupCategoryMapping(sheet, transactionsSheet, headerMap, ss) {
  // Row 1: Title
  sheet.getRange("A1")
    .setValue("Category Mapping")
    .setFontSize(18)
    .setFontWeight("bold")
    .setFontColor("#023820")
    .setHorizontalAlignment("left");

  // Row 2: Description
  sheet.getRange("A2")
    .setValue("Maps Plaid transaction categories to your custom categories, type, and statement.")
    .setFontSize(11)
    .setWrap(false)
    .setHorizontalAlignment("left");

  sheet.setRowHeight(3, 21);

  // Row 4: Column headers
  const headers = ["Plaid Category", "Custom Category", "Type", "Statement"];
  sheet.getRange(4, 1, 1, 4).setValues([headers]);
  sheet.getRange(4, 1, 1, 4)
    .setFontWeight("bold")
    .setBackground("#0b703a")
    .setFontColor("white");

  // Default mappings: Plaid Category -> [Custom Category, Type, Statement]
  const defaultMappings = {
    "INCOME":                    ["Income",                    "Revenue",  "P&L"],
    "TRANSFER_IN":               ["Transfers In",              "Transfer", "Balance Sheet"],
    "TRANSFER_OUT":              ["Transfers Out",             "Transfer", "Balance Sheet"],
    "LOAN_DISBURSEMENTS":        ["Transfers In",              "Transfer", "Balance Sheet"],
    "FOOD_AND_DRINK":            ["Meals & Entertainment",     "Expense",  "P&L"],
    "GENERAL_MERCHANDISE":       ["General Merchandise",       "Expense",  "P&L"],
    "GENERAL_SERVICES":          ["Services",                  "Expense",  "P&L"],
    "ENTERTAINMENT":             ["Entertainment",             "Expense",  "P&L"],
    "TRANSPORTATION":            ["Transportation",            "Expense",  "P&L"],
    "TRAVEL":                    ["Travel",                    "Expense",  "P&L"],
    "RENT_AND_UTILITIES":        ["Rent & Utilities",          "Expense",  "P&L"],
    "HOME_IMPROVEMENT":          ["Home Improvement",          "Expense",  "P&L"],
    "MEDICAL":                   ["Healthcare",                "Expense",  "P&L"],
    "PERSONAL_CARE":             ["Personal Care",             "Expense",  "P&L"],
    "LOAN_PAYMENTS":             ["Interest & Loan Payments",  "Expense",  "P&L"],
    "BANK_FEES":                 ["Bank Fees",                 "Expense",  "P&L"],
    "GOVERNMENT_AND_NON_PROFIT": ["Taxes & Government",        "Expense",  "P&L"],
    "OTHER":                     ["Other Expenses",            "Expense",  "P&L"],
    "BANK_ACCOUNT":              ["Cash - Checking",           "Asset",    "Balance Sheet"],
    "CREDIT_CARD":               ["Credit Card Payable",       "Liability","Balance Sheet"],
    "LOAN":                      ["Loans Payable",             "Liability","Balance Sheet"]
  };

  // Scan Transactions for all unique Plaid categories
  const categoryCol = getColumnIndex(headerMap, 'category_primary');
  const transactionsData = transactionsSheet.getDataRange().getValues();
  const uniqueCategories = new Set();
  for (let i = 1; i < transactionsData.length; i++) {
    const category = transactionsData[i][categoryCol - 1];
    if (category && category !== "") uniqueCategories.add(category);
  }

  // Read existing user mappings to preserve edits
  // New column order: Plaid Category (A) | Custom Category (B) | Type (C) | Statement (D)
  const existingMappings = {};
  const customOnlyRows = []; // Rows with no Plaid category (user-added custom categories)
  if (sheet.getLastRow() >= 5) {
    try {
      const existingData = sheet.getRange(5, 1, sheet.getLastRow() - 4, 4).getValues();
      existingData.forEach(row => {
        const plaidCategory = row[0];
        const customCategory = row[1];
        if (plaidCategory && plaidCategory !== "") {
          existingMappings[plaidCategory] = {
            customCategory: row[1],
            type: row[2],
            statement: row[3]
          };
        } else if (customCategory && customCategory !== "") {
          // Preserve custom-only rows (no Plaid category) — user-added entries
          customOnlyRows.push([row[0], row[1], row[2], row[3]]);
        }
      });
    } catch (e) {
      // If reading fails, use defaults
    }
  }

  // Build final mappings — Plaid-backed rows first
  const mappings = [];
  uniqueCategories.forEach(category => {
    if (existingMappings[category]) {
      mappings.push([category, existingMappings[category].customCategory, existingMappings[category].type, existingMappings[category].statement]);
    } else if (defaultMappings[category]) {
      mappings.push([category, ...defaultMappings[category]]);
    } else {
      mappings.push([category, "Other Expenses", "Expense", "P&L"]);
    }
  });

  // Sort by Type, then Custom Category
  mappings.sort((a, b) => {
    if (a[2] !== b[2]) return a[2].localeCompare(b[2]);
    return a[1].localeCompare(b[1]);
  });

  // Append custom-only rows (no Plaid category) after Plaid-backed rows
  const allMappings = [...mappings, ...customOnlyRows];

  // Clear old data rows only (preserve header rows)
  if (sheet.getLastRow() >= 5) {
    sheet.getRange(5, 1, sheet.getLastRow() - 4, 4).clearContent().clearFormat();
  }

  sheet.getRange(5, 1, allMappings.length, 4).setValues(allMappings);
  // Plaid Category (col A) read-only feel — no background
  // Custom Category, Type, Statement — yellow to signal editable
  sheet.getRange(5, 2, allMappings.length, 3).setBackground("#fffbea");

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setFrozenRows(4);

  // Ensure extra rows below data so the sheet doesn't feel cut off
  const cmEmptyRows = sheet.getMaxRows() - sheet.getLastRow();
  if (cmEmptyRows < 1000) sheet.insertRowsAfter(sheet.getMaxRows(), 1000 - cmEmptyRows);
}

/**
 * Setup General Ledger v3
 * Columns: Date | Transaction ID | Vendor | Plaid Category | Custom Category | Type | Statement | Account Name | Account Type | Debit | Credit | Memo
 * Custom Category, Type, Statement are formula-driven from Category Mapping but user-overridable.
 * Manual overrides persist across re-runs: if a cell value differs from what the formula would produce, it is preserved.
 */
function setupGeneralLedgerV3(sheet, transactionsSheet, headerMap, mappingSheet, ss) {
  let currentRow = 1;

  // Get column indices from transactions sheet
  const dateCol = getColumnIndex(headerMap, 'date');
  const amountCol = getColumnIndex(headerMap, 'amount');
  const pendingCol = getColumnIndex(headerMap, 'pending');
  const transactionIdCol = getColumnIndex(headerMap, 'transaction_id');
  const categoryPrimaryCol = getColumnIndex(headerMap, 'category_primary');
  const merchantCol = getColumnIndex(headerMap, 'merchant_name');
  const accountNameCol = getColumnIndex(headerMap, 'account_name');
  const descriptionRawCol = getColumnIndex(headerMap, 'description_raw');

  const dateColLetter = columnIndexToLetter(dateCol);
  const amountColLetter = columnIndexToLetter(amountCol);
  const pendingColLetter = columnIndexToLetter(pendingCol);
  const txnIdColLetter = columnIndexToLetter(transactionIdCol);
  const categoryPrimaryColLetter = columnIndexToLetter(categoryPrimaryCol);
  const merchantColLetter = columnIndexToLetter(merchantCol);
  const accountNameColLetter = columnIndexToLetter(accountNameCol);
  const descriptionRawColLetter = columnIndexToLetter(descriptionRawCol);

  // --- Build Category Mapping lookup in memory for override detection ---
  // mappingSheet columns: A=Plaid Category, B=Custom Category, C=Type, D=Statement
  const mappingData = mappingSheet.getLastRow() >= 5
    ? mappingSheet.getRange(5, 1, mappingSheet.getLastRow() - 4, 4).getValues()
    : [];
  const mappingLookup = {}; // plaidCategory -> {customCategory, type, statement}
  mappingData.forEach(row => {
    const plaid = row[0];
    if (plaid && plaid !== "") {
      mappingLookup[plaid] = { customCategory: row[1], type: row[2], statement: row[3] };
    }
  });
  // Also build reverse lookup: customCategory -> {type, statement} for XLOOKUP fallback resolution
  const customCategoryLookup = {};
  mappingData.forEach(row => {
    const customCat = row[1];
    if (customCat && customCat !== "" && !customCategoryLookup[customCat]) {
      customCategoryLookup[customCat] = { type: row[2], statement: row[3] };
    }
  });

  // --- Read existing GL to detect manual overrides ---
  // GL columns (v3): A=Date, B=TxnId, C=Vendor, D=PlaidCat, E=CustomCat, F=Type, G=Statement, H=AccountName, I=AccountType, J=Debit, K=Credit, L=Memo
  // Find existing ledger header row to locate data
  let existingLedgerHeaderRow = -1;
  const existingLastRow = sheet.getLastRow();
  if (existingLastRow > 0) {
    const existingData = sheet.getRange(1, 1, existingLastRow, 2).getValues();
    for (let i = 0; i < existingData.length; i++) {
      if (existingData[i][0] === 'Date' && existingData[i][1] === 'Transaction ID') {
        existingLedgerHeaderRow = i + 1; // 1-indexed
        break;
      }
    }
  }

  // Map of transaction_id -> {customCategory, type, statement} for manual overrides
  const manualOverrides = {};
  if (existingLedgerHeaderRow > 0 && existingLastRow > existingLedgerHeaderRow) {
    const numDataRows = existingLastRow - existingLedgerHeaderRow;
    const existingGLData = sheet.getRange(existingLedgerHeaderRow + 1, 1, numDataRows, 12).getValues();
    existingGLData.forEach(row => {
      const txnId    = row[1];  // Col B
      const plaidCat = row[3];  // Col D
      const customCat = row[4]; // Col E
      const type     = row[5];  // Col F
      const statement = row[6]; // Col G

      if (!txnId) return;

      // Resolve what the formula would produce for this Plaid Category
      const expected = mappingLookup[plaidCat] || { customCategory: "Uncategorized", type: "Expense", statement: "P&L" };

      const customCatOverridden = customCat !== "" && customCat !== expected.customCategory;
      // For Type and Statement, resolve from custom category (may differ from plaid mapping if custom cat was overridden)
      const resolvedType = customCategoryLookup[customCat] ? customCategoryLookup[customCat].type : expected.type;
      const resolvedStatement = customCategoryLookup[customCat] ? customCategoryLookup[customCat].statement : expected.statement;

      const typeOverridden = type !== "" && type !== resolvedType;
      const statementOverridden = statement !== "" && statement !== resolvedStatement;

      if (customCatOverridden || typeOverridden || statementOverridden) {
        manualOverrides[txnId] = {
          customCategory: customCatOverridden ? customCat : null,
          type: typeOverridden ? type : null,
          statement: statementOverridden ? statement : null
        };
      }
    });
  }

  // --- Clear and rebuild sheet ---
  // Row 1: Title
  sheet.getRange("A1")
    .setValue("General Ledger")
    .setFontSize(18)
    .setFontWeight("bold")
    .setFontColor("#023820")
    .setHorizontalAlignment("left");
  currentRow++;

  // Row 2: Description
  sheet.getRange("A2")
    .setValue("Complete transaction history with debit/credit accounting format.")
    .setFontSize(11)
    .setWrap(false)
    .setHorizontalAlignment("left");
  currentRow++;

  // Row 3: Blank
  currentRow++;

  // Row 4: Account Balance Configuration
  sheet.getRange(currentRow, 1).setValue("Account Balance Configuration")
    .setFontSize(12)
    .setFontWeight("bold")
    .setFontColor("#023820");
  currentRow++;

  sheet.getRange(currentRow, 1).setValue("Auto-detected accounts from Transactions. Edit Account Type and Starting Balance:");
  sheet.getRange(currentRow, 1).setFontStyle("italic").setFontSize(10);
  currentRow++;

  // Account config table headers
  const accountConfigHeaders = ["Account Name (from Transactions)", "Account Type", "Starting Balance", "As of Date"];
  sheet.getRange(currentRow, 1, 1, 4).setValues([accountConfigHeaders]);
  sheet.getRange(currentRow, 1, 1, 4).setFontWeight("bold").setBackground("#f3f3f3");
  currentRow++;

  const accountConfigStartRow = currentRow;
  const accountNameColIdx = getColumnIndex(headerMap, 'account_name');

  // Read existing account config to preserve user edits (Starting Balance, Account Type, As of Date)
  const existingAccountConfig = {};
  if (existingLedgerHeaderRow > 0) {
    const glAllData = sheet.getRange(1, 1, existingLedgerHeaderRow, 4).getValues();
    // Find config section by scanning for rows between "Account Name" header and blank row
    let inConfig = false;
    for (let i = 0; i < glAllData.length; i++) {
      if (glAllData[i][0] === "Account Name (from Transactions)") { inConfig = true; continue; }
      if (inConfig && (!glAllData[i][0] || glAllData[i][0] === "Date")) break;
      if (inConfig && glAllData[i][0]) {
        existingAccountConfig[glAllData[i][0]] = {
          accountType: glAllData[i][1],
          startingBalance: glAllData[i][2],
          asOfDate: glAllData[i][3]
        };
      }
    }
  }

  const lastTxnRowForAccounts = transactionsSheet.getLastRow();
  if (lastTxnRowForAccounts > 1) {
    const accountNames = transactionsSheet.getRange(2, accountNameColIdx, lastTxnRowForAccounts - 1, 1).getValues();
    const uniqueAccounts = [...new Set(accountNames.map(r => r[0]).filter(n => n && n.trim() !== ''))].sort();

    uniqueAccounts.forEach(accountName => {
      let accountType = "Asset";
      const nameLower = accountName.toLowerCase();
      if (nameLower.includes("credit card") || nameLower.includes("credit")) accountType = "Liability";
      else if (nameLower.includes("loan") || nameLower.includes("mortgage")) accountType = "Liability";

      const existing = existingAccountConfig[accountName];
      const now = new Date();
      const todayDate = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;

      sheet.getRange(currentRow, 1).setValue(accountName).setBackground("#fffbea");
      sheet.getRange(currentRow, 2).setValue(existing ? existing.accountType : accountType).setBackground("#fffbea");
      sheet.getRange(currentRow, 3).setValue(existing ? existing.startingBalance : 0).setNumberFormat("$#,##0.00").setBackground("#fffbea");
      sheet.getRange(currentRow, 4).setValue(existing ? existing.asOfDate : todayDate).setNumberFormat("yyyy-mm-dd").setBackground("#fffbea");
      currentRow++;
    });
  }

  const accountConfigEndRow = currentRow - 1;
  createNamedRange(sheet, "GL_AccountConfig", `A${accountConfigStartRow}:D${accountConfigEndRow}`);
  if (accountConfigEndRow >= accountConfigStartRow) {
    createNamedRange(sheet, "GL_CashBalance", `C${accountConfigStartRow}`);
  }

  currentRow++; // Blank row

  // Ledger headers — new column order (12 cols)
  // A=Date, B=TxnID, C=Vendor, D=Plaid Category, E=Custom Category, F=Type, G=Statement, H=Account Name, I=Account Type, J=Debit, K=Credit, L=Memo
  const headers = ["Date", "Transaction ID", "Vendor", "Plaid Category", "Custom Category", "Type", "Statement", "Account Name", "Account Type", "Debit", "Credit", "Memo"];
  sheet.getRange(currentRow, 1, 1, 12).setValues([headers]);
  sheet.getRange(currentRow, 1, 1, 12)
    .setFontWeight("bold")
    .setBackground("#0b703a")
    .setFontColor("white")
    .setHorizontalAlignment("left");

  const ledgerHeaderRow = currentRow;
  currentRow++;

  const lastTxnRow = transactionsSheet.getLastRow();

  if (lastTxnRow > 1) {
    const dataStartRow = currentRow;
    const numTxns = lastTxnRow - 1;

    // Read all transaction IDs to resolve overrides
    const txnIdValues = transactionsSheet.getRange(2, transactionIdCol, numTxns, 1).getValues();

    // Hoist mapping range — getLastRow() is an API call, do it once not per-row
    const mappingLastRow = mappingSheet.getLastRow();
    const mappingRange = `'Category Mapping'!$A$5:$D$${Math.max(mappingLastRow, 5)}`;
    const mappingRangeA = `'Category Mapping'!$A$5:$A$${Math.max(mappingLastRow, 5)}`;
    const mappingRangeB = `'Category Mapping'!$B$5:$B$${Math.max(mappingLastRow, 5)}`;
    const mappingRangeC = `'Category Mapping'!$C$5:$C$${Math.max(mappingLastRow, 5)}`;
    const mappingRangeD = `'Category Mapping'!$D$5:$D$${Math.max(mappingLastRow, 5)}`;

    // Build all formulas in memory, write in one batch
    const formulas = [];
    const overrideCells = []; // Track cells that need manual override values applied after formula write

    for (let i = 0; i < numTxns; i++) {
      const txnRow = i + 2;
      const ledgerRow = dataStartRow + i;
      const txnId = txnIdValues[i][0];
      const override = manualOverrides[txnId] || {};

      formulas.push([
        // A: Date
        `=DATEVALUE(Transactions!${dateColLetter}${txnRow})`,
        // B: Transaction ID
        `=Transactions!${txnIdColLetter}${txnRow}`,
        // C: Vendor
        `=Transactions!${merchantColLetter}${txnRow}`,
        // D: Plaid Category (direct reference)
        `=Transactions!${categoryPrimaryColLetter}${txnRow}`,
        // E: Custom Category — XLOOKUP Plaid Category against mapping col A, return col B
        `=IFERROR(XLOOKUP(D${ledgerRow}, ${mappingRangeA}, ${mappingRangeB}), "Uncategorized")`,
        // F: Type — XLOOKUP Custom Category against mapping col B, return col C
        `=IFERROR(XLOOKUP(E${ledgerRow}, ${mappingRangeB}, ${mappingRangeC}), "Expense")`,
        // G: Statement — XLOOKUP Custom Category against mapping col B, return col D
        `=IFERROR(XLOOKUP(E${ledgerRow}, ${mappingRangeB}, ${mappingRangeD}), "P&L")`,
        // H: Account Name
        `=Transactions!${accountNameColLetter}${txnRow}`,
        // I: Account Type — XLOOKUP Account Name against config table
        `=IFERROR(XLOOKUP(H${ledgerRow}, 'General Ledger'!$A$${accountConfigStartRow}:$A$${accountConfigEndRow}, 'General Ledger'!$B$${accountConfigStartRow}:$B$${accountConfigEndRow}), "")`,
        // J: Debit
        `=IF(AND(Transactions!${amountColLetter}${txnRow}<0, I${ledgerRow}="Asset"), ABS(Transactions!${amountColLetter}${txnRow}), IF(AND(Transactions!${amountColLetter}${txnRow}<0, I${ledgerRow}="Liability"), ABS(Transactions!${amountColLetter}${txnRow}), ""))`,
        // K: Credit
        `=IF(AND(Transactions!${amountColLetter}${txnRow}>0, I${ledgerRow}="Liability"), ABS(Transactions!${amountColLetter}${txnRow}), IF(AND(Transactions!${amountColLetter}${txnRow}>0, I${ledgerRow}="Asset"), ABS(Transactions!${amountColLetter}${txnRow}), ""))`,
        // L: Memo
        `=Transactions!${descriptionRawColLetter}${txnRow}`
      ]);

      // Track overrides to apply after formula write
      if (override.customCategory || override.type || override.statement) {
        overrideCells.push({ ledgerRow, override });
      }
    }

    // Single batch write for all formulas
    sheet.getRange(dataStartRow, 1, formulas.length, 12).setFormulas(formulas);

    // Apply yellow background to Custom Category (E), Type (F), Statement (G) columns
    sheet.getRange(dataStartRow, 5, numTxns, 3).setBackground("#fffbea");

    // Format date and currency columns
    sheet.getRange(dataStartRow, 1, numTxns, 1).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(dataStartRow, 10, numTxns, 2).setNumberFormat("$#,##0.00"); // Debit (J) and Credit (K)

    // Re-apply manual overrides as values — batch by column to minimize API calls
    if (overrideCells.length > 0) {
      overrideCells.forEach(({ ledgerRow, override }) => {
        if (override.customCategory) sheet.getRange(ledgerRow, 5).setValue(override.customCategory);
        if (override.type) sheet.getRange(ledgerRow, 6).setValue(override.type);
        if (override.statement) sheet.getRange(ledgerRow, 7).setValue(override.statement);
      });
    }
  }

  sheet.setFrozenRows(ledgerHeaderRow);
  sheet.setColumnWidth(1, 100);  // Date
  sheet.setColumnWidth(2, 160);  // Transaction ID
  sheet.setColumnWidth(3, 200);  // Vendor
  sheet.setColumnWidth(4, 180);  // Plaid Category
  sheet.setColumnWidth(5, 180);  // Custom Category
  sheet.setColumnWidth(6, 120);  // Type
  sheet.setColumnWidth(7, 120);  // Statement
  sheet.setColumnWidth(8, 180);  // Account Name
  sheet.setColumnWidth(9, 120);  // Account Type
  sheet.setColumnWidth(10, 100); // Debit
  sheet.setColumnWidth(11, 100); // Credit
  sheet.setColumnWidth(12, 400); // Memo

  // Ensure extra rows below data so the sheet doesn't feel cut off
  const glLastDataRow = sheet.getLastRow();
  const glEmptyRows = sheet.getMaxRows() - glLastDataRow;
  if (glEmptyRows < 1000) sheet.insertRowsAfter(sheet.getMaxRows(), 1000 - glEmptyRows);
  sheet.getRange(glLastDataRow + 1, 1, 1000, 12).clearFormat();
}

/**
 * Setup Financial Statements v3 - Consolidated with Monthly Trending
 * GL columns: A=Date, B=TxnID, C=Vendor, D=Plaid Category, E=Custom Category, F=Type, G=Statement, H=Account Name, I=Account Type, J=Debit, K=Credit, L=Memo
 * SUMIFS reference: Category=col E ($E:$E), Type=col F ($F:$F), Statement=col G ($G:$G), Debit=col J ($J:$J), Credit=col K ($K:$K), Date=col A ($A:$A), Account=col H ($H:$H)
 */
function setupFinancialStatementsV3(sheet, ledgerSheet, transactionsSheet, headerMap, ss) {
  sheet.clear();

  let currentRow = 1;

  // Get column letters
  const dateCol = getColumnIndex(headerMap, 'date');
  const amountCol = getColumnIndex(headerMap, 'amount');
  const categoryCol = getColumnIndex(headerMap, 'category_primary');
  const pendingCol = getColumnIndex(headerMap, 'pending');
  const accountNameCol = getColumnIndex(headerMap, 'account_name');

  const dateColLetter = columnIndexToLetter(dateCol);
  const amountColLetter = columnIndexToLetter(amountCol);
  const categoryColLetter = columnIndexToLetter(categoryCol);
  const pendingColLetter = columnIndexToLetter(pendingCol);
  const accountNameColLetter = columnIndexToLetter(accountNameCol);

  // Get all transactions to determine date range
  const transactions = getTransactionData(transactionsSheet);
  const dates = transactions.map(t => parseDate(t.date)).filter(d => d);
  const minDate = new Date(Math.min(...dates));
  const maxDate = new Date(Math.max(...dates));

  // Generate month list (last 12 months)
  const months = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date();
    d.setMonth(d.getMonth() - i);
    months.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  }

  // Row 1: Title (matching cashflow/ledger styling)
  sheet.getRange("A1")
    .setValue("Financial Statements")
    .setFontSize(18)
    .setFontWeight("bold")
    .setFontColor("#023820")
    .setHorizontalAlignment("left");
  currentRow++;

  // Row 2: Description
  sheet.getRange("A2")
    .setValue("Consolidated financial statements with monthly trending showing Profit & Loss, Balance Sheet, and Cash Flow.")
    .setFontSize(11)
    .setWrap(false)
    .setHorizontalAlignment("left");
  currentRow++;

  // Row 3: Blank
  currentRow++;

  // PROFIT & LOSS SECTION
  const plSectionRow = currentRow;
  sheet.getRange(currentRow, 1).setValue("PROFIT & LOSS")
    .setFontWeight("bold")
    .setFontSize(12);
  // Apply background color across entire table width (2 fixed columns + 12 month columns = 14)
  sheet.getRange(currentRow, 1, 1, 2 + months.length)
    .setBackground("#0b703a")
    .setFontColor("white");
  currentRow++;

  // P&L Headers - set dates as 1st of each month
  const plHeaders = ["Account", "Type"];
  const monthDates = [];

  // Add text headers first
  sheet.getRange(currentRow, 1, 1, 2).setValues([["Account", "Type"]]);
  sheet.getRange(currentRow, 1, 1, 2).setFontWeight("bold").setBackground("#f3f3f3");

  // Add date headers as a batch
  months.forEach((month, index) => {
    const parts = month.split('-');
    const dateString = `${parts[0]}-${parts[1]}-01`;
    monthDates.push(dateString);
  });
  sheet.getRange(currentRow, 3, 1, months.length).setValues([monthDates]);
  sheet.getRange(currentRow, 3, 1, months.length).setNumberFormat("mmm-yy").setFontWeight("bold").setBackground("#f3f3f3");

  const plHeaderRow = currentRow;
  currentRow++;

  // REVENUE section
  sheet.getRange(currentRow, 1).setValue("REVENUE").setFontWeight("bold");
  currentRow++;

  // Revenue rows (pull from General Ledger — deduped Custom Category, Type, Statement)
  const plStartRow = currentRow;

  // Read GL data once: col E=Custom Category, F=Type, G=Statement
  // GL header row is dynamic — find it by locating the "Date" header row
  const glAllData = ledgerSheet.getDataRange().getValues();
  let glDataStartRow = -1;
  for (let i = 0; i < glAllData.length; i++) {
    if (glAllData[i][0] === 'Date' && glAllData[i][1] === 'Transaction ID') {
      glDataStartRow = i + 1; // 0-indexed row after header
      break;
    }
  }

  // Build deduped category map from GL rows: customCategory -> { type, statement }
  const glCategoryMap = {}; // customCategory -> { type, statement }
  if (glDataStartRow >= 0) {
    for (let i = glDataStartRow; i < glAllData.length; i++) {
      const customCat = glAllData[i][4]; // Col E
      const type      = glAllData[i][5]; // Col F
      const statement = glAllData[i][6]; // Col G
      if (customCat && customCat !== "" && !glCategoryMap[customCat]) {
        glCategoryMap[customCat] = { type: type || "Expense", statement: statement || "P&L" };
      }
    }
  }

  const revenueAccountsMap = {}; // customCategory -> type
  for (const [customCat, meta] of Object.entries(glCategoryMap)) {
    if (meta.type === "Revenue" && meta.statement === "P&L") {
      revenueAccountsMap[customCat] = meta.type;
    }
  }

  const revenueNames = Object.keys(revenueAccountsMap);
  if (revenueNames.length > 0) {
    const revenueValues = [];
    const revenueFormulas = [];
    revenueNames.forEach((accountName, rowOffset) => {
      const type = revenueAccountsMap[accountName];
      const ledgerRow = currentRow + rowOffset;
      const accountRef = `$A${ledgerRow}`;
      const typeRef = `$B${ledgerRow}`;
      revenueValues.push([accountName, type]);
      revenueFormulas.push(months.map((m, idx) => {
        const headerCol = columnIndexToLetter(3 + idx);
        return `=SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$E:$E, ${accountRef}, 'General Ledger'!$F:$F, ${typeRef}, ` +
          `'General Ledger'!$A:$A, ">="&${headerCol}$${plHeaderRow}, ` +
          `'General Ledger'!$A:$A, "<"&EOMONTH(${headerCol}$${plHeaderRow}, 0)+1) - ` +
          `SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$E:$E, ${accountRef}, 'General Ledger'!$F:$F, ${typeRef}, ` +
          `'General Ledger'!$A:$A, ">="&${headerCol}$${plHeaderRow}, ` +
          `'General Ledger'!$A:$A, "<"&EOMONTH(${headerCol}$${plHeaderRow}, 0)+1)`;
      }));
    });
    sheet.getRange(currentRow, 1, revenueNames.length, 2).setValues(revenueValues);
    sheet.getRange(currentRow, 3, revenueNames.length, months.length).setFormulas(revenueFormulas);
    currentRow += revenueNames.length;
  }

  // Total Revenue
  sheet.getRange(currentRow, 1).setValue("Total Revenue").setFontWeight("bold");
  const hasRevenueRows = currentRow > plStartRow;
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    if (hasRevenueRows) {
      sheet.getRange(currentRow, 3 + idx).setFormula(`=SUM(${col}${plStartRow}:${col}${currentRow - 1})`);
    } else {
      sheet.getRange(currentRow, 3 + idx).setValue(0);
    }
  });
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#d9ead3").setFontWeight("bold");
  const totalRevenueRow = currentRow;
  currentRow++;

  // Blank row
  currentRow++;

  // Expenses (pull from COA dynamically)
  sheet.getRange(currentRow, 1).setValue("EXPENSES").setFontWeight("bold");
  currentRow++;

  const expenseStartRow = currentRow;

  const expenseAccountsMap = {}; // customCategory -> type
  for (const [customCat, meta] of Object.entries(glCategoryMap)) {
    if (meta.type === "Expense" && meta.statement === "P&L") {
      expenseAccountsMap[customCat] = meta.type;
    }
  }

  const expenseNames = Object.keys(expenseAccountsMap);
  if (expenseNames.length > 0) {
    const expenseValues = [];
    const expenseFormulas = [];
    expenseNames.forEach((accountName, rowOffset) => {
      const type = expenseAccountsMap[accountName];
      const ledgerRow = currentRow + rowOffset;
      const accountRef = `$A${ledgerRow}`;
      const typeRef = `$B${ledgerRow}`;
      expenseValues.push([accountName, type]);
      expenseFormulas.push(months.map((m, idx) => {
        const headerCol = columnIndexToLetter(3 + idx);
        return `=SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$E:$E, ${accountRef}, 'General Ledger'!$F:$F, ${typeRef}, ` +
          `'General Ledger'!$A:$A, ">="&${headerCol}$${plHeaderRow}, ` +
          `'General Ledger'!$A:$A, "<"&EOMONTH(${headerCol}$${plHeaderRow}, 0)+1) - ` +
          `SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$E:$E, ${accountRef}, 'General Ledger'!$F:$F, ${typeRef}, ` +
          `'General Ledger'!$A:$A, ">="&${headerCol}$${plHeaderRow}, ` +
          `'General Ledger'!$A:$A, "<"&EOMONTH(${headerCol}$${plHeaderRow}, 0)+1)`;
      }));
    });
    sheet.getRange(currentRow, 1, expenseNames.length, 2).setValues(expenseValues);
    sheet.getRange(currentRow, 3, expenseNames.length, months.length).setFormulas(expenseFormulas);
    currentRow += expenseNames.length;
  }

  // Total Expenses
  sheet.getRange(currentRow, 1).setValue("Total Expenses").setFontWeight("bold");
  const hasExpenseRows = currentRow > expenseStartRow;
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    if (hasExpenseRows) {
      sheet.getRange(currentRow, 3 + idx).setFormula(`=SUM(${col}${expenseStartRow}:${col}${currentRow - 1})`);
    } else {
      sheet.getRange(currentRow, 3 + idx).setValue(0);
    }
  });
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#f4cccc").setFontWeight("bold");
  const totalExpensesRow = currentRow;
  currentRow++;

  // Net Income
  sheet.getRange(currentRow, 1).setValue("NET INCOME").setFontWeight("bold").setFontSize(11);
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(`=${col}${totalRevenueRow}+${col}${totalExpensesRow}`);
  });
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#cfe2f3").setFontWeight("bold");
  currentRow += 2;

  // Format currency (now starting at column 3)
  sheet.getRange(plStartRow, 3, currentRow - plStartRow, months.length).setNumberFormat("$#,##0.00_);($#,##0.00);\"-  \"");

  const netIncomeRow = currentRow - 2; // Save for later reference

  // Get config table range from General Ledger for Balance Sheet formulas
  const namedRanges = ledgerSheet.getNamedRanges();
  let accountConfigStartRow, accountConfigEndRow;

  const configRange = namedRanges.find(nr => nr.getName() === 'GL_AccountConfig');
  if (configRange) {
    const range = configRange.getRange();
    accountConfigStartRow = range.getRow();
    accountConfigEndRow = range.getLastRow();
  } else {
    // Fallback: assume config starts at row 7 (standard structure)
    accountConfigStartRow = 7;
    // Count rows in ledgerSheet to find end
    const glData = ledgerSheet.getDataRange().getValues();
    // Find the first row that contains "Date" header (marks end of config table)
    for (let i = 6; i < glData.length; i++) {
      if (glData[i][0] === 'Date') {
        accountConfigEndRow = i; // Row before the ledger headers
        break;
      }
    }
    if (!accountConfigEndRow) accountConfigEndRow = accountConfigStartRow + 10; // Safe default
  }

  // BALANCE SHEET SECTION
  currentRow += 2; // Extra spacing
  sheet.getRange(currentRow, 1).setValue("BALANCE SHEET")
    .setFontWeight("bold")
    .setFontSize(12);
  // Apply background color across entire table width (2 fixed columns + 12 month columns = 14)
  sheet.getRange(currentRow, 1, 1, 2 + months.length)
    .setBackground("#0b703a")
    .setFontColor("white");
  currentRow++;

  // Balance Sheet Headers - set dates as 1st of each month
  const bsHeaders = ["Account", "Type"];

  // Add text headers first
  sheet.getRange(currentRow, 1, 1, 2).setValues([["Account", "Type"]]);
  sheet.getRange(currentRow, 1, 1, 2).setFontWeight("bold").setBackground("#f3f3f3");

  // Add date headers as a batch
  sheet.getRange(currentRow, 3, 1, months.length).setValues([monthDates]);
  sheet.getRange(currentRow, 3, 1, months.length).setNumberFormat("mmm-yy").setFontWeight("bold").setBackground("#f3f3f3");

  const bsHeaderRow = currentRow;
  currentRow++;

  // ASSETS
  sheet.getRange(currentRow, 1).setValue("ASSETS").setFontWeight("bold");
  currentRow++;

  const assetsStartRow = currentRow;

  // Dynamically add all Asset accounts from GL_AccountConfig
  // Track Balance Sheet row numbers for cash accounts
  const configData = ledgerSheet.getRange("GL_AccountConfig").getValues();
  const cashAccountBalanceSheetRows = {};

  const assetRows = configData.filter(row => row[0] && row[1] === "Asset");
  const assetValues = [];
  const assetFormulas = [];
  assetRows.forEach((row, rowOffset) => {
    const accountName = row[0];
    const ledgerRow = currentRow + rowOffset;
    cashAccountBalanceSheetRows[accountName] = ledgerRow;
    const accountRef = `$A${ledgerRow}`;
    assetValues.push([accountName, "Asset"]);
    assetFormulas.push(months.map((m, monthIdx) => {
      const headerCol = columnIndexToLetter(3 + monthIdx);
      const monthEndDate = `EOMONTH(${headerCol}$${bsHeaderRow}, 0)`;
      return `=IF(INT(${monthEndDate})=INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0)), ` +
        `IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 3, FALSE), 0), ` +
        `IF(${monthEndDate}<IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0), ` +
        `IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 3, FALSE), 0) - ` +
        `SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$H:$H, ${accountRef}, ` +
        `'General Ledger'!$A:$A, ">"&INT(${monthEndDate}), ` +
        `'General Ledger'!$A:$A, "<="&INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0))) + ` +
        `SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$H:$H, ${accountRef}, ` +
        `'General Ledger'!$A:$A, ">"&INT(${monthEndDate}), ` +
        `'General Ledger'!$A:$A, "<="&INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0))), ` +
        `IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 3, FALSE), 0) + ` +
        `SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$H:$H, ${accountRef}, ` +
        `'General Ledger'!$A:$A, ">"&INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0)), ` +
        `'General Ledger'!$A:$A, "<="&INT(${monthEndDate})) - ` +
        `SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$H:$H, ${accountRef}, ` +
        `'General Ledger'!$A:$A, ">"&INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0)), ` +
        `'General Ledger'!$A:$A, "<="&INT(${monthEndDate}))))`;
    }));
  });
  if (assetRows.length > 0) {
    sheet.getRange(currentRow, 1, assetRows.length, 2).setValues(assetValues);
    sheet.getRange(currentRow, 3, assetRows.length, months.length).setFormulas(assetFormulas);
    currentRow += assetRows.length;
  }

  // Total Assets
  sheet.getRange(currentRow, 1).setValue("Total Assets").setFontWeight("bold");
  const totalAssetsFormulas = months.map((m, idx) => [`=SUM(${columnIndexToLetter(3 + idx)}${assetsStartRow}:${columnIndexToLetter(3 + idx)}${currentRow - 1})`]);
  sheet.getRange(currentRow, 3, 1, months.length).setFormulas([totalAssetsFormulas.map(f => f[0])]);
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#d9d2e9").setFontWeight("bold");
  const totalAssetsRow = currentRow;
  currentRow++;

  currentRow++; // Blank row

  // LIABILITIES
  sheet.getRange(currentRow, 1).setValue("LIABILITIES").setFontWeight("bold");
  currentRow++;

  const liabilitiesStartRow = currentRow;

  // Dynamically add all Liability accounts from GL_AccountConfig
  // Track Balance Sheet row numbers for Cash Flow adjustments
  const liabilityBalanceSheetRows = {};

  const liabilityRows = configData.filter(row => row[0] && row[1] === "Liability");
  const liabilityValues = [];
  const liabilityFormulas = [];
  liabilityRows.forEach((row, rowOffset) => {
    const accountName = row[0];
    const ledgerRow = currentRow + rowOffset;
    liabilityBalanceSheetRows[accountName] = ledgerRow;
    const accountRef = `$A${ledgerRow}`;
    liabilityValues.push([accountName, "Liability"]);
    liabilityFormulas.push(months.map((m, monthIdx) => {
      const headerCol = columnIndexToLetter(3 + monthIdx);
      const monthEndDate = `EOMONTH(${headerCol}$${bsHeaderRow}, 0)`;
      return `=IF(INT(${monthEndDate})=INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0)), ` +
        `IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 3, FALSE), 0), ` +
        `IF(${monthEndDate}<IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0), ` +
        `IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 3, FALSE), 0) - ` +
        `SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$H:$H, ${accountRef}, ` +
        `'General Ledger'!$A:$A, ">"&INT(${monthEndDate}), ` +
        `'General Ledger'!$A:$A, "<="&INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0))) + ` +
        `SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$H:$H, ${accountRef}, ` +
        `'General Ledger'!$A:$A, ">"&INT(${monthEndDate}), ` +
        `'General Ledger'!$A:$A, "<="&INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0))), ` +
        `IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 3, FALSE), 0) + ` +
        `SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$H:$H, ${accountRef}, ` +
        `'General Ledger'!$A:$A, ">"&INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0)), ` +
        `'General Ledger'!$A:$A, "<="&INT(${monthEndDate})) - ` +
        `SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$H:$H, ${accountRef}, ` +
        `'General Ledger'!$A:$A, ">"&INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0)), ` +
        `'General Ledger'!$A:$A, "<="&INT(${monthEndDate}))))`;
    }));
  });
  if (liabilityRows.length > 0) {
    sheet.getRange(currentRow, 1, liabilityRows.length, 2).setValues(liabilityValues);
    sheet.getRange(currentRow, 3, liabilityRows.length, months.length).setFormulas(liabilityFormulas);
    currentRow += liabilityRows.length;
  }

  // Total Liabilities
  sheet.getRange(currentRow, 1).setValue("Total Liabilities").setFontWeight("bold");
  sheet.getRange(currentRow, 3, 1, months.length).setFormulas([months.map((m, idx) => `=SUM(${columnIndexToLetter(3+idx)}${liabilitiesStartRow}:${columnIndexToLetter(3+idx)}${currentRow - 1})`)]);
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#f4cccc").setFontWeight("bold");
  const totalLiabilitiesRow = currentRow;
  currentRow++;

  currentRow++; // Blank row

  // EQUITY (calculated as Assets - Liabilities)
  sheet.getRange(currentRow, 1).setValue("EQUITY").setFontWeight("bold");
  currentRow++;

  const equityStartRow = currentRow;

  // Retained Earnings (will reference Total Equity row calculated below)
  sheet.getRange(currentRow, 1, 1, 2).setValues([["Retained Earnings", "Equity"]]);
  const retainedEarningsRow = currentRow;
  currentRow++;

  // Total Equity
  sheet.getRange(currentRow, 1).setValue("Total Equity").setFontWeight("bold");
  sheet.getRange(currentRow, 3, 1, months.length).setFormulas([months.map((m, idx) => `=${columnIndexToLetter(3+idx)}${totalAssetsRow}-${columnIndexToLetter(3+idx)}${totalLiabilitiesRow}`)]);
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#d9d2e9").setFontWeight("bold");
  const totalEquityRow = currentRow;
  currentRow++;

  // Retained Earnings = Total Equity (batch)
  sheet.getRange(retainedEarningsRow, 3, 1, months.length).setFormulas([months.map((m, idx) => `=${columnIndexToLetter(3+idx)}${totalEquityRow}`)]);

  // Check: Total Liabilities + Equity should equal Total Assets
  sheet.getRange(currentRow, 1).setValue("Check: Liabilities + Equity").setFontStyle("italic");
  sheet.getRange(currentRow, 3, 1, months.length).setFormulas([months.map((m, idx) => `=${columnIndexToLetter(3+idx)}${totalLiabilitiesRow}+${columnIndexToLetter(3+idx)}${totalEquityRow}`)]);
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setFontStyle("italic").setBackground("#fff2cc");
  currentRow += 2;

  // Format Balance Sheet currency
  sheet.getRange(assetsStartRow, 3, currentRow - assetsStartRow, months.length).setNumberFormat("$#,##0.00_);($#,##0.00);\"-  \"");

  // CASH FLOW STATEMENT SECTION
  currentRow += 2; // Extra spacing
  sheet.getRange(currentRow, 1).setValue("CASH FLOW STATEMENT")
    .setFontWeight("bold")
    .setFontSize(12);
  // Apply background color across entire table width (2 fixed columns + 12 month columns = 14)
  sheet.getRange(currentRow, 1, 1, 2 + months.length)
    .setBackground("#0b703a")
    .setFontColor("white");
  currentRow++;

  // Cash Flow Headers - set dates as 1st of each month
  const cfHeaders = ["Account", "Type"];

  // Add text headers first
  sheet.getRange(currentRow, 1, 1, 2).setValues([["Account", "Type"]]);
  sheet.getRange(currentRow, 1, 1, 2).setFontWeight("bold").setBackground("#f3f3f3");

  // Add date headers as a batch
  sheet.getRange(currentRow, 3, 1, months.length).setValues([monthDates]);
  sheet.getRange(currentRow, 3, 1, months.length).setNumberFormat("mmm-yy").setFontWeight("bold").setBackground("#f3f3f3");

  const cfHeaderRow = currentRow;
  currentRow++;

  // OPERATING ACTIVITIES (Indirect Method)
  sheet.getRange(currentRow, 1).setValue("OPERATING ACTIVITIES").setFontWeight("bold");
  currentRow++;

  const cfOperatingStartRow = currentRow;

  // Net Income (from P&L)
  sheet.getRange(currentRow, 1, 1, 2).setValues([["Net Income", "Operating"]]);
  sheet.getRange(currentRow, 3, 1, months.length).setFormulas([months.map((m, idx) => `=${columnIndexToLetter(3+idx)}${netIncomeRow}`)]);
  const netIncomeRowCF = currentRow;
  currentRow++;

  // Adjustments for changes in working capital
  sheet.getRange(currentRow, 1).setValue("Changes in Working Capital:").setFontStyle("italic");
  currentRow++;

  // Changes in Liabilities — batch write
  const liabilityAccountNames = Object.keys(liabilityBalanceSheetRows);
  if (liabilityAccountNames.length > 0) {
    const liabCFValues = [];
    const liabCFFormulas = [];
    liabilityAccountNames.forEach((accountName, rowOffset) => {
      const bsRowNum = liabilityBalanceSheetRows[accountName];
      const accountCell = `$A${bsRowNum}`;
      liabCFValues.push([`  Increase in ${accountName}`, "Operating"]);
      liabCFFormulas.push(months.map((m, monthIdx) => {
        const headerCol = columnIndexToLetter(3 + monthIdx);
        return `=SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$H:$H, ${accountCell}, ` +
          `'General Ledger'!$A:$A, ">="&${headerCol}$${cfHeaderRow}, ` +
          `'General Ledger'!$A:$A, "<="&INT(EOMONTH(${headerCol}$${cfHeaderRow}, 0))) - ` +
          `SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$H:$H, ${accountCell}, ` +
          `'General Ledger'!$A:$A, ">="&${headerCol}$${cfHeaderRow}, ` +
          `'General Ledger'!$A:$A, "<="&INT(EOMONTH(${headerCol}$${cfHeaderRow}, 0)))`;
      }));
    });
    sheet.getRange(currentRow, 1, liabilityAccountNames.length, 2).setValues(liabCFValues);
    sheet.getRange(currentRow, 3, liabilityAccountNames.length, months.length).setFormulas(liabCFFormulas);
    currentRow += liabilityAccountNames.length;
  }

  // Other Working Capital (plug to reconcile: Total Cash Change - Net Income - Liabilities - CapEx - Loan Proceeds)
  sheet.getRange(currentRow, 1).setValue("  Other Working Capital");
  sheet.getRange(currentRow, 2).setValue("Operating");
  const otherWorkingCapitalRow = currentRow;
  // Formula will be added after we know CapEx and Loan Proceeds rows
  currentRow++;

  // Cash from Operating Activities
  sheet.getRange(currentRow, 1).setValue("Cash from Operating Activities").setFontWeight("bold");
  sheet.getRange(currentRow, 3, 1, months.length).setFormulas([months.map((m, idx) => `=SUM(${columnIndexToLetter(3+idx)}${cfOperatingStartRow}:${columnIndexToLetter(3+idx)}${currentRow - 1})`)]);
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#d9ead3").setFontWeight("bold");
  const totalCFOperatingRow = currentRow;
  currentRow++;

  currentRow++; // Blank row

  // INVESTING ACTIVITIES
  sheet.getRange(currentRow, 1).setValue("INVESTING ACTIVITIES").setFontWeight("bold");
  currentRow++;

  const cfInvestingStartRow = currentRow;

  // Manual input for capital expenditures
  sheet.getRange(currentRow, 1, 1, 2).setValues([["Capital Expenditures", "Investing"]]);
  sheet.getRange(currentRow, 3, 1, months.length).setValues([months.map(() => 0)]).setBackground("#fffbea");
  const capExRow = currentRow;
  currentRow++;

  // Cash from Investing Activities
  sheet.getRange(currentRow, 1).setValue("Cash from Investing Activities").setFontWeight("bold");
  sheet.getRange(currentRow, 3, 1, months.length).setFormulas([months.map((m, idx) => `=SUM(${columnIndexToLetter(3+idx)}${cfInvestingStartRow}:${columnIndexToLetter(3+idx)}${currentRow - 1})`)]);
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#d9ead3").setFontWeight("bold");
  const totalCFInvestingRow = currentRow;
  currentRow++;

  currentRow++; // Blank row

  // FINANCING ACTIVITIES
  sheet.getRange(currentRow, 1).setValue("FINANCING ACTIVITIES").setFontWeight("bold");
  currentRow++;

  const cfFinancingStartRow = currentRow;

  // Manual input for loan proceeds/repayments
  sheet.getRange(currentRow, 1, 1, 2).setValues([["Loan Proceeds / Repayments", "Financing"]]);
  sheet.getRange(currentRow, 3, 1, months.length).setValues([months.map(() => 0)]).setBackground("#fffbea");
  const loanProceedsRow = currentRow;
  currentRow++;

  // Cash from Financing Activities
  sheet.getRange(currentRow, 1).setValue("Cash from Financing Activities").setFontWeight("bold");
  sheet.getRange(currentRow, 3, 1, months.length).setFormulas([months.map((m, idx) => `=SUM(${columnIndexToLetter(3+idx)}${cfFinancingStartRow}:${columnIndexToLetter(3+idx)}${currentRow - 1})`)]);
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#d9ead3").setFontWeight("bold");
  const totalCFFinancingRow = currentRow;
  currentRow++;

  currentRow++; // Blank row

  // NET CHANGE IN CASH
  sheet.getRange(currentRow, 1).setValue("NET CHANGE IN CASH").setFontWeight("bold").setFontSize(11);
  sheet.getRange(currentRow, 3, 1, months.length).setFormulas([months.map((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    return `=${col}${totalCFOperatingRow}+${col}${totalCFInvestingRow}+${col}${totalCFFinancingRow}`;
  })]);
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#fce5cd").setFontWeight("bold");
  const netChangeInCashRow = currentRow;
  currentRow++;

  currentRow++; // Blank row

  // RECONCILIATION - Change in Cash from Balance Sheet (batch)
  sheet.getRange(currentRow, 1).setValue("Change in Cash (from Balance Sheet):").setFontStyle("italic");
  const cashNamesForRecon = Object.keys(cashAccountBalanceSheetRows);
  const reconChangeFormulas = months.map((m, monthIdx) => {
    if (monthIdx === 0) return "";
    const col = columnIndexToLetter(3 + monthIdx);
    const prevCol = columnIndexToLetter(3 + monthIdx - 1);
    return cashNamesForRecon.length > 0
      ? "=" + cashNamesForRecon.map(n => `(${col}${cashAccountBalanceSheetRows[n]}-${prevCol}${cashAccountBalanceSheetRows[n]})`).join("+")
      : "0";
  });
  // setFormulas requires no empty strings — write values and formulas separately
  sheet.getRange(currentRow, 3, 1, 1).setValue("");
  if (reconChangeFormulas.slice(1).length > 0) {
    sheet.getRange(currentRow, 4, 1, months.length - 1).setFormulas([reconChangeFormulas.slice(1)]);
  }
  const changeFromBalanceSheetRow = currentRow;
  currentRow++;

  // RECONCILIATION - Difference (batch)
  sheet.getRange(currentRow, 1).setValue("Difference (to investigate):").setFontStyle("italic").setFontColor("#cc0000");
  const diffFormulas = months.map((m, idx) => {
    if (idx === 0) return "";
    const col = columnIndexToLetter(3 + idx);
    return `=ROUND(${col}${netChangeInCashRow}-${col}${changeFromBalanceSheetRow},2)`;
  });
  sheet.getRange(currentRow, 3, 1, 1).setValue("");
  if (diffFormulas.slice(1).length > 0) {
    sheet.getRange(currentRow, 4, 1, months.length - 1).setFormulas([diffFormulas.slice(1)]);
  }
  sheet.getRange(currentRow, 3, 1, months.length).setFontColor("#cc0000");
  currentRow++;

  // Now go back and fill in "Other Working Capital" formula (plug to reconcile) — batch
  const cashAccountNames = Object.keys(cashAccountBalanceSheetRows);
  const owcFormulas = months.map((m, monthIdx) => {
    const col = columnIndexToLetter(3 + monthIdx);
    const cashChangeFormula = cashAccountNames.length > 0
      ? cashAccountNames.map(accountName => {
          const bsRowNum = cashAccountBalanceSheetRows[accountName];
          return `(SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$H:$H, $A${bsRowNum}, ` +
                 `'General Ledger'!$A:$A, ">="&${col}$${cfHeaderRow}, ` +
                 `'General Ledger'!$A:$A, "<="&INT(EOMONTH(${col}$${cfHeaderRow}, 0))) - ` +
                 `SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$H:$H, $A${bsRowNum}, ` +
                 `'General Ledger'!$A:$A, ">="&${col}$${cfHeaderRow}, ` +
                 `'General Ledger'!$A:$A, "<="&INT(EOMONTH(${col}$${cfHeaderRow}, 0))))`;
        }).join("+")
      : "0";
    const liabNames = Object.keys(liabilityBalanceSheetRows);
    const liabilitySum = liabNames.length > 0
      ? `SUM(${col}${cfOperatingStartRow + 2}:${col}${otherWorkingCapitalRow - 1})`
      : "0";
    return `=${cashChangeFormula}-${col}${cfOperatingStartRow}-${liabilitySum}-${col}${capExRow}-${col}${loanProceedsRow}`;
  });
  sheet.getRange(otherWorkingCapitalRow, 3, 1, months.length).setFormulas([owcFormulas]);

  // Format Cash Flow currency
  sheet.getRange(cfOperatingStartRow, 3, currentRow - cfOperatingStartRow, months.length).setNumberFormat("$#,##0.00_);($#,##0.00);\"-  \"");

  // Set column widths
  sheet.setColumnWidth(1, 250); // Account
  sheet.setColumnWidth(2, 100); // Type
  for (let i = 0; i < months.length; i++) {
    sheet.setColumnWidth(3 + i, 100);
  }

  // Freeze after column B (columns A and B frozen)
  sheet.setFrozenRows(plHeaderRow);
  sheet.setFrozenColumns(2);

  // Ensure extra rows below data so the sheet doesn't feel cut off
  const fsLastDataRow = sheet.getLastRow();
  const fsEmptyRows = sheet.getMaxRows() - fsLastDataRow;
  if (fsEmptyRows < 1000) sheet.insertRowsAfter(sheet.getMaxRows(), 1000 - fsEmptyRows);
}

/**
 * Helper to get column letter from index
 */
function columnIndexToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Helper to create a named range
 * @param {Sheet} sheet - The sheet to create the named range in
 * @param {string} name - Name of the range
 * @param {string} a1Notation - A1 notation for the range (e.g., "B3")
 */
function createNamedRange(sheet, name, a1Notation) {
  try {
    const ss = sheet.getParent();
    const existingRange = ss.getNamedRanges().find(nr => nr.getName() === name);

    if (existingRange) {
      existingRange.remove();
    }

    const range = sheet.getRange(a1Notation);
    ss.setNamedRange(name, range);
  } catch (error) {
    Logger.log(`Error creating named range ${name}: ${error.message}`);
  }
}


// ========================================
// MENU INTEGRATION
// ========================================
// Menu is now managed by the unified SheetLink Recipes menu system
