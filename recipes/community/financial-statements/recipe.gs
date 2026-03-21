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

    // Create or get Category Mapping sheet
    const mappingSheet = getOrCreateSheet(ss, "Category Mapping");
    setupCategoryMapping(mappingSheet, transactionsSheet, headerMap, ss);

    // Create General Ledger
    const ledgerSheet = getOrCreateSheet(ss, "General Ledger");
    setupGeneralLedgerV3(ledgerSheet, transactionsSheet, headerMap, mappingSheet, ss);

    // Create consolidated Financial Statements
    const statementsSheet = getOrCreateSheet(ss, "Financial Statements");
    setupFinancialStatementsV3(statementsSheet, ledgerSheet, transactionsSheet, headerMap, mappingSheet, ss);

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

  // Row 3: Extended instructions
  sheet.getRange("A3")
    .setValue("You can reuse the same Custom Category across multiple Plaid categories. You can also add custom categories — just note that any custom category not tied to a Plaid category must be entered manually in the General Ledger (including Type and Statement). After editing mappings, re-run the recipe to rebuild the Financial Statements tab.")
    .setFontSize(10)
    .setFontColor("#6b7280")
    .setWrap(true)
    .setHorizontalAlignment("left");
  sheet.setRowHeight(3, 50);

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
  if (sheet.getLastRow() >= 5) {
    try {
      const existingData = sheet.getRange(5, 1, sheet.getLastRow() - 4, 4).getValues();
      existingData.forEach(row => {
        const plaidCategory = row[0];
        if (plaidCategory && plaidCategory !== "") {
          existingMappings[plaidCategory] = {
            customCategory: row[1],
            type: row[2],
            statement: row[3]
          };
        }
      });
    } catch (e) {
      // If reading fails, use defaults
    }
  }

  // Build final mappings
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

  // Clear old data rows only (preserve header rows)
  if (sheet.getLastRow() >= 5) {
    sheet.getRange(5, 1, sheet.getLastRow() - 4, 4).clearContent().clearFormat();
  }

  sheet.getRange(5, 1, mappings.length, 4).setValues(mappings);
  // Plaid Category (col A) read-only feel — no background
  // Custom Category, Type, Statement — yellow to signal editable
  sheet.getRange(5, 2, mappings.length, 3).setBackground("#fffbea");

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setFrozenRows(4);
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

  // Row 3: Override instructions
  sheet.getRange("A3")
    .setValue("Custom Category, Type, and Statement (yellow columns) are auto-mapped from the Category Mapping tab. You can override any cell directly — manual overrides will persist when the recipe is re-run. To reset a cell back to auto-mapping, clear it and re-run the recipe.")
    .setFontSize(10)
    .setFontColor("#6b7280")
    .setWrap(true)
    .setHorizontalAlignment("left");
  sheet.setRowHeight(3, 50);
  currentRow++;

  // Row 4: Blank
  currentRow++;

  // Row 5: Account Balance Configuration
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

    // Build formulas batch
    const formulas = [];
    const overrideCells = []; // Track cells that need manual override values applied after formula write

    for (let i = 0; i < numTxns; i++) {
      const txnRow = i + 2;
      const ledgerRow = dataStartRow + i;
      const txnId = txnIdValues[i][0];
      const override = manualOverrides[txnId] || {};

      // Mapping sheet data range: A5:D down to last row
      const mappingLastRow = mappingSheet.getLastRow();
      const mappingRange = `'Category Mapping'!$A$5:$D$${Math.max(mappingLastRow, 5)}`;

      const row = [
        // A: Date
        `=DATEVALUE(Transactions!${dateColLetter}${txnRow})`,
        // B: Transaction ID
        `=Transactions!${txnIdColLetter}${txnRow}`,
        // C: Vendor
        `=Transactions!${merchantColLetter}${txnRow}`,
        // D: Plaid Category (direct reference)
        `=Transactions!${categoryPrimaryColLetter}${txnRow}`,
        // E: Custom Category — XLOOKUP Plaid Category against mapping col A, return col B
        `=IFERROR(XLOOKUP(D${ledgerRow}, ${mappingRange.replace('$D', '$A').replace('A$5:$D', 'A$5:$A')}, ${mappingRange.replace('$A$5:$D', '$B$5:$B')}), "Uncategorized")`,
        // F: Type — XLOOKUP Custom Category against mapping col B, return col C
        `=IFERROR(XLOOKUP(E${ledgerRow}, ${mappingRange.replace('$A$5:$D', '$B$5:$B')}, ${mappingRange.replace('$A$5:$D', '$C$5:$C')}), "Expense")`,
        // G: Statement — XLOOKUP Custom Category against mapping col B, return col D
        `=IFERROR(XLOOKUP(E${ledgerRow}, ${mappingRange.replace('$A$5:$D', '$B$5:$B')}, ${mappingRange.replace('$A$5:$D', '$D$5:$D')}), "P&L")`,
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
      ];

      formulas.push(row);

      // Track overrides to apply after formula write
      if (override.customCategory || override.type || override.statement) {
        overrideCells.push({ ledgerRow, override });
      }

      if (formulas.length >= 1000 || i === numTxns - 1) {
        const startRow = dataStartRow + i - formulas.length + 1;
        sheet.getRange(startRow, 1, formulas.length, 12).setFormulas(formulas);
        formulas.length = 0;
        SpreadsheetApp.flush();
      }
    }

    // Apply yellow background to Custom Category (E), Type (F), Statement (G) columns
    sheet.getRange(dataStartRow, 5, numTxns, 3).setBackground("#fffbea");

    // Format date and currency columns
    sheet.getRange(dataStartRow, 1, numTxns, 1).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(dataStartRow, 10, numTxns, 2).setNumberFormat("$#,##0.00"); // Debit (J) and Credit (K)

    // Re-apply manual overrides as values (overwrite formulas for overridden cells only)
    overrideCells.forEach(({ ledgerRow, override }) => {
      if (override.customCategory) {
        sheet.getRange(ledgerRow, 5).setValue(override.customCategory);
      }
      if (override.type) {
        sheet.getRange(ledgerRow, 6).setValue(override.type);
      }
      if (override.statement) {
        sheet.getRange(ledgerRow, 7).setValue(override.statement);
      }
    });
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
}

/**
 * Setup Financial Statements v3 - Consolidated with Monthly Trending
 * GL columns: A=Date, B=TxnID, C=Vendor, D=Plaid Category, E=Custom Category, F=Type, G=Statement, H=Account Name, I=Account Type, J=Debit, K=Credit, L=Memo
 * SUMIFS reference: Category=col E ($E:$E), Type=col F ($F:$F), Statement=col G ($G:$G), Debit=col J ($J:$J), Credit=col K ($K:$K), Date=col A ($A:$A), Account=col H ($H:$H)
 */
function setupFinancialStatementsV3(sheet, ledgerSheet, transactionsSheet, headerMap, mappingSheet, ss) {
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

  // Add date headers individually to control formatting
  months.forEach((month, index) => {
    const parts = month.split('-');
    const year = parseInt(parts[0]);
    const monthNum = parseInt(parts[1]);
    const col = 3 + index;

    // Set as string date "YYYY-MM-01" which Sheets will parse as date
    const dateString = `${year}-${String(monthNum).padStart(2, '0')}-01`;
    sheet.getRange(currentRow, col).setValue(dateString);
    sheet.getRange(currentRow, col).setNumberFormat("mmm-yy");
    sheet.getRange(currentRow, col).setFontWeight("bold").setBackground("#f3f3f3");

    monthDates.push(dateString);
  });

  const plHeaderRow = currentRow;
  currentRow++;

  // REVENUE section
  sheet.getRange(currentRow, 1).setValue("REVENUE").setFontWeight("bold");
  currentRow++;

  // Revenue rows (pull from Category Mapping dynamically)
  const plStartRow = currentRow;

  // Category Mapping columns: A=Plaid Category, B=Custom Category, C=Type, D=Statement
  const mappingData = mappingSheet.getLastRow() >= 5
    ? mappingSheet.getRange(5, 1, mappingSheet.getLastRow() - 4, 4).getValues()
    : [];

  const revenueAccountsMap = {}; // customCategory -> type
  for (let i = 0; i < mappingData.length; i++) {
    const type      = mappingData[i][2]; // Col C: Type
    const customCat = mappingData[i][1]; // Col B: Custom Category
    const statement = mappingData[i][3]; // Col D: Statement
    if (type === "Revenue" && statement === "P&L" && customCat && !revenueAccountsMap[customCat]) {
      revenueAccountsMap[customCat] = type;
    }
  }

  Object.keys(revenueAccountsMap).forEach((accountName) => {
    const type = revenueAccountsMap[accountName];

    sheet.getRange(currentRow, 1).setValue(accountName);
    sheet.getRange(currentRow, 2).setValue(type);

    months.forEach((month, idx) => {
      const headerCol = columnIndexToLetter(3 + idx);
      const accountRef = `$A${currentRow}`;
      const typeRef = `$B${currentRow}`;

      // Revenue = Debits - Credits (GL col E=Custom Category, F=Type, J=Debit, K=Credit, A=Date)
      sheet.getRange(currentRow, 3 + idx).setFormula(
        `=SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$E:$E, ${accountRef}, 'General Ledger'!$F:$F, ${typeRef}, ` +
        `'General Ledger'!$A:$A, ">="&${headerCol}$${plHeaderRow}, ` +
        `'General Ledger'!$A:$A, "<"&EOMONTH(${headerCol}$${plHeaderRow}, 0)+1) - ` +
        `SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$E:$E, ${accountRef}, 'General Ledger'!$F:$F, ${typeRef}, ` +
        `'General Ledger'!$A:$A, ">="&${headerCol}$${plHeaderRow}, ` +
        `'General Ledger'!$A:$A, "<"&EOMONTH(${headerCol}$${plHeaderRow}, 0)+1)`
      );
    });
    currentRow++;
  });

  // Total Revenue
  sheet.getRange(currentRow, 1).setValue("Total Revenue").setFontWeight("bold");
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(`=SUM(${col}${plStartRow}:${col}${currentRow - 1})`);
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
  for (let i = 0; i < mappingData.length; i++) {
    const type      = mappingData[i][2];
    const customCat = mappingData[i][1];
    const statement = mappingData[i][3];
    if (type === "Expense" && statement === "P&L" && customCat && !expenseAccountsMap[customCat]) {
      expenseAccountsMap[customCat] = type;
    }
  }

  Object.keys(expenseAccountsMap).forEach((accountName) => {
    const type = expenseAccountsMap[accountName];

    sheet.getRange(currentRow, 1).setValue(accountName);
    sheet.getRange(currentRow, 2).setValue(type);

    months.forEach((month, idx) => {
      const headerCol = columnIndexToLetter(3 + idx);
      const accountRef = `$A${currentRow}`;
      const typeRef = `$B${currentRow}`;

      // Expenses = Debits - Credits (GL col E=Custom Category, F=Type, J=Debit, K=Credit)
      sheet.getRange(currentRow, 3 + idx).setFormula(
        `=SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$E:$E, ${accountRef}, 'General Ledger'!$F:$F, ${typeRef}, ` +
        `'General Ledger'!$A:$A, ">="&${headerCol}$${plHeaderRow}, ` +
        `'General Ledger'!$A:$A, "<"&EOMONTH(${headerCol}$${plHeaderRow}, 0)+1) - ` +
        `SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$E:$E, ${accountRef}, 'General Ledger'!$F:$F, ${typeRef}, ` +
        `'General Ledger'!$A:$A, ">="&${headerCol}$${plHeaderRow}, ` +
        `'General Ledger'!$A:$A, "<"&EOMONTH(${headerCol}$${plHeaderRow}, 0)+1)`
      );
    });
    currentRow++;
  });

  // Total Expenses
  sheet.getRange(currentRow, 1).setValue("Total Expenses").setFontWeight("bold");
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(`=SUM(${col}${expenseStartRow}:${col}${currentRow - 1})`);
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

  // Add date headers individually to control formatting
  months.forEach((month, index) => {
    const parts = month.split('-');
    const year = parseInt(parts[0]);
    const monthNum = parseInt(parts[1]);
    const col = 3 + index;

    // Set as string date "YYYY-MM-01" which Sheets will parse as date
    const dateString = `${year}-${String(monthNum).padStart(2, '0')}-01`;
    sheet.getRange(currentRow, col).setValue(dateString);
    sheet.getRange(currentRow, col).setNumberFormat("mmm-yy");
    sheet.getRange(currentRow, col).setFontWeight("bold").setBackground("#f3f3f3");
  });

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

  configData.forEach((row, idx) => {
    const accountName = row[0];
    const accountType = row[1];
    const startingBalance = row[2];
    const asOfDate = row[3];

    // Skip empty rows or non-Asset accounts
    if (!accountName || accountType !== "Asset") return;

    // Track this row for Cash Flow reconciliation
    cashAccountBalanceSheetRows[accountName] = currentRow;

    sheet.getRange(currentRow, 1).setValue(accountName);
    sheet.getRange(currentRow, 2).setValue("Asset");

    // Monthly balances - date-aware calculation using General Ledger
    months.forEach((month, monthIdx) => {
      const headerCol = columnIndexToLetter(3 + monthIdx);
      const accountRef = `$A${currentRow}`; // Reference to Account Name in column A
      const monthEndDate = `EOMONTH(${headerCol}$${bsHeaderRow}, 0)`; // Last day of month

      // Date-aware formula for Assets:
      // 1. If month end = as-of date: show exact starting balance
      // 2. If month end < as-of date: work backwards (reverse transactions between month-end and as-of date)
      // 3. If month end > as-of date: work forwards (add net change from as-of to month-end)
      // GL columns: A=Date, H=Account Name, J=Debit, K=Credit
      const formula =
        `=IF(INT(${monthEndDate})=INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0)), ` +
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

      sheet.getRange(currentRow, 3 + monthIdx).setFormula(formula);
    });
    currentRow++;
  });

  // Total Assets
  sheet.getRange(currentRow, 1).setValue("Total Assets").setFontWeight("bold");
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(`=SUM(${col}${assetsStartRow}:${col}${currentRow - 1})`);
  });
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

  configData.forEach((row, idx) => {
    const accountName = row[0];
    const accountType = row[1];
    const startingBalance = row[2];
    const asOfDate = row[3];

    // Skip empty rows or non-Liability accounts
    if (!accountName || accountType !== "Liability") return;

    // Track this row for Cash Flow adjustments
    liabilityBalanceSheetRows[accountName] = currentRow;

    sheet.getRange(currentRow, 1).setValue(accountName);
    sheet.getRange(currentRow, 2).setValue("Liability");

    // Monthly balances - date-aware calculation using General Ledger (inverted logic for liabilities)
    months.forEach((month, monthIdx) => {
      const headerCol = columnIndexToLetter(3 + monthIdx);
      const accountRef = `$A${currentRow}`; // Reference to Account Name in column A
      const monthEndDate = `EOMONTH(${headerCol}$${bsHeaderRow}, 0)`; // Last day of month

      // GL columns: A=Date, H=Account Name, J=Debit, K=Credit (inverted logic for liabilities)
      const formula =
        `=IF(INT(${monthEndDate})=INT(IFERROR(VLOOKUP(${accountRef}, 'General Ledger'!$A$${accountConfigStartRow}:$D$${accountConfigEndRow}, 4, FALSE), 0)), ` +
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
        `'General Ledger'!$A:$A, "<="&INT(${monthEndDate}))))`);

      sheet.getRange(currentRow, 3 + monthIdx).setFormula(formula);
    });
    currentRow++;
  });

  // Total Liabilities
  sheet.getRange(currentRow, 1).setValue("Total Liabilities").setFontWeight("bold");
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(`=SUM(${col}${liabilitiesStartRow}:${col}${currentRow - 1})`);
  });
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#f4cccc").setFontWeight("bold");
  const totalLiabilitiesRow = currentRow;
  currentRow++;

  currentRow++; // Blank row

  // EQUITY (calculated as Assets - Liabilities)
  sheet.getRange(currentRow, 1).setValue("EQUITY").setFontWeight("bold");
  currentRow++;

  const equityStartRow = currentRow;

  // Retained Earnings (will reference Total Equity row calculated below)
  sheet.getRange(currentRow, 1).setValue("Retained Earnings");
  sheet.getRange(currentRow, 2).setValue("Equity");
  const retainedEarningsRow = currentRow;
  // Formula will be added after Total Equity row is defined
  currentRow++;

  // Total Equity
  sheet.getRange(currentRow, 1).setValue("Total Equity").setFontWeight("bold");
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(`=${col}${totalAssetsRow}-${col}${totalLiabilitiesRow}`);
  });
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#d9d2e9").setFontWeight("bold");
  const totalEquityRow = currentRow;
  currentRow++;

  // Now go back and set Retained Earnings to equal Total Equity
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(retainedEarningsRow, 3 + idx).setFormula(`=${col}${totalEquityRow}`);
  });

  // Check: Total Liabilities + Equity should equal Total Assets
  sheet.getRange(currentRow, 1).setValue("Check: Liabilities + Equity").setFontStyle("italic");
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(`=${col}${totalLiabilitiesRow}+${col}${totalEquityRow}`);
  });
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

  // Add date headers individually to control formatting
  months.forEach((month, index) => {
    const parts = month.split('-');
    const year = parseInt(parts[0]);
    const monthNum = parseInt(parts[1]);
    const col = 3 + index;

    // Set as string date "YYYY-MM-01" which Sheets will parse as date
    const dateString = `${year}-${String(monthNum).padStart(2, '0')}-01`;
    sheet.getRange(currentRow, col).setValue(dateString);
    sheet.getRange(currentRow, col).setNumberFormat("mmm-yy");
    sheet.getRange(currentRow, col).setFontWeight("bold").setBackground("#f3f3f3");
  });

  const cfHeaderRow = currentRow;
  currentRow++;

  // OPERATING ACTIVITIES (Indirect Method)
  sheet.getRange(currentRow, 1).setValue("OPERATING ACTIVITIES").setFontWeight("bold");
  currentRow++;

  const cfOperatingStartRow = currentRow;

  // Net Income (from P&L)
  sheet.getRange(currentRow, 1).setValue("Net Income");
  sheet.getRange(currentRow, 2).setValue("Operating");
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(`=${col}${netIncomeRow}`);
  });
  const netIncomeRowCF = currentRow;
  currentRow++;

  // Adjustments for changes in working capital
  sheet.getRange(currentRow, 1).setValue("Changes in Working Capital:").setFontStyle("italic");
  currentRow++;

  // Changes in Liabilities (increase = source of cash, add back)
  // For each liability account, sum credits - debits from General Ledger
  Object.keys(liabilityBalanceSheetRows).forEach(accountName => {
    const bsRowNum = liabilityBalanceSheetRows[accountName];

    sheet.getRange(currentRow, 1).setValue(`  Increase in ${accountName}`);
    sheet.getRange(currentRow, 2).setValue("Operating");

    months.forEach((m, monthIdx) => {
      const headerCol = columnIndexToLetter(3 + monthIdx);
      const accountCell = `$A${bsRowNum}`; // Account Name from Balance Sheet (absolute column reference)

      // Calculate net change from General Ledger: Credits - Debits (increase is positive)
      sheet.getRange(currentRow, 3 + monthIdx).setFormula(
        `=SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$H:$H, ${accountCell}, ` +
        `'General Ledger'!$A:$A, ">="&${headerCol}$${cfHeaderRow}, ` +
        `'General Ledger'!$A:$A, "<="&INT(EOMONTH(${headerCol}$${cfHeaderRow}, 0))) - ` +
        `SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$H:$H, ${accountCell}, ` +
        `'General Ledger'!$A:$A, ">="&${headerCol}$${cfHeaderRow}, ` +
        `'General Ledger'!$A:$A, "<="&INT(EOMONTH(${headerCol}$${cfHeaderRow}, 0)))`
      );
    });

    currentRow++;
  });

  // Other Working Capital (plug to reconcile: Total Cash Change - Net Income - Liabilities - CapEx - Loan Proceeds)
  sheet.getRange(currentRow, 1).setValue("  Other Working Capital");
  sheet.getRange(currentRow, 2).setValue("Operating");
  const otherWorkingCapitalRow = currentRow;
  // Formula will be added after we know CapEx and Loan Proceeds rows
  currentRow++;

  // Cash from Operating Activities
  sheet.getRange(currentRow, 1).setValue("Cash from Operating Activities").setFontWeight("bold");
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(`=SUM(${col}${cfOperatingStartRow}:${col}${currentRow - 1})`);
  });
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#d9ead3").setFontWeight("bold");
  const totalCFOperatingRow = currentRow;
  currentRow++;

  currentRow++; // Blank row

  // INVESTING ACTIVITIES
  sheet.getRange(currentRow, 1).setValue("INVESTING ACTIVITIES").setFontWeight("bold");
  currentRow++;

  const cfInvestingStartRow = currentRow;

  // Manual input for capital expenditures
  sheet.getRange(currentRow, 1).setValue("Capital Expenditures");
  sheet.getRange(currentRow, 2).setValue("Investing");
  months.forEach((m, idx) => {
    sheet.getRange(currentRow, 3 + idx).setValue(0).setBackground("#fffbea");
  });
  const capExRow = currentRow;
  currentRow++;

  // Cash from Investing Activities
  sheet.getRange(currentRow, 1).setValue("Cash from Investing Activities").setFontWeight("bold");
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(`=SUM(${col}${cfInvestingStartRow}:${col}${currentRow - 1})`);
  });
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#d9ead3").setFontWeight("bold");
  const totalCFInvestingRow = currentRow;
  currentRow++;

  currentRow++; // Blank row

  // FINANCING ACTIVITIES
  sheet.getRange(currentRow, 1).setValue("FINANCING ACTIVITIES").setFontWeight("bold");
  currentRow++;

  const cfFinancingStartRow = currentRow;

  // Manual input for loan proceeds/repayments
  sheet.getRange(currentRow, 1).setValue("Loan Proceeds / Repayments");
  sheet.getRange(currentRow, 2).setValue("Financing");
  months.forEach((m, idx) => {
    sheet.getRange(currentRow, 3 + idx).setValue(0).setBackground("#fffbea");
  });
  const loanProceedsRow = currentRow;
  currentRow++;

  // Cash from Financing Activities
  sheet.getRange(currentRow, 1).setValue("Cash from Financing Activities").setFontWeight("bold");
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(`=SUM(${col}${cfFinancingStartRow}:${col}${currentRow - 1})`);
  });
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#d9ead3").setFontWeight("bold");
  const totalCFFinancingRow = currentRow;
  currentRow++;

  currentRow++; // Blank row

  // NET CHANGE IN CASH
  sheet.getRange(currentRow, 1).setValue("NET CHANGE IN CASH").setFontWeight("bold").setFontSize(11);
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    sheet.getRange(currentRow, 3 + idx).setFormula(
      `=${col}${totalCFOperatingRow}+${col}${totalCFInvestingRow}+${col}${totalCFFinancingRow}`
    );
  });
  sheet.getRange(currentRow, 1, 1, 2 + months.length).setBackground("#fce5cd").setFontWeight("bold");
  const netChangeInCashRow = currentRow;
  currentRow++;

  currentRow++; // Blank row

  // RECONCILIATION - Change in Cash from Balance Sheet
  sheet.getRange(currentRow, 1).setValue("Change in Cash (from Balance Sheet):").setFontStyle("italic");
  months.forEach((m, monthIdx) => {
    const col = columnIndexToLetter(3 + monthIdx);
    const prevCol = monthIdx > 0 ? columnIndexToLetter(3 + monthIdx - 1) : null;

    // Skip first month (no previous month to compare)
    if (monthIdx === 0) {
      sheet.getRange(currentRow, 3 + monthIdx).setValue("");
    } else {
      // Sum all changes in cash accounts (Asset accounts)
      let formula = "=";
      const cashAccountNames = Object.keys(cashAccountBalanceSheetRows);

      if (cashAccountNames.length > 0) {
        // For each cash account, add (current month - previous month)
        const changes = cashAccountNames.map(accountName => {
          const bsRowNum = cashAccountBalanceSheetRows[accountName];
          return `(${col}${bsRowNum}-${prevCol}${bsRowNum})`;
        });
        formula += changes.join("+");
      } else {
        // No cash accounts - show 0
        formula = "0";
      }

      sheet.getRange(currentRow, 3 + monthIdx).setFormula(formula);
    }
  });
  currentRow++;

  // RECONCILIATION - Difference
  sheet.getRange(currentRow, 1).setValue("Difference (to investigate):").setFontStyle("italic").setFontColor("#cc0000");
  months.forEach((m, idx) => {
    const col = columnIndexToLetter(3 + idx);
    // Skip first month (no previous month to compare)
    if (idx === 0) {
      sheet.getRange(currentRow, 3 + idx).setValue("");
    } else {
      sheet.getRange(currentRow, 3 + idx).setFormula(
        `=${col}${netChangeInCashRow}-${col}${currentRow - 1}`
      );
    }
  });
  sheet.getRange(currentRow, 3, 1, months.length).setFontColor("#cc0000");
  const changeFromBalanceSheetRow = currentRow - 1;
  currentRow++;

  // Now go back and fill in "Other Working Capital" formula (plug to reconcile)
  // Formula: Change in Cash (from GL) - Net Income - Sum of Liability Increases - CapEx - Loan Proceeds
  months.forEach((m, monthIdx) => {
    const col = columnIndexToLetter(3 + monthIdx);
    const headerCol = columnIndexToLetter(3 + monthIdx);

    // Build formula to calculate change in cash from General Ledger for all cash accounts
    const cashAccountNames = Object.keys(cashAccountBalanceSheetRows);
    let cashChangeFormula = "";
    if (cashAccountNames.length > 0) {
      // For each cash account, sum (Debits - Credits) from General Ledger
      const cashChanges = cashAccountNames.map(accountName => {
        const bsRowNum = cashAccountBalanceSheetRows[accountName];
        return `(SUMIFS('General Ledger'!$J:$J, 'General Ledger'!$H:$H, $A${bsRowNum}, ` +
               `'General Ledger'!$A:$A, ">="&${headerCol}$${cfHeaderRow}, ` +
               `'General Ledger'!$A:$A, "<="&INT(EOMONTH(${headerCol}$${cfHeaderRow}, 0))) - ` +
               `SUMIFS('General Ledger'!$K:$K, 'General Ledger'!$H:$H, $A${bsRowNum}, ` +
               `'General Ledger'!$A:$A, ">="&${headerCol}$${cfHeaderRow}, ` +
               `'General Ledger'!$A:$A, "<="&INT(EOMONTH(${headerCol}$${cfHeaderRow}, 0))))`;
      });
      cashChangeFormula = cashChanges.join("+");
    } else {
      cashChangeFormula = "0";
    }

    // Build formula to sum all liability increase rows
    const liabilityAccountNames = Object.keys(liabilityBalanceSheetRows);
    let liabilitySum = "";
    if (liabilityAccountNames.length > 0) {
      // Find the range of rows that contain liability increases
      // These are between cfOperatingStartRow+1 and otherWorkingCapitalRow-1
      const firstLiabilityRow = cfOperatingStartRow + 1; // After "Net Income"
      const lastLiabilityRow = otherWorkingCapitalRow - 1; // Before "Other Working Capital"
      liabilitySum = `SUM(${col}${firstLiabilityRow}:${col}${lastLiabilityRow})`;
    } else {
      liabilitySum = "0";
    }

    // Other Working Capital = Change in Cash (from GL) - Net Income - Liability Increases - CapEx - Loan Proceeds
    sheet.getRange(otherWorkingCapitalRow, 3 + monthIdx).setFormula(
      `=${cashChangeFormula}-${col}${cfOperatingStartRow}-${liabilitySum}-${col}${capExRow}-${col}${loanProceedsRow}`
    );
  });

  // Format Cash Flow currency
  sheet.getRange(cfOperatingStartRow, 3, currentRow - cfOperatingStartRow, months.length).setNumberFormat("$#,##0.00_);($#,##0.00);\"-  \"");

  // Set column widths
  sheet.setColumnWidth(1, 250); // Account
  sheet.setColumnWidth(2, 100); // Type
  months.forEach((m, idx) => {
    sheet.setColumnWidth(3 + idx, 100); // Monthly columns
  });

  // Freeze after column B (columns A and B frozen)
  sheet.setFrozenRows(plHeaderRow);
  sheet.setFrozenColumns(2);
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
