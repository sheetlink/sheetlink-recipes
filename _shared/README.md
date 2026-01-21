# Shared Utilities

This folder contains utilities that are automatically injected by the SheetLink Recipes installer when recipes are installed to a user's Google Sheet.

## Files

### `utils.gs`
Common utility functions available to all recipes:

- **`TRANSACTIONS_SHEET_NAME`**: Constant for the transactions sheet name ("Transactions")
- **`getOrCreateSheet(ss, sheetName)`**: Get or create a sheet by name
- **`getTransactionsSheet(ss)`**: Get the transactions sheet
- **`validateTransactionsSheet(ss)`**: Validate transactions sheet exists and has data
- **`formatCurrency(value)`**: Format a number as currency
- **`getColumnIndexByHeader(sheet, headerName)`**: Find column index by header name
- **`parseDate(dateStr)`**: Parse date string to Date object
- **`formatDate(date)`**: Format date as YYYY-MM-DD
- **`getMonthKey(date)`**: Get month key in YYYY-MM format
- **`clearSheetKeepHeaders(sheet)`**: Clear sheet data but preserve headers

### `menu-template.gs`
Template for the dynamically generated menu system. The installer modifies this template to add menu items for each installed recipe.

## For Recipe Developers

When creating recipes for SheetLink, you can rely on these utilities being available. **Do not inline these utilities in your recipe files** - they are automatically injected by the installer.

### Recipe Structure

A recipe should contain:
1. Recipe-specific logic and functions
2. A menu entry point function: `function run_<recipe-id>() { ... }`

Do NOT include:
- Utils functions (they're provided by `utils.gs`)
- `onOpen()` function (managed by the menu system)
- Manifest file (created by installer)

### Example Recipe Template

```javascript
/**
 * SheetLink Recipe: My Recipe Name
 * Version: 1.0.0
 * Description: What this recipe does
 */

// ========================================
// RECIPE LOGIC
// ========================================

function runMyRecipe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Use shared utilities
  const validation = validateTransactionsSheet(ss);
  if (!validation.valid) {
    SpreadsheetApp.getUi().alert(validation.error);
    return;
  }

  const txnSheet = getTransactionsSheet(ss);
  const outputSheet = getOrCreateSheet(ss, "My Output");

  // Your recipe logic here...
}

// ========================================
// MENU INTEGRATION
// ========================================

/**
 * Menu entry point for my-recipe
 * Called from SheetLink Recipes menu
 */
function run_my_recipe() {
  return runMyRecipe();
}
```

## Version History

- **2.0.0** (2026-01): Initial shared utilities framework
