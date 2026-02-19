# Contributing to SheetLink Recipes

Thank you for your interest in contributing! This guide covers everything you need to build and submit a recipe.

## How Recipes Work

When a user installs a recipe from the SheetLink extension, two Apps Script files are written to their spreadsheet project:

1. **`utils.gs`** — Shared utilities from `_shared/utils.gs` in this repo. Installed once, shared by all recipes.
2. **`recipe.gs`** — Your recipe code, installed as a separate file.

Because `utils.gs` is always present, your recipe code can call any utility function directly without defining it. **Do not redefine or copy utility functions** — the installer strips any inlined copies to avoid duplication errors.

---

## File Structure

```
recipes/community/your-recipe-name/
├── recipe.gs        # Your Apps Script code
├── metadata.json    # Recipe metadata (required for manifest)
└── README.md        # User-facing docs (optional but encouraged)
```

---

## Writing `recipe.gs`

### Entry function

Your recipe must have exactly one entry function named `run<Something>Recipe`:

```javascript
function runMyRecipe(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  // ...
}
```

The installer auto-detects this function and generates a menu entry point that calls it. You do not need to define `onOpen()` or manage the menu yourself.

### Using `utils.gs` functions

All functions in `_shared/utils.gs` are globally available. Call them directly:

```javascript
function runMyRecipe(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  // Validate transactions exist before doing any work
  if (!checkTransactionsOrPrompt(ss)) return;

  const txSheet = getTransactionsSheet(ss);
  const headerMap = getHeaderMap(txSheet);
  const transactions = getTransactionData(txSheet);

  // ... your logic ...

  const outputSheet = getOrCreateSheet(ss, "My Output");
  setHeaders(outputSheet, ["Month", "Amount", "Count"]);
  formatSheet(outputSheet);

  showToast("Done!", "My Recipe", 3);
}
```

### Full `utils.gs` API reference

**Transaction access**
| Function | Returns | Description |
|----------|---------|-------------|
| `getTransactionsSheet(ss)` | `Sheet\|null` | Gets the "Transactions" sheet |
| `validateTransactionsSheet(ss)` | `{valid, error}` | Checks sheet exists and has data |
| `checkTransactionsOrPrompt(ss)` | `boolean` | Validates or shows a "no data" dialog — returns `false` if user cancels |
| `getTransactionData(txSheet)` | `Object[]` | All rows as objects keyed by header name |

**Header/column utilities**
| Function | Returns | Description |
|----------|---------|-------------|
| `getHeaderMap(sheet)` | `Object` | Maps column names → 1-based column index |
| `getColumnIndex(headerMap, name)` | `number\|null` | Gets column index from header map |

**Sheet management**
| Function | Description |
|----------|-------------|
| `getOrCreateSheet(ss, name)` | Gets sheet by name, creates it (at end) if missing |
| `clearSheetData(sheet, preserveHeaders)` | Clears data; set `preserveHeaders=true` to keep row 1 |
| `setHeaders(sheet, headersArray)` | Writes bold headers to row 1 with gray background |
| `formatSheet(sheet)` | Freezes row 1 and auto-resizes all columns |
| `createNamedRange(sheet, name, a1)` | Creates/replaces a named range |

**Date utilities**
| Function | Returns | Description |
|----------|---------|-------------|
| `parseDate(value)` | `Date\|null` | Parses string or Date; returns null if invalid |
| `getCurrentMonth()` | `string` | `"YYYY-MM"` for today |
| `getMostRecentMonth(transactions)` | `string` | `"YYYY-MM"` from the most recent transaction date |
| `isInMonth(date, targetMonth)` | `boolean` | Whether a date falls in `"YYYY-MM"` |
| `getISOWeek(date)` | `string` | `"YYYY-WW"` ISO week number |

**Data formatting**
| Function | Returns | Description |
|----------|---------|-------------|
| `formatCurrency(value)` | `string` | `"$12.34"` (absolute value, 2 decimal places) |

**Data type fixups**

Dates and the `pending` column are written as text strings by the extension (RAW mode). If your recipe uses these for formulas or comparisons, call these before processing:

| Function | Description |
|----------|-------------|
| `formatTransactionDateColumns(txSheet, headerMap)` | Strips leading apostrophes from `date` and `authorized_date` columns |
| `formatTransactionPendingColumn(txSheet, headerMap)` | Converts `"TRUE"`/`"FALSE"` strings to actual booleans in `pending` column |

**Logging & UI**
| Function | Description |
|----------|-------------|
| `showToast(message, title, timeoutSeconds)` | Shows a toast notification in the spreadsheet |
| `showError(message)` | Shows a blocking error dialog |
| `logRecipe(recipeName, message)` | Logs to Apps Script console with `[recipeName]` prefix |

### Transaction data shape

The `Transactions` sheet has these columns (all values are strings unless `formatTransactionDateColumns` / `formatTransactionPendingColumn` are called):

| Column | Type | Notes |
|--------|------|-------|
| `transaction_id` | string | Plaid transaction ID |
| `account_id` | string | Plaid account ID |
| `account_name` | string | Human-readable account name |
| `date` | string | `"YYYY-MM-DD"` — text, call `formatTransactionDateColumns` if using in formulas |
| `authorized_date` | string | `"YYYY-MM-DD"` or empty |
| `name` | string | Raw transaction name |
| `merchant_name` | string | Cleaned merchant name (may be empty) |
| `amount` | number | Positive = expense, negative = income (Plaid convention) |
| `category_primary` | string | e.g., `"Food and Drink"` |
| `category_detailed` | string | e.g., `"Food and Drink > Restaurants"` |
| `payment_channel` | string | `"online"`, `"in store"`, `"other"` |
| `pending` | string | `"TRUE"` or `"FALSE"` — call `formatTransactionPendingColumn` to get booleans |
| `iso_currency_code` | string | e.g., `"USD"` |

> **Amount sign convention:** Plaid uses positive amounts for debits (money out) and negative for credits (money in). Most recipes treat positive as expense and negative as income.

### What NOT to do

- **Do not define** `getOrCreateSheet`, `getTransactionsSheet`, `validateTransactionsSheet`, `getHeaderMap`, `getColumnIndex`, `clearSheetData`, `setHeaders`, `formatSheet`, `getTransactionData`, `parseDate`, `showToast`, `showError`, or `logRecipe` — these exist in `utils.gs` and the installer will conflict
- **Do not define** `TRANSACTIONS_SHEET_NAME` — use `getTransactionsSheet(ss)` instead
- **Do not call** `UrlFetchApp` — recipes must be fully offline
- **Do not define** `onOpen()` — the extension manages the menu

---

## Writing `metadata.json`

This file must be added to `manifest.json` at the repo root for the recipe to appear in the marketplace.

```json
{
  "id": "your-recipe-name",
  "name": "Your Recipe Display Name",
  "version": "1.0.0",
  "author": "Your GitHub username",
  "type": "community",
  "source": "community",
  "githubUser": "your-github-username",
  "contributed": "YYYY-MM-DD",
  "description": "One-line description, max 80 chars",
  "longDescription": "2-3 sentences explaining what the recipe does, who it's for, and what output it creates.",
  "requirements": {
    "sheets": ["Transactions"],
    "columns": ["date", "amount", "category_primary"]
  },
  "outputs": ["My Output Sheet"],
  "files": ["recipe.gs"],
  "dependencies": [],
  "menuName": "🎯 Your Recipe Name",
  "entryFunction": "runMyRecipe",
  "tags": ["budgeting", "analysis"]
}
```

**Field notes:**
- `id` — lowercase, hyphenated, unique across all recipes
- `entryFunction` — must exactly match your `run<X>Recipe` function name
- `menuName` — include an emoji; shown in the SheetLink Recipes menu
- `requirements.columns` — list every column your recipe reads from `Transactions`
- `outputs` — list every sheet name your recipe creates or overwrites
- `tags` — used for filtering; see existing recipes for common tags

---

## Testing Locally

Before submitting, test your recipe manually:

1. Open a Google Sheet that has a `Transactions` sheet with real or dummy data (use SheetLink's "Populate Dummy Data" from the Settings menu)
2. Go to **Extensions → Apps Script**
3. Create a new file, paste the contents of `_shared/utils.gs` followed by your `recipe.gs`
4. Save and run your entry function
5. Check the execution log for errors
6. Verify output sheets are created correctly

**Edge cases to test:**
- Empty `Transactions` sheet
- Missing required columns
- Transactions with null/empty values
- Single transaction vs. many transactions
- Multiple months of data

---

## Submitting a PR

```bash
# 1. Fork and clone
git clone https://github.com/sheetlink/sheetlink-recipes.git

# 2. Create a branch
git checkout -b recipe/your-recipe-name

# 3. Add your files
mkdir -p recipes/community/your-recipe-name
# ... add recipe.gs, metadata.json, README.md

# 4. Add your entry to manifest.json (in the "recipes" array)
# ... edit manifest.json

# 5. Commit and push
git add recipes/community/your-recipe-name/ manifest.json
git commit -m "feat: Add [Your Recipe Name] community recipe"
git push origin recipe/your-recipe-name
```

Open a PR with:
- **Title:** `feat: Add [Recipe Name] community recipe`
- **What it does** — brief description
- **Columns required** — from the `Transactions` sheet
- **What you tested** — edge cases covered
- **Screenshots** — of the output sheet(s)

We review for privacy compliance (no external calls), code quality, error handling, and documentation. Once merged, the recipe appears in the marketplace automatically.

---

## Recipe Ideas

**Personal finance**
- Tax prep helper — categorize transactions by deduction type
- Savings goals tracker — progress toward named goals
- Debt payoff planner — snowball/avalanche tracking
- Net worth tracker — assets and liabilities over time

**Business**
- Invoice tracker — match expenses to client projects
- Expense report generator — formatted for reimbursement
- 1099 contractor report — track contractor payments
- Mileage/reimbursement tracker

**Analysis**
- Category trends — YoY/MoM spending by category
- Merchant analysis — top merchants by spend with trends
- Seasonal spending patterns
- Account reconciliation

---

## Questions?

- **Discussions:** [github.com/sheetlink/sheetlink-recipes/discussions](https://github.com/sheetlink/sheetlink-recipes/discussions)
- **Issues:** [github.com/sheetlink/sheetlink-recipes/issues](https://github.com/sheetlink/sheetlink-recipes/issues)

By contributing, you agree to license your recipe under the MIT License.
