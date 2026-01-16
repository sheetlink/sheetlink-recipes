# SheetLink Recipes

Open-source financial analysis recipes for Google Sheets. Install pre-built budget trackers, cash flow forecasts, and business reporting tools directly into your spreadsheet.

## What are Recipes?

Recipes are self-contained Google Apps Script files that add powerful financial analysis features to your SheetLink-powered spreadsheet. Each recipe:

- **Runs entirely in your spreadsheet** - No data leaves your Google account
- **Is fully open source** - Audit the code before installing
- **Works with your transaction data** - Analyzes data synced by SheetLink
- **Adds custom menu items** - One-click execution from your spreadsheet

## Available Recipes

### Budget Tracker
Track spending vs budget by category with multi-month trending.

**Outputs:** `Budget Monthly` sheet
**Use Case:** Personal budgeting, expense tracking
[View Code](recipes/official/budget-tracker/) | [Documentation](recipes/official/budget-tracker/README.md)

---

### Budget Tracker (by Account)
Track spending by category broken down by individual accounts.

**Outputs:** `Budget Monthly (by Account)` sheet
**Use Case:** Multi-account budgeting, understanding which accounts drive spending
[View Code](recipes/official/budget-by-account/) | [Documentation](recipes/official/budget-by-account/README.md)

---

### Recurring Spend Detector
Identify subscriptions and recurring charges with trends and annualized costs.

**Outputs:** `Recurring Analysis` sheet
**Use Case:** Subscription auditing, finding hidden recurring charges
[View Code](recipes/official/recurring-analysis/) | [Documentation](recipes/official/recurring-analysis/README.md)

---

### Cash Flow Forecast
Weekly cash flow projection with income, expenses, and running balance.

**Outputs:** `CashFlow Weekly` sheet
**Use Case:** Cash runway planning, liquidity management
[View Code](recipes/official/cash-flow/) | [Documentation](recipes/official/cash-flow/README.md)

---

### Financial Statements Suite
Professional financial reporting: Chart of Accounts, General Ledger, P&L, Balance Sheet, Cash Flow Statement.

**Outputs:** `Chart of Accounts`, `General Ledger`, `Financial Statements` sheets
**Use Case:** Small business accounting, freelancer financials, GAAP-style reporting
[View Code](recipes/official/financial-statements/) | [Documentation](recipes/official/financial-statements/README.md)

---

## Installation

### Option 1: Install via SheetLink Extension (Recommended)

1. Install the [SheetLink Chrome Extension](https://chrome.google.com/webstore/detail/sheetlink)
2. Sync your transactions to a Google Sheet
3. Click the extension icon → **Recipes** tab
4. Browse recipes and click **Install**
5. Grant Apps Script permissions when prompted
6. Run recipes from the **SheetLink Recipes** menu in your spreadsheet

### Option 2: Manual Installation

1. Open your SheetLink-powered Google Sheet
2. Go to **Extensions** → **Apps Script**
3. Delete any existing code in `Code.gs`
4. Copy the contents of a recipe's `recipe.gs` file
5. Paste into Apps Script editor
6. Click **Save** (💾)
7. Refresh your spreadsheet
8. Look for the recipe menu item under **SheetLink Recipes**

## Privacy & Security

SheetLink Recipes are designed with privacy first:

- **All code runs in YOUR Google Sheet** - No data sent to external servers
- **Fully open source** - Audit every line before installing
- **No tracking or analytics** - We never see your financial data
- **Standalone files** - No hidden dependencies or network calls

The SheetLink extension only installs the code to your spreadsheet. It never reads your transaction data.

## Development

### Recipe Structure

Each recipe is a standalone directory containing:

```
recipe-name/
├── recipe.gs          # Complete Apps Script code (with utilities inlined)
├── metadata.json      # Recipe metadata for the marketplace
└── README.md          # User-facing documentation
```

### Contributing

We welcome community contributions! See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines on:

- Creating new recipes
- Recipe coding standards
- Testing and submission process
- Community recipe guidelines

## Requirements

- **SheetLink Extension** - Syncs bank transactions to Google Sheets
- **Transaction Data** - Recipes analyze your synced transaction history
- **Google Apps Script** - Recipes run as Apps Script projects bound to your spreadsheet

## Support

- **Documentation:** [sheetlink.app/recipes](https://sheetlink.app/recipes)
- **Issues:** [GitHub Issues](https://github.com/sheetlink/sheetlink-recipes/issues)
- **Email:** support@sheetlink.app

## License

MIT License - see [LICENSE](LICENSE) for details.

## Credits

Built by the SheetLink team.

Recipes powered by:
- Google Apps Script
- Plaid transaction categories
- Your financial data (that never leaves your control)

---

**Ready to get started?** Install the [SheetLink Extension](https://chrome.google.com/webstore/detail/sheetlink) and start analyzing your finances with recipes!
