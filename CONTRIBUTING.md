# Contributing to SheetLink Recipes

Thank you for your interest in contributing to SheetLink Recipes! This guide will help you create and submit recipes that the community can use.

## Recipe Guidelines

### Privacy & Security First

**All recipes must:**
- Run entirely within Google Apps Script (no external API calls)
- Never send transaction data to external servers
- Be fully auditable (clear, readable code)
- Include no tracking, analytics, or telemetry

**Prohibited:**
- `UrlFetchApp` calls to external services
- Obfuscated or minified code
- Third-party libraries (unless well-known and security-audited)
- Data collection or transmission

### Code Quality

- **Standalone:** Each recipe must be a single `recipe.gs` file with all utilities inlined
- **Documented:** Clear comments explaining logic, especially complex calculations
- **Error Handling:** Graceful handling of missing data, invalid formats, etc.
- **User Feedback:** Use `showToast()` or `showError()` to communicate with users

## Recipe Structure

Each recipe lives in its own directory:

```
recipes/community/your-recipe-name/
├── recipe.gs          # Complete Apps Script code
├── metadata.json      # Recipe metadata
└── README.md          # User-facing documentation
```

### recipe.gs

Must include:

```javascript
/**
 * Recipe Name
 * Version X.X.X
 * Author: Your Name
 *
 * Description of what the recipe does.
 */

/**
 * Entry point function
 * This function will be called from the spreadsheet menu
 */
function runYourRecipe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Your recipe logic here

  showToast("✓ Recipe complete!", "SheetLink Recipes", 3);
}

/**
 * Creates custom menu on spreadsheet open
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('SheetLink Recipes')
    .addItem('🎯 Your Recipe Name', 'runYourRecipe')
    .addToUi();
}

// === Utilities (inline all helper functions here) ===

function showToast(message, title, timeout) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeout);
}

function showError(message) {
  SpreadsheetApp.getUi().alert("Error", message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// Add other utility functions as needed
```

### metadata.json

```json
{
  "id": "your-recipe-name",
  "name": "Your Recipe Display Name",
  "version": "1.0.0",
  "author": "Your Name",
  "type": "community",
  "description": "Short one-line description (max 80 chars)",
  "longDescription": "Longer description explaining what the recipe does, who it's for, and what outputs it creates. 2-3 sentences.",
  "requirements": {
    "sheets": ["Transactions"],
    "columns": ["date", "amount", "category_primary"]
  },
  "outputs": ["Your Output Sheet Name"],
  "files": ["recipe.gs"],
  "dependencies": [],
  "menuName": "🎯 Your Recipe Name",
  "entryFunction": "runYourRecipe",
  "tags": ["tag1", "tag2", "tag3"]
}
```

**Field descriptions:**
- `id`: Lowercase, hyphenated (e.g., `cash-flow`, `tax-report`)
- `type`: Use `"community"` for community recipes
- `requirements.columns`: List all transaction columns your recipe needs
- `outputs`: Names of sheets your recipe creates/updates
- `menuName`: Include an emoji for visual appeal
- `tags`: Help users find your recipe (e.g., `budgeting`, `taxes`, `business`)

### README.md

```markdown
# Recipe Name

Brief description of what the recipe does.

## What It Does

Explain the recipe's purpose and value proposition.

## Requirements

- **Transaction Columns:** `date`, `amount`, `category_primary`
- **Minimum Data:** 30 days of transactions recommended

## Outputs

### Sheet Name

Description of the output sheet and its columns.

| Column | Description |
|--------|-------------|
| Column A | What this column shows |
| Column B | What this column shows |

## How to Use

1. Install the recipe via SheetLink extension
2. Ensure you have synced transactions
3. Go to **SheetLink Recipes** → **🎯 Your Recipe Name**
4. View results in the `Your Output Sheet` tab

## Configuration

Explain any user-configurable settings (if applicable).

## Example Use Cases

- Use case 1
- Use case 2
- Use case 3

## Known Limitations

List any limitations or edge cases users should be aware of.

## Support

Questions? Open an issue at [github.com/sheetlink/sheetlink-recipes/issues](https://github.com/sheetlink/sheetlink-recipes/issues)
```

## Submission Process

### 1. Fork the Repository

```bash
git clone https://github.com/sheetlink/sheetlink-recipes.git
cd sheetlink-recipes
```

### 2. Create Your Recipe

```bash
mkdir -p recipes/community/your-recipe-name
cd recipes/community/your-recipe-name
touch recipe.gs metadata.json README.md
```

### 3. Test Your Recipe

**Manual Testing:**
1. Open a SheetLink-powered Google Sheet with transaction data
2. Go to **Extensions** → **Apps Script**
3. Paste your `recipe.gs` code
4. Save and run your entry function
5. Verify:
   - No errors in execution log
   - Output sheets created correctly
   - Handles missing data gracefully
   - Menu appears on spreadsheet refresh

**Edge Cases to Test:**
- Empty Transactions sheet
- Missing required columns
- Transactions with null/undefined values
- Single transaction vs. thousands of transactions
- Date ranges (single day, single month, multiple years)

### 4. Submit a Pull Request

```bash
git checkout -b recipe/your-recipe-name
git add recipes/community/your-recipe-name/
git commit -m "feat: Add [Your Recipe Name] community recipe"
git push origin recipe/your-recipe-name
```

Open a PR with:
- **Title:** `feat: Add [Recipe Name] community recipe`
- **Description:**
  - What the recipe does
  - Who it's for
  - What you tested
  - Screenshots (if applicable)

### 5. Review Process

We'll review your recipe for:
- Privacy/security compliance (no external calls)
- Code quality and readability
- Error handling
- Documentation completeness
- Testing coverage

We may request changes or improvements. Once approved, your recipe will be merged and available in the marketplace!

## Recipe Ideas

Looking for inspiration? Here are some recipe ideas the community might love:

### Personal Finance
- **Tax Prep Helper** - Categorize transactions by tax deduction type
- **Savings Goals Tracker** - Track progress toward savings goals
- **Debt Payoff Planner** - Snowball/avalanche debt repayment tracking
- **Net Worth Tracker** - Track assets and liabilities over time

### Business
- **Invoice Tracker** - Match expenses to client projects for invoicing
- **Expense Report Generator** - Format transactions for expense reimbursement
- **1099 Contractor Report** - Track contractor payments for 1099 filing
- **Sales Tax Calculator** - Calculate sales tax owed by jurisdiction

### Advanced Analysis
- **Category Trends** - YoY/MoM spending trends by category
- **Seasonal Spending** - Identify seasonal spending patterns
- **Merchant Analysis** - Top merchants by spend with trends
- **Account Reconciliation** - Match transactions across accounts

## Community

- **Discussions:** [GitHub Discussions](https://github.com/sheetlink/sheetlink-recipes/discussions)
- **Issues:** [Report bugs or request features](https://github.com/sheetlink/sheetlink-recipes/issues)
- **Email:** recipes@sheetlink.app

## Code of Conduct

- Be respectful and inclusive
- Provide constructive feedback
- Focus on helping users solve real problems
- Prioritize privacy and security

## License

By contributing, you agree to license your recipe under the MIT License.

---

**Ready to contribute?** Fork the repo and start building!
