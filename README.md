# Split Sheets

A flexible, dynamic bill-splitting script for Google Sheets. Perfect for roommates, couples, or groups who need to track shared expenses and settle up fairly.

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)

## Features

- **Dynamic People Support** - Works with any number of people (2+)
- **Three Split Types**:
  - **Equally** - Checkbox-based equal splits among selected people
  - **Variably** - Percentage-based splits with auto-calculation
  - **Fixed** - Dollar amount splits with auto-calculation
- **Auto-Calculate** - When you enter one person's share, the remainder auto-fills
- **Real-Time Updates** - Calculations update automatically on every edit
- **Summary Row** - Shows who owes what at a glance
- **Per-Row Breakdown** - See exactly how each expense is split

## Setup

### 1. Create Your Google Sheet

Set up your spreadsheet with these columns:

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| Description | Who Paid | Amount | How to split | Person 1 | Person 2 | Breakdown |

> **Note:** Add as many person columns as needed (E, F, G, etc.). The "Breakdown" column is auto-created after the last person.

### 2. Add the Script

1. Open your Google Sheet
2. Go to **Extensions** → **Apps Script**
3. Delete any existing code
4. Copy and paste the contents of `Split-Sheets.js`
5. Click **Save**
6. Refresh your spreadsheet

### 3. Set Up Data Validation (Optional but Recommended)

For the **"Who Paid"** column (B):
1. Select column B (excluding header)
2. Go to **Data** → **Data validation**
3. Set criteria to **List of items** with your names (e.g., `Alice,Bob`)

For the **"How to split"** column (D):
1. Select column D (excluding header)
2. Go to **Data** → **Data validation**
3. Set criteria to **List of items**: `Equally,Variably,Fixed`

## How to Use

### Equal Splits
1. Select **"Equally"** in the "How to split" column
2. Checkboxes appear in person columns
3. Check the people who share the expense
4. The amount is split evenly among checked people

### Variable (Percentage) Splits
1. Select **"Variably"** in the "How to split" column
2. Percentage inputs appear in person columns
3. Enter one person's percentage (e.g., 60%)
4. The other person's percentage auto-calculates (40%)

### Fixed Amount Splits
1. Select **"Fixed"** in the "How to split" column
2. Dollar inputs appear in person columns
3. Enter one person's fixed amount
4. The remainder auto-calculates for others

## Example

| Description | Who Paid | Amount | How to split | Alice | Bob | Breakdown |
|-------------|----------|--------|--------------|-------|-----|-----------|
| Rent | Alice | $2,000 | Equally | ☑️ | ☑️ | Bob Pays: $1,000.00 |
| Internet | Bob | $80 | Equally | ☑️ | ☑️ | Alice Pays: $40.00 |
| Groceries | Alice | $150 | Variably | 60% | 40% | Bob Pays: $60.00 |
| Dinner | Bob | $75 | Fixed | $25 | $50 | Alice Pays: $25.00 |
| **Due: 1st** | **Summary** | **$2,305** | | | | **Bob owes $1,085.00** |

## Custom Menu

After installation, a **"Split Sheets"** menu appears in your spreadsheet:

- **Recalculate** - Manually trigger a full recalculation

## File Structure

```
Split-Sheets/
├── Split-Sheets.js    # Main Google Apps Script
└── README.md          # This file
```

## Configuration

You can customize these constants at the top of the script:

```javascript
var DESCRIPTION_COL = 1;    // Column A
var WHO_PAID_COL = 2;       // Column B
var AMOUNT_COL = 3;         // Column C
var SPLIT_TYPE_COL = 4;     // Column D
var FIRST_SPLITTER_COL = 5; // Column E (first person)
```

## Contributing

Contributions are welcome! Feel free to:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Tips

- The script runs on every cell edit, so calculations are always current
- Use the "Recalculate" menu option if something seems off
- The summary row automatically positions itself after the last expense row
- Works best with Chrome and the Google Sheets desktop app

## Known Limitations

- Requires Google Sheets (not compatible with Excel)
- Checkbox state may not persist if you change split types frequently
- Very large spreadsheets (1000+ rows) may experience slight delays
