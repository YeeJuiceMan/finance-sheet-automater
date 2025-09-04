# Finance Sheet Automater

Automate and streamline your personal finance tracking directly within Google Sheets.

## Overview

**Finance Sheet Automater** is a Google Apps Script-based tool for automating the organization, entry, and modification of financial data in Google Sheets. It is designed to make tracking expenses, reimbursements, income, and notes in complex finance sheets easy and error-free.

## Features

- **Automated Data Entry:** Easily add or modify financial entries (expenses, income, reimbursements) across multiple sheets.
- **Expense Categorization:** Distinguishes between needs, wants, and reimbursements, supporting custom categories.
- **Notes Management:** Attach and update notes to specific entries, with advanced logic for splitting, formatting, and transferring notes.
- **Reimbursement Tracking:** Detects unreimbursed items, updates reimbursement status, and assists with monthly reconciliation.
- **Sheet Configuration:** Works with multiple sheets (e.g., College Savings, Console) and dynamically finds rows/columns for each operation.
- **Dropdown Validation:** Ensures all entry fields use validated dropdown lists to prevent errors.
- **Customizable:** Easily extend or adapt to new sheets, categories, or financial workflows.

## How It Works

- Utilizes Google Apps Script (`src/subbutton&notemod.js`) bound to your spreadsheet.
- Provides custom functions and triggers to interact with sheet buttons, edit events, and automation scripts.
- Main workflow involves finding the correct cell/range, validating data, and recording the necessary information with notes and categorization.

## Example Use Cases

- Track monthly expenses and income with detailed notes.
- Automate reimbursement verification and update relevant sheet fields.
- Quickly add new financial transactions via sheet buttons or script triggers.
- Split transaction notes and costs into subcategories or specific tracking columns.

## Getting Started

1. **Setup:** Bind the script to your Google Spreadsheet containing finance tracking sheets.
2. **Configuration:** Update sheet names and categories in the script to match your own setup.
3. **Usage:** Use provided functions or sheet buttons to add, modify, and track financial data.

## License

This project is licensed under the MIT License. See [LICENSE.txt](LICENSE.txt) for details.

## Author

[YeeJuiceMan](https://github.com/YeeJuiceMan)
