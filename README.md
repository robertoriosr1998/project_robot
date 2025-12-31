# Project Robot

A Python utility for creating Excel workbooks with predefined sheet structures for OPC data management.

## Features

- Creates an Excel workbook with multiple sheets:
  - **Parameters** - Configuration settings (OPC folder path)
  - **OPC** - 50-column data sheet
  - **TIPS** - Fund house information and contact details
  - **CN Database** - Contract notes database with transaction details

## Requirements

- Python 3.x
- openpyxl

## Installation

```bash
pip install openpyxl
```

## Usage

```bash
python add_macro.py
```

This will generate an `OPC_TEST.xlsx` file with the predefined sheet structure.

## Files

- `add_macro.py` - Main script to generate the Excel workbook
- `OPC_TEST.xlsm` - Example Excel file with macros
- `ProcessEmailAttachments.bas` - VBA macro for processing email attachments

## VBA Macro: ProcessEmailAttachments

### What it does

1. **Gets the selected row** - Reads the row of the currently selected cell (green frame)
2. **Extracts search value** - Gets data from column E (5th column) of that row
3. **Searches TIPS sheet** - Looks for the value in column A of the TIPS worksheet
4. **Gets MY TIPS value** - Retrieves contents of column Q (MY TIPS) from the matching row
5. **Searches emails** - Finds emails from the address specified in Parameters!B4 containing the search value
6. **Processes attachments** - Downloads and attempts to open each attachment
7. **Handles passwords** - For PDFs, tries passwords from TIPS columns R, S, T in sequence
8. **Logs to CN Database** - Adds successfully opened files to CN Database with auto-incrementing ID

### Installation

1. Open `OPC_TEST.xlsm` in Excel
2. Press `Alt+F11` to open the VBA Editor
3. Go to **File > Import File** and select `ProcessEmailAttachments.bas`
4. Add required references via **Tools > References**:
   - Microsoft Outlook XX.0 Object Library
   - Adobe Acrobat XX.0 Type Library (optional, for better PDF handling)
5. Close VBA Editor and save the workbook

### Usage

1. Ensure Parameters!B4 contains the email address to search
2. Ensure TIPS sheet has data with passwords in columns R, S, T
3. Select a cell in any row (the row you want to process)
4. Run the macro: **Developer > Macros > ProcessEmailAttachments**

### Column References

| Sheet | Column | Purpose |
|-------|--------|---------|
| Active Sheet | E (5) | Search value |
| TIPS | A (1) | Search target |
| TIPS | Q (17) | MY TIPS - email subject filter |
| TIPS | R (18) | Password 1 |
| TIPS | S (19) | Password 2 |
| TIPS | T (20) | Password 3 |
| Parameters | B4 | Email address to search |
| CN Database | A (1) | Auto-increment ID |
| CN Database | B (2) | File path |

