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

