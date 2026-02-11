# Excel to Access Importer

## About This Tool

While I generally recommend modern database solutions like PostgreSQL, SQLite, or cloud-based options for data analysis, I recognize that in academic and research environments, there are situations where collaborators need or prefer to work with Microsoft Access. This may be due to institutional requirements, existing workflows, legacy systems, or simply the tools they are most comfortable with.

This script provides a straightforward solution for researchers and data analysts who need to import Excel data into Access databases. It automates the tedious process of creating tables, handling column name restrictions, and importing data—allowing scientists to focus on their analysis rather than data wrangling.

## Features

- Imports all worksheets from an Excel file into separate Access tables
- Automatically sanitizes column names to meet Access requirements (max 30 characters, no special characters)
- Preserves original column names in a mapping table (__column_map)
- Handles data type conversion (INTEGER, FLOAT, BOOLEAN, DATETIME, TEXT)
- Batch inserts for improved stability
- Command-line interface for easy automation

## Requirements

- Python 3.7+
- pandas
- pyodbc
- Microsoft Access (to create the initial database file)
- Microsoft Access ODBC Driver installed on your system

## Installation

Install required Python packages:

```bash
pip install pandas pyodbc
```

or with conda:

```bash
conda install pandas pyodbc
```

## Usage

### Step 1: Create an empty Access database

Before running the script, you must create an empty Access database:

1. Open Microsoft Access
2. Click "Blank database"
3. Choose a location and filename (e.g., `database.accdb`)
4. Click "Create"
5. Close Access

### Step 2: Run the import script

```bash
python DataImportToAccess.py <input.xlsx> <output.accdb>
```

Example:

```bash
python DataImportToAccess.py survey_data.xlsx survey_database.accdb
```

The script will:
1. Read all worksheets from the Excel file
2. Create one Access table per worksheet
3. Sanitize column names (replace special characters, limit to 30 chars, ensure uniqueness)
4. Create a mapping table (__column_map) to preserve original column names
5. Import all data in batches of 300 rows

## Configuration

You can modify these constants at the top of the script:

- `MAX_COLNAME_LEN`: Maximum column name length (default: 30)
- `NAMING_MODE`: "short" for sanitized names or "letters" for A, B, C style
- `INSERT_BATCH_SIZE`: Number of rows per batch insert (default: 300)
- `USE_FAST_EXECUTEMANY`: Enable fast executemany mode (default: False)

## Column Mapping Table

The script creates a special table `__column_map` with three columns:
- `table_name`: Name of the Access table
- `original_name`: Original column header from Excel
- `short_name`: Sanitized column name used in Access

This allows you to trace back shortened or modified column names to their original Excel headers.

## Troubleshooting

### "Access database not found" error
Make sure you created an empty .accdb file in Microsoft Access before running the script.

### "Database is already open" error
Close Microsoft Access before running the script.

### "ODBC driver not found" error
Install the Microsoft Access Database Engine:
- Download from: https://www.microsoft.com/en-us/download/details.aspx?id=54920
- Install the version matching your Python (32-bit or 64-bit)

## License

MIT License

Copyright (c) 2026

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.