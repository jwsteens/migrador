# excel_to_sql

Standalone Python package for interactively mapping Excel/CSV columns
to a SQLite table, with a CLI and Textual TUI.

## Usage

### As a library
    from excel_to_sql import ExcelMapper
    import sqlite3

    mapper = ExcelMapper("data.xlsx", header_row=1, start_col="A")
    print(mapper.columns())
    conn = sqlite3.connect("mydb.db")
    rows = mapper.to_sqlite({"A": "tag", "B": "description"}, conn, "my_table")

### CLI
    python -m excel_to_sql.cli map-excel

### TUI
    python -m excel_to_sql.cli map-excel --tui

## Installation
    pip install -r requirements.txt
