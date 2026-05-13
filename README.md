# migrador

Migrate Excel/CSV data into a SQLite table — interactively via a TUI, or in one command.

## Usage

### TUI

Launch without arguments to open the interactive TUI:

```
migrador
```

Use the TUI to configure your migration and optionally save it as a JSON config file for later reuse.

### One-liner

Run a saved migration config non-interactively:

```
migrador --input data.xlsx --database mydb.db --config migration.json --mode append
```

**Options**

| Option | Description |
|---|---|
| `--input` | Path to the Excel or CSV file |
| `--database` | Path to the SQLite database |
| `--config` | Path to the migration JSON (created via the TUI) |
| `--mode` | `append`, `replace`, or `upsert` |
| `--join-col` | Column to upsert on (required when `--mode=upsert`) |

### As a library

```python
from migrador import ExcelMigrator
import sqlite3

migrator = ExcelMigrator("data.xlsx", header_row=1, start_col="A")
print(migrator.columns())
conn = sqlite3.connect("mydb.db")
rows = migrator.to_sqlite({"A": "tag", "B": "description"}, conn, "my_table")
```
