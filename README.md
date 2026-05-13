# migrador

Migrate Excel/CSV data into a SQLite table — interactively via a TUI, or in one command.

## Usage

### TUI

Launch without arguments to open the interactive TUI:

```
migrador
```
<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/82547bc4-b31e-4e35-aac7-2f2ccb97d24a" />
<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/8e7671aa-a107-4a54-8b4a-0ac2b36f3381" />
<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/e6154733-3029-44e1-bca1-54f93137f735" />
<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/e4b18f03-cf04-49ac-a06a-e0a04b3ff47e" />

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
