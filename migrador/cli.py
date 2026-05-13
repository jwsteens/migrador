import sqlite3

import click

from .helpers import load_migration
from .migrator import ExcelMigrator


@click.command()
@click.option("--input", "input_file", default=None, help="Path to the Excel or CSV file")
@click.option("--database", default=None, help="Path to the SQLite database")
@click.option("--config", "migration_file", default=None, help="Path to the migration JSON")
@click.option(
    "--mode",
    type=click.Choice(["append", "replace", "upsert"], case_sensitive=False),
    default=None,
    help="Import mode: append, replace, or upsert",
)
@click.option("--join-col", default=None, help="Column to upsert on (required when --mode=upsert)")
def cli(input_file, database, migration_file, mode, join_col):
    if not any([input_file, database, migration_file, mode]):
        from .tui import ExcelMigratorApp
        ExcelMigratorApp().run()
        return

    missing = [name for name, val in [("--input", input_file), ("--database", database), ("--config", migration_file), ("--mode", mode)] if not val]
    if missing:
        raise click.UsageError(f"Missing options when running without TUI: {', '.join(missing)}")

    if mode == "upsert" and not join_col:
        raise click.UsageError("--join-col is required when --mode=upsert")

    migration = load_migration(migration_file)

    header_row = migration.get("header_row", 1)
    start_col = migration.get("start_col", "A")
    sheet_name = migration.get("sheet_name", 0)
    target_table = migration.get("target_table")
    if not target_table:
        raise click.UsageError("Migration JSON is missing 'target_table'")

    migrator = ExcelMigrator(input_file, header_row=header_row, start_col=start_col, sheet_name=sheet_name)

    columns = migration.get("columns", [])
    if isinstance(columns, list):
        col_map = {entry["source"]: entry["name"] for entry in columns if entry.get("source") != "skip" and entry.get("name")}
    else:
        col_map = {src: tgt for src, tgt in columns.items()}

    df = migrator.migrate(col_map)

    conn = sqlite3.connect(database)
    try:
        if mode == "append":
            df.to_sql(target_table, conn, if_exists="append", index=False)
        elif mode == "replace":
            df.to_sql(target_table, conn, if_exists="replace", index=False)
        elif mode == "upsert":
            cols = list(df.columns)
            placeholders = ", ".join(["?" for _ in cols])
            update_set = ", ".join(f"{c}=excluded.{c}" for c in cols if c != join_col)
            sql = (
                f"INSERT INTO {target_table} ({', '.join(cols)}) "
                f"VALUES ({placeholders}) "
                f"ON CONFLICT({join_col}) DO UPDATE SET {update_set}"
            )
            conn.cursor().executemany(sql, df.values.tolist())
            conn.commit()
        click.echo(f"Imported {len(df)} rows into '{target_table}'.")
    finally:
        conn.close()


if __name__ == "__main__":
    cli()
