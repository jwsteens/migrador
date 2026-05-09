import glob
import readline
import sqlite3

import click

from .helpers import (
    build_col_map,
    load_mapping,
    preview_table,
    print_columns,
    prompt_column,
    save_mapping,
    val,
)
from .mapper import ExcelMapper


def prompt_path(message: str, default: str = None) -> str:
    def _completer(text, state):
        matches = glob.glob(text + "*")
        return matches[state] if state < len(matches) else None

    readline.set_completer(_completer)
    readline.parse_and_bind("tab: complete")
    try:
        return click.prompt(message, default=default)
    finally:
        readline.set_completer(None)


@click.group()
def cli():
    pass


@cli.command("map-excel")
@click.option("--mapping-file", default=None, help="Path to a saved mapping JSON")
@click.option(
    "--save-mapping",
    "save_mapping_path",
    default=None,
    help="Save completed mapping to this path",
)
@click.option(
    "--tui",
    is_flag=True,
    default=False,
    help="Launch the full Textual TUI instead",
)
def map_excel(mapping_file, save_mapping_path, tui):
    if tui:
        from .tui import ExcelMapperApp
        ExcelMapperApp().run()
        return

    # Step 1: file path
    filepath = prompt_path("Excel/CSV file path")

    # Step 2: header row and start column
    saved = {}
    if mapping_file:
        saved = load_mapping(mapping_file)
        header_row = saved.get("header_row", 1)
        start_col = saved.get("start_col", "A")
    else:
        header_row = click.prompt("Header row", default=1, type=int)
        while True:
            start_col = click.prompt("Start column", default="A").strip()
            if start_col.isalpha():
                break
            click.echo("Start column must be a non-empty alphabetic string (e.g. 'A', 'B').")

    # Step 3: load file and show columns
    mapper = ExcelMapper(filepath, header_row=header_row, start_col=start_col)
    print_columns(mapper.df)

    # Step 4-5: DB path and table
    db_path = prompt_path("SQLite DB path")
    table = click.prompt("Target table name", default=saved.get("target_table", ""))

    # Step 6: introspect target table columns
    conn = sqlite3.connect(db_path)
    cursor = conn.execute(f"PRAGMA table_info({table})")
    pragma_rows = cursor.fetchall()
    # PRAGMA table_info columns: cid, name, type, notnull, dflt_value, pk
    target_cols = [
        row[1]
        for row in pragma_rows
        if not (row[5] == 1 and row[2].upper() == "INTEGER")
    ]
    if not target_cols:
        click.echo(
            f"Error: table '{table}' does not exist or has no eligible columns."
        )
        conn.close()
        return

    # Step 7: map each target column
    col_map = build_col_map(mapper.df)
    saved_columns = saved.get("columns", {})
    saved_renames = saved.get("renames", {})

    columns = {}   # {excel_source: sql_target}
    renames = {}   # {sql_target: final_name}

    for col_name in target_cols:
        click.echo(f"\nColumn: {col_name}")
        default_source = saved_columns.get(col_name, col_name)
        source_key = prompt_column(
            f"  Map from Excel column",
            default_name=default_source,
            col_map=col_map,
            df=mapper.df,
            required=True,
        )
        columns[str(source_key)] = col_name

        default_rename = saved_renames.get(col_name, col_name)
        rename_to = click.prompt(
            f"  Rename to? (Enter to keep '{col_name}')",
            default=default_rename,
        ).strip()
        if rename_to and rename_to != col_name:
            renames[col_name] = rename_to

    # Step 8: save mapping if requested
    mapping_dict = {
        "header_row": header_row,
        "start_col": start_col,
        "target_table": table,
        "columns": columns,
        "renames": renames,
    }
    if save_mapping_path:
        save_mapping(save_mapping_path, mapping_dict)
        click.echo(f"Mapping saved to {save_mapping_path}")

    # Step 9: build final mapping applying renames
    final_mapping = {
        source: renames.get(target, target)
        for source, target in columns.items()
    }

    # Step 10: build mapped DataFrame
    df = mapper.map(final_mapping)

    # Step 11: launch preview TUI
    from .tui import ExcelPreviewApp
    app = ExcelPreviewApp(df, conn, table)
    app.run()
    if app.confirmed:
        click.echo(f"Imported {len(df)} rows into '{table}'.")
    else:
        click.echo("Import cancelled.")

    conn.close()


if __name__ == "__main__":
    cli()
