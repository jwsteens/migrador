import json

import click
import pandas as pd
from rich.console import Console
from rich.table import Table


def col_letters(n: int) -> list[str]:
    """Generate Excel-style column labels for n columns: A, B, …, Z, AA, AB, …"""
    labels = []
    for i in range(n):
        label = ""
        col = i
        while True:
            label = chr(ord("A") + col % 26) + label
            col = col // 26 - 1
            if col < 0:
                break
        labels.append(label)
    return labels


def col_letter_to_index(letter: str) -> int:
    """Convert an Excel-style column letter to a 0-based integer index.

    "A" → 0, "B" → 1, "Z" → 25, "AA" → 26, etc.
    """
    if not letter or not letter.isalpha():
        raise ValueError(
            f"Invalid column letter: {letter!r}. Expected a non-empty alphabetic string like 'A', 'Z', or 'AA'."
        )
    index = 0
    for char in letter.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def build_col_map(df: pd.DataFrame) -> dict[str, object]:
    """Return a dict mapping both letter (e.g. 'A') and column name to the DataFrame column key."""
    letters = col_letters(len(df.columns))
    col_map = {}
    for letter, col in zip(letters, df.columns):
        col_map[letter] = col
        col_map[str(col)] = col
    return col_map


def print_columns(df: pd.DataFrame) -> None:
    """Print available columns with their letter labels using click.echo."""
    letters = col_letters(len(df.columns))
    for letter, col in zip(letters, df.columns):
        click.echo(f"  {letter:<4} {col}")


def prompt_column(
    label: str,
    default_name: str,
    col_map: dict,
    df: pd.DataFrame,
    required: bool = True,
) -> object | None:
    """Interactive prompt to select a column by letter or name.

    Returns the DataFrame column key, or None if skipped (required=False).
    """
    has_default = default_name and default_name in col_map

    while True:
        if required:
            value = click.prompt(
                label,
                default=default_name if has_default else None,
                show_default=has_default,
            ).strip()
        else:
            raw = click.prompt(
                f"{label} (letter or name, Enter to skip)",
                default="",
                show_default=False,
            ).strip()
            if not raw:
                return None
            value = raw

        if not value:
            click.echo("This column is required. Please enter a letter or column name.")
            print_columns(df)
            continue

        key = col_map.get(value.upper()) or col_map.get(value)
        if key is None or key not in df.columns:
            click.echo(f"Column '{value}' not found. Use a letter (e.g. 'A') or exact column name.")
            print_columns(df)
            continue

        return key


def val(row: pd.Series, col: object | None) -> str | None:
    """Extract a stripped string value from a row, None for missing/NaN/empty."""
    if col is None or col not in row.index:
        return None
    v = row[col]
    if pd.isna(v):
        return None
    s = str(v).strip()
    return s if s else None


def preview_table(records: list[dict]) -> None:
    """Print a rich bordered table for a list of row dicts."""
    if not records:
        click.echo("(no rows to preview)")
        return
    console = Console()
    table = Table(show_header=True, show_lines=True, overflow="fold")
    columns = list(records[0].keys())
    for col in columns:
        table.add_column(str(col))
    for record in records:
        table.add_row(*[str(record.get(col, "")) for col in columns])
    console.print(table)


def load_migration(path: str) -> dict:
    """Read a migration JSON file. Accepts both TUI format (list columns) and CLI format (dict columns)."""
    with open(path, "r", encoding="utf-8") as f:
        migration = json.load(f)
    new_keys = {"sheet_name", "header_row", "start_col", "target_table", "columns"}
    old_keys = {"header_row", "start_col", "target_table", "columns", "renames"}
    if new_keys.issubset(migration.keys()):
        if not isinstance(migration["columns"], list):
            raise ValueError(
                f"Migration file '{path}': 'columns' must be a list of {{name, source, type}} dicts"
            )
    elif old_keys.issubset(migration.keys()):
        pass  # CLI format, accepted as-is
    else:
        missing = new_keys - migration.keys()
        raise ValueError(
            f"Migration file '{path}' is missing required keys: {sorted(missing)}"
        )
    return migration


def save_migration(path: str, migration: dict) -> None:
    """Write a migration dict to a JSON file with indent=2."""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(migration, f, indent=2)
