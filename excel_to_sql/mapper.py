import sqlite3

import pandas as pd

from .helpers import build_col_map, col_letter_to_index, col_letters


class ExcelMapper:

    def __init__(
        self,
        filepath: str,
        header_row: int = 1,
        start_col: str = "A",
        sheet_name: str | int = 0,
    ):
        """
        Load an Excel (.xlsx, .xls) or CSV file into a DataFrame.

        Args:
            filepath:   Path to the source file.
            header_row: 1-indexed row number of the header (0 = no header).
            start_col:  Excel-style column letter where the table begins.
            sheet_name: Sheet name or 0-based index (Excel only; ignored for CSV).
        """
        self.filepath = filepath
        self.start_col_index = col_letter_to_index(start_col.upper())
        pandas_header = None if header_row == 0 else header_row - 1
        ext = filepath.rsplit(".", 1)[-1].lower()

        if ext in ("xlsx", "xls"):
            self.df = pd.read_excel(
                filepath, header=pandas_header, sheet_name=sheet_name
            ).iloc[:, self.start_col_index:]
        elif ext == "csv":
            df = pd.read_csv(filepath, header=pandas_header)
            self.df = df.iloc[:, self.start_col_index:]
        else:
            raise ValueError(
                f"Unsupported file type: .{ext} (expected .xlsx, .xls, or .csv)"
            )

    def columns(self) -> list[tuple[str, str]]:
        """
        Return (letter, column_name) pairs for all columns in the loaded DataFrame.
        Letters always start from A regardless of start_col offset.
        """
        letters = col_letters(len(self.df.columns))
        return list(zip(letters, self.df.columns))

    def map(self, mapping: dict[str, str]) -> pd.DataFrame:
        """
        Apply a column mapping to the DataFrame.

        Args:
            mapping: {excel_col_letter_or_name: target_sql_column_name}

        Returns:
            New DataFrame with only the mapped columns, renamed to target names.

        Raises:
            ValueError if any source column letter or name is not found.
        """
        col_map = build_col_map(self.df)
        selected = []
        rename = {}
        for source, target in mapping.items():
            key = col_map.get(source.upper()) or col_map.get(source)
            if key is None or key not in self.df.columns:
                raise ValueError(
                    f"Column '{source}' not found. "
                    f"Available: {list(self.df.columns)}"
                )
            selected.append(key)
            rename[key] = target
        return self.df[selected].rename(columns=rename)

    def preview(self, mapping: dict[str, str], n: int = 5) -> list[dict]:
        """Return the first n rows of the mapped DataFrame as a list of dicts."""
        return self.map(mapping).head(n).to_dict(orient="records")

    def to_sqlite(
        self,
        mapping: dict[str, str],
        conn: sqlite3.Connection,
        table: str,
        if_exists: str = "append",
    ) -> int:
        """
        Write the mapped DataFrame to a SQLite table.

        Args:
            mapping:   Column mapping dict (same format as map()).
            conn:      Open sqlite3 connection.
            table:     Target table name.
            if_exists: "append", "replace", or "fail".

        Returns:
            Number of rows written.
        """
        df = self.map(mapping)
        df.to_sql(table, conn, if_exists=if_exists, index=False)
        return len(df)
