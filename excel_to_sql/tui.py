import asyncio
import sqlite3

import pandas as pd
from textual.app import App, ComposeResult
from textual.screen import Screen, ModalScreen
from textual.widgets import (
    Button, Checkbox, DataTable, DirectoryTree, Input,
    Label, RadioButton, RadioSet, Select, Static,
)
from textual.containers import Horizontal, Vertical, VerticalScroll
from textual import work
from rich.text import Text

from .mapper import ExcelMapper
from .helpers import build_col_map, col_letters, load_mapping, save_mapping


# ─────────────────────────────────────────────────────────────
# Shared helpers
# ─────────────────────────────────────────────────────────────

def build_display_rows(
    df: pd.DataFrame,
) -> tuple[list[dict], list[bool]]:
    cols = list(df.columns)
    sep = {col: "…" for col in cols}

    if len(df) < 50:
        rows = df.to_dict(orient="records")
        return rows, [False] * len(rows)

    head = df.iloc[:5].to_dict(orient="records")
    middle_pool = df.iloc[5:-5]
    n_sample = min(40, len(middle_pool))
    middle = (
        middle_pool.sample(n=n_sample, random_state=42).to_dict(orient="records")
        if n_sample > 0
        else []
    )
    tail = df.iloc[-5:].to_dict(orient="records")

    display_rows = head + [sep] + middle + [sep] + tail
    is_separator = (
        [False] * len(head)
        + [True]
        + [False] * len(middle)
        + [True]
        + [False] * len(tail)
    )
    return display_rows, is_separator


def _collapse_separators(
    rows: list[dict], is_sep: list[bool]
) -> tuple[list[dict], list[bool]]:
    out_rows: list[dict] = []
    out_sep: list[bool] = []
    pending_sep: dict | None = None

    for row, sep in zip(rows, is_sep):
        if sep:
            pending_sep = row
        else:
            if pending_sep is not None and out_rows:
                out_rows.append(pending_sep)
                out_sep.append(True)
            pending_sep = None
            out_rows.append(row)
            out_sep.append(False)

    return out_rows, out_sep


# ─────────────────────────────────────────────────────────────
# FilePickerModal
# ─────────────────────────────────────────────────────────────

class FilePickerModal(ModalScreen):
    def __init__(self, title: str, start_path: str = ".") -> None:
        super().__init__()
        self._title = title
        self._start_path = start_path

    def compose(self) -> ComposeResult:
        with Vertical(id="filepicker-container"):
            yield Label(self._title)
            yield DirectoryTree(self._start_path, id="filepicker-tree")
            yield Input(placeholder="Selected path", id="filepicker-input")
            with Horizontal(id="filepicker-buttons"):
                yield Button("Cancel", id="btn-cancel")
                yield Button("OK", variant="primary", id="btn-ok")

    def on_directory_tree_file_selected(
        self, event: DirectoryTree.FileSelected
    ) -> None:
        self.query_one("#filepicker-input", Input).value = str(event.path)

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "btn-ok":
            path = self.query_one("#filepicker-input", Input).value.strip()
            self.dismiss(path or None)
        else:
            self.dismiss(None)

    def on_key(self, event) -> None:
        if event.key == "escape":
            self.dismiss(None)


# ─────────────────────────────────────────────────────────────
# PathInputModal
# ─────────────────────────────────────────────────────────────

class PathInputModal(ModalScreen):
    def __init__(self, title: str, placeholder: str = "") -> None:
        super().__init__()
        self._title = title
        self._placeholder = placeholder

    def compose(self) -> ComposeResult:
        with Vertical(id="modal-container"):
            yield Label(self._title)
            yield Input(placeholder=self._placeholder, id="modal-input")
            with Horizontal(id="modal-buttons"):
                yield Button("Cancel", id="btn-cancel")
                yield Button("OK", variant="primary", id="btn-ok")

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "btn-ok":
            path = self.query_one("#modal-input", Input).value.strip()
            self.dismiss(path or None)
        else:
            self.dismiss(None)

    def on_key(self, event) -> None:
        if event.key == "escape":
            self.dismiss(None)


# ─────────────────────────────────────────────────────────────
# Screen 1: FileScreen
# ─────────────────────────────────────────────────────────────

class FileScreen(Screen):
    def compose(self) -> ComposeResult:
        with Vertical(id="file-screen-content"):
            yield Label("Excel / CSV file")
            with Horizontal(classes="file-picker-row"):
                yield Input(placeholder="path/to/file.xlsx", id="excel-path")
                yield Button("Browse", id="btn-browse-excel")
            yield Label("SQLite database file")
            with Horizontal(classes="file-picker-row"):
                yield Input(placeholder="path/to/database.db", id="db-path")
                yield Button("Browse", id="btn-browse-db")
            yield Label("(will be created if not found)", classes="hint-label")
            with Horizontal(id="file-buttons"):
                yield Button("Exit", id="btn-exit", variant="error")
                yield Button(
                    "Load mapping from JSON",
                    id="btn-load-mapping",
                    variant="primary",
                )
                yield Button(
                    "Create mapping",
                    id="btn-create-mapping",
                    variant="primary",
                )

    def on_button_pressed(self, event: Button.Pressed) -> None:
        bid = event.button.id
        if bid == "btn-browse-excel":
            self._browse("excel")
        elif bid == "btn-browse-db":
            self._browse("db")
        elif bid == "btn-load-mapping":
            self._load_mapping_flow()
        elif bid == "btn-create-mapping":
            self._create_mapping_flow()
        elif bid == "btn-exit":
            self.app.exit()

    def _browse(self, target: str) -> None:
        title = "Select Excel/CSV file" if target == "excel" else "Select SQLite database"
        input_id = "excel-path" if target == "excel" else "db-path"

        def _on_path(path: str | None) -> None:
            if path:
                self.query_one(f"#{input_id}", Input).value = path

        self.app.push_screen(FilePickerModal(title), _on_path)

    def _validate_paths(self) -> tuple[str, str] | None:
        excel_path = self.query_one("#excel-path", Input).value.strip()
        db_path = self.query_one("#db-path", Input).value.strip()
        if not excel_path:
            self.app.notify("Excel/CSV file path is required.", severity="warning")
            return None
        if not db_path:
            self.app.notify("SQLite DB path is required.", severity="warning")
            return None
        return excel_path, db_path

    def _load_mapping_flow(self) -> None:
        paths = self._validate_paths()
        if not paths:
            return
        excel_path, db_path = paths
        self.app.excel_path = excel_path
        self.app.db_path = db_path

        def _on_json_path(json_path: str | None) -> None:
            if not json_path:
                return
            try:
                m = load_mapping(json_path)
                self.app.loaded_mapping = m
                self.app.push_screen(SheetTableScreen())
            except Exception as e:
                self.app.notify(str(e), severity="error")

        self.app.push_screen(FilePickerModal("Load mapping JSON"), _on_json_path)

    def _create_mapping_flow(self) -> None:
        paths = self._validate_paths()
        if not paths:
            return
        excel_path, db_path = paths
        self.app.excel_path = excel_path
        self.app.db_path = db_path
        self.app.loaded_mapping = None
        self.app.push_screen(SheetTableScreen())


# ─────────────────────────────────────────────────────────────
# Screen 2: SheetTableScreen
# ─────────────────────────────────────────────────────────────

class SheetTableScreen(Screen):
    def __init__(self) -> None:
        super().__init__()
        self._sheet_names: list[str] = []
        self._db_tables: list[str] = []

    def compose(self) -> ComposeResult:
        with VerticalScroll(id="sheet-screen-content"):
            yield Label("Select sheet")
            yield RadioSet(id="sheet-radioset")
            with Horizontal(classes="form-row"):
                yield Label("Header row")
                yield Input(value="1", id="header-row")
            with Horizontal(classes="form-row"):
                yield Label("Start column")
                yield Input(value="A", id="start-col")
            yield Label("Target table")
            yield RadioSet(id="table-radioset")
            yield Input(placeholder="Or type a new table name", id="new-table-name")
            with Horizontal(id="sheet-buttons"):
                yield Button("Back", id="btn-back")
                yield Button("Continue", variant="primary", id="btn-continue")

    def on_mount(self) -> None:
        self._load_sheets_and_tables()

    @work
    async def _load_sheets_and_tables(self) -> None:
        excel_path = self.app.excel_path
        db_path = self.app.db_path

        ext = excel_path.rsplit(".", 1)[-1].lower()
        if ext == "csv":
            sheet_names = ["CSV"]
        else:
            try:
                sheet_names = await asyncio.to_thread(
                    lambda: pd.ExcelFile(excel_path).sheet_names
                )
            except Exception as e:
                self.app.notify(f"Could not read Excel file: {e}", severity="error")
                sheet_names = []

        def _get_tables() -> list[str]:
            conn = sqlite3.connect(db_path, check_same_thread=False)
            rows = conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
            conn.close()
            return [r[0] for r in rows]

        try:
            db_tables = await asyncio.to_thread(_get_tables)
        except Exception as e:
            self.app.notify(f"Could not read DB: {e}", severity="warning")
            db_tables = []

        self._sheet_names = sheet_names
        self._db_tables = db_tables
        self._populate_radiosets()

        if self.app.loaded_mapping:
            self._prefill_from_mapping(self.app.loaded_mapping)

    def _populate_radiosets(self) -> None:
        sheet_rs = self.query_one("#sheet-radioset", RadioSet)
        for name in self._sheet_names:
            sheet_rs.mount(RadioButton(str(name)))
        if len(self._sheet_names) == 1:
            buttons = list(sheet_rs.query(RadioButton))
            if buttons:
                buttons[0].value = True

        table_rs = self.query_one("#table-radioset", RadioSet)
        for name in self._db_tables:
            table_rs.mount(RadioButton(str(name)))

    def _prefill_from_mapping(self, m: dict) -> None:
        sheet_name = str(m.get("sheet_name", ""))
        if sheet_name in [str(s) for s in self._sheet_names]:
            idx = [str(s) for s in self._sheet_names].index(sheet_name)
            buttons = list(self.query_one("#sheet-radioset", RadioSet).query(RadioButton))
            if idx < len(buttons):
                buttons[idx].value = True

        self.query_one("#header-row", Input).value = str(m.get("header_row", 1))
        self.query_one("#start-col", Input).value = str(m.get("start_col", "A"))

        target_table = str(m.get("target_table", ""))
        if target_table in self._db_tables:
            idx = self._db_tables.index(target_table)
            buttons = list(self.query_one("#table-radioset", RadioSet).query(RadioButton))
            if idx < len(buttons):
                buttons[idx].value = True
        else:
            self.query_one("#new-table-name", Input).value = target_table

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "btn-back":
            self.app.pop_screen()
        elif event.button.id == "btn-continue":
            self._go_continue()

    @work
    async def _go_continue(self) -> None:
        sheet_rs = self.query_one("#sheet-radioset", RadioSet)
        sheet_idx = sheet_rs.pressed_index
        if sheet_idx < 0 or not self._sheet_names:
            self.app.notify("Select a sheet.", severity="warning")
            return
        sheet_name = self._sheet_names[sheet_idx]

        header_row_str = self.query_one("#header-row", Input).value.strip()
        start_col = self.query_one("#start-col", Input).value.strip().upper()
        try:
            header_row = int(header_row_str) if header_row_str else 1
        except ValueError:
            self.app.notify("Header row must be an integer.", severity="warning")
            return
        if not start_col.isalpha():
            self.app.notify("Start column must be alphabetic (e.g. 'A').", severity="warning")
            return

        table_rs = self.query_one("#table-radioset", RadioSet)
        table_idx = table_rs.pressed_index
        new_table_input = self.query_one("#new-table-name", Input).value.strip()
        if 0 <= table_idx < len(self._db_tables):
            target_table = self._db_tables[table_idx]
        elif new_table_input:
            target_table = new_table_input
        else:
            self.app.notify("Select or enter a target table name.", severity="warning")
            return

        try:
            mapper = await asyncio.to_thread(
                ExcelMapper, self.app.excel_path, header_row, start_col, sheet_name
            )
        except Exception as e:
            self.app.notify(str(e), severity="error")
            return

        def _open_and_introspect() -> tuple[sqlite3.Connection, list[tuple[str, str]]]:
            if self.app.conn:
                self.app.conn.close()
                self.app.conn = None
            conn = sqlite3.connect(self.app.db_path, check_same_thread=False)
            rows = conn.execute(f"PRAGMA table_info({target_table})").fetchall()
            cols = [
                (r[1], r[2] or "TEXT")
                for r in rows
                if not (r[5] == 1 and r[2].upper() == "INTEGER")
            ]
            return conn, cols

        try:
            conn, existing_cols = await asyncio.to_thread(_open_and_introspect)
        except Exception as e:
            self.app.notify(str(e), severity="error")
            return

        self.app.mapper = mapper
        self.app.conn = conn
        self.app.sheet_name = sheet_name
        self.app.header_row = header_row
        self.app.start_col = start_col
        self.app.target_table = target_table
        self.app.existing_table_cols = existing_cols
        self.app.push_screen(MappingScreen())


# ─────────────────────────────────────────────────────────────
# Screen 3: MappingScreen — column card widget
# ─────────────────────────────────────────────────────────────

_SQL_TYPES = [("TEXT", "TEXT"), ("INTEGER", "INTEGER"), ("REAL", "REAL"), ("BLOB", "BLOB"), ("NUMERIC", "NUMERIC")]


class ColumnCard(Vertical):
    def __init__(
        self,
        card_index: int,
        col_name: str,
        source: str,
        col_type: str,
        unique: bool,
        excel_options: list[tuple[str, str]],
        card_width: int = 28,
    ) -> None:
        super().__init__(id=f"col-card-{card_index}", classes="column-card")
        self._idx = card_index
        self._col_name = col_name
        self._source = source
        self._col_type = col_type
        self._unique = unique
        self._excel_options = excel_options
        self._card_width = card_width

    def on_mount(self) -> None:
        self.styles.width = self._card_width

    def compose(self) -> ComposeResult:
        source_val = self._source if self._source != "skip" else Select.NULL
        yield Input(
            value=self._col_name,
            placeholder="column name",
            id=f"card-name-{self._idx}",
        )
        yield Select(
            options=self._excel_options,
            value=source_val,
            allow_blank=True,
            id=f"card-source-{self._idx}",
        )
        yield Select(
            options=_SQL_TYPES,
            value=self._col_type if self._col_type in dict(_SQL_TYPES) else "TEXT",
            allow_blank=False,
            id=f"card-type-{self._idx}",
        )
        yield Checkbox("Unique", value=self._unique, id=f"card-unique-{self._idx}")
        yield Button("×", id=f"card-delete-{self._idx}", classes="delete-btn")

    def on_select_changed(self, event: Select.Changed) -> None:
        if event.select.id == f"card-source-{self._idx}":
            is_skip = event.value is Select.NULL
            try:
                self.query_one(f"#card-name-{self._idx}", Input).disabled = is_skip
                self.query_one(f"#card-type-{self._idx}", Select).disabled = is_skip
                self.query_one(f"#card-unique-{self._idx}", Checkbox).disabled = is_skip
            except Exception:
                pass


# ─────────────────────────────────────────────────────────────
# Screen 3: MappingScreen
# ─────────────────────────────────────────────────────────────

class MappingScreen(Screen):
    def __init__(self) -> None:
        super().__init__()
        self._next_card_index: int = 0
        self._card_indices: list[int] = []

    def _excel_options(self) -> list[tuple[str, str]]:
        return [
            (f"{letter}  {name}", letter)
            for letter, name in self.app.mapper.columns()
        ] + [("(skip)", "skip")]

    def _card_width(self) -> int:
        opts = self._excel_options()
        max_label = max((len(label) for label, _ in opts), default=10)
        return max(max_label, 12) + 4  # padding on each side

    def _initial_cards(self) -> list[tuple[str, str, str, bool]]:
        opts_values = {val for _, val in self._excel_options()}
        if self.app.loaded_mapping and isinstance(
            self.app.loaded_mapping.get("columns"), list
        ):
            result = []
            for entry in self.app.loaded_mapping["columns"]:
                src = entry.get("source", "skip")
                if src not in opts_values:
                    src = "skip"
                result.append((
                    entry.get("name", ""),
                    src,
                    entry.get("type", "TEXT"),
                    bool(entry.get("unique", False)),
                ))
            return result
        if self.app.existing_table_cols:
            return [
                (name, "skip", sql_type or "TEXT", False)
                for name, sql_type in self.app.existing_table_cols
            ]
        return []

    def compose(self) -> ComposeResult:
        opts = self._excel_options()
        w = self._card_width()
        with Vertical(id="mapping-screen"):
            with Vertical(id="excel-panel"):
                yield Label("Excel columns  (reference)")
                yield DataTable(id="excel-col-table")
            with Vertical(id="db-panel"):
                yield Label("DB columns")
                with Horizontal(id="card-scroll"):
                    for col_name, source, col_type, unique in self._initial_cards():
                        idx = self._next_card_index
                        self._card_indices.append(idx)
                        self._next_card_index += 1
                        yield ColumnCard(idx, col_name, source, col_type, unique, opts, w)
                    yield Button("+ Add column", id="btn-add-col", classes="add-col-btn")
                with Horizontal(id="mapping-buttons"):
                    yield Button("Back", id="btn-back")
                    yield Button("Save mapping to JSON", id="btn-save-mapping")
                    yield Button("Preview", variant="primary", id="btn-preview")

    def on_mount(self) -> None:
        table = self.query_one("#excel-col-table", DataTable)
        df = self.app.mapper.df
        col_pairs = self.app.mapper.columns()   # [(letter, col_name), ...]

        # First column = row number; remaining = Excel letter labels
        table.add_column("", key="_row")
        for letter, _ in col_pairs:
            table.add_column(letter, key=letter)

        # Row 0: header names (like Excel row 1)
        table.add_row(
            Text("", style="dim"),
            *[Text(str(col), style="bold") for _, col in col_pairs],
        )

        # Rows 1-5: actual sample data
        for i, (_, row) in enumerate(df.head(5).iterrows(), start=1):
            table.add_row(
                Text(str(i), style="dim"),
                *[str(row[col]) for _, col in col_pairs],
            )

    def on_button_pressed(self, event: Button.Pressed) -> None:
        bid = event.button.id
        if bid == "btn-back":
            self.app.pop_screen()
        elif bid == "btn-add-col":
            self._add_card()
        elif bid == "btn-preview":
            self._go_preview()
        elif bid == "btn-save-mapping":
            self._save_mapping_dialog()
        elif bid and bid.startswith("card-delete-"):
            try:
                idx = int(bid.split("-")[-1])
                self._delete_card(idx)
            except (ValueError, IndexError):
                pass

    def _add_card(self) -> None:
        opts = self._excel_options()
        idx = self._next_card_index
        self._next_card_index += 1
        self._card_indices.append(idx)
        card = ColumnCard(idx, "", "skip", "TEXT", False, opts, self._card_width())
        add_btn = self.query_one("#btn-add-col", Button)
        self.query_one("#card-scroll", Horizontal).mount(card, before=add_btn)

    def _delete_card(self, idx: int) -> None:
        try:
            self.query_one(f"#col-card-{idx}").remove()
            self._card_indices.remove(idx)
        except Exception:
            pass

    def get_current_mapping(self) -> dict:
        columns = []
        for idx in self._card_indices:
            try:
                name = self.query_one(f"#card-name-{idx}", Input).value.strip()
                src_widget = self.query_one(f"#card-source-{idx}", Select)
                type_widget = self.query_one(f"#card-type-{idx}", Select)
                source = (
                    str(src_widget.value)
                    if src_widget.value is not Select.NULL
                    else "skip"
                )
                col_type = (
                    str(type_widget.value)
                    if type_widget.value is not Select.NULL
                    else "TEXT"
                )
                unique = self.query_one(f"#card-unique-{idx}", Checkbox).value
                columns.append({"name": name, "source": source, "type": col_type, "unique": unique})
            except Exception:
                continue
        return {
            "sheet_name": self.app.sheet_name,
            "header_row": self.app.header_row,
            "start_col": self.app.start_col,
            "target_table": self.app.target_table,
            "columns": columns,
        }

    def _go_preview(self) -> None:
        mapping = self.get_current_mapping()
        col_map = {
            entry["source"]: entry["name"]
            for entry in mapping["columns"]
            if entry["source"] != "skip" and entry["name"]
        }
        if not col_map:
            self.app.notify("Map at least one column.", severity="warning")
            return
        try:
            df = self.app.mapper.map(col_map)
        except ValueError as e:
            self.app.notify(str(e), severity="error")
            return
        self.app.current_mapping = mapping
        self.app.push_screen(PreviewScreen(df, self.app.conn, self.app.target_table))

    def _save_mapping_dialog(self) -> None:
        def _on_dismiss(path: str | None) -> None:
            if not path:
                return
            try:
                save_mapping(path, self.get_current_mapping())
                self.app.notify(f"Mapping saved to {path}")
            except Exception as e:
                self.app.notify(str(e), severity="error")

        self.app.push_screen(FilePickerModal("Save mapping JSON"), _on_dismiss)


# ─────────────────────────────────────────────────────────────
# Screen 4: PreviewScreen
# ─────────────────────────────────────────────────────────────

class PreviewScreen(Screen):
    def __init__(self, df: pd.DataFrame, conn: sqlite3.Connection, table: str) -> None:
        super().__init__()
        self._df = df
        self._conn = conn
        self._table = table
        self._all_rows: list[dict] = []
        self._display_rows: list[dict] = []
        self._is_separator: list[bool] = []
        self._columns: list[str] = list(df.columns)

    def compose(self) -> ComposeResult:
        with Horizontal(id="search-bar"):
            for i, col in enumerate(self._columns):
                yield Input(placeholder=str(col), id=f"search-{i}")
        yield DataTable(id="preview-table")
        with Horizontal(id="bottom-bar"):
            yield Static("", id="row-count")
            yield Button("Back", id="btn-back")
            yield Button("Save mapping to JSON", id="btn-save-mapping")
            yield Button("Next", variant="primary", id="btn-next")

    def on_mount(self) -> None:
        self._all_rows = self._df.to_dict(orient="records")
        self._display_rows, self._is_separator = build_display_rows(self._df)
        table = self.query_one("#preview-table", DataTable)
        for col in self._columns:
            table.add_column(str(col), key=str(col))
        self._refresh_table(self._display_rows, self._is_separator)

    def _refresh_table(self, rows: list[dict], is_sep: list[bool]) -> None:
        table = self.query_one("#preview-table", DataTable)
        table.clear()
        for row, sep in zip(rows, is_sep):
            if sep:
                cells = [Text("…", style="dim italic")] * len(self._columns)
            else:
                cells = [str(row.get(col, "")) for col in self._columns]
            table.add_row(*cells)
        data_count = sum(1 for s in is_sep if not s)
        total_data = sum(1 for s in self._is_separator if not s)
        self.query_one("#row-count", Static).update(
            f"Showing {data_count} / {total_data} data rows"
        )

    def on_input_changed(self, event: Input.Changed) -> None:
        if not event.input.id or not event.input.id.startswith("search-"):
            return
        filters: dict[str, str] = {}
        for i, col in enumerate(self._columns):
            v = self.query_one(f"#search-{i}", Input).value.strip()
            if v:
                filters[col] = v.lower()

        if not filters:
            self._refresh_table(self._display_rows, self._is_separator)
            return

        filtered = [
            row for row in self._all_rows
            if all(
                filters[col] in str(row.get(col, "")).lower()
                for col in filters
            )
        ]
        if len(filtered) < 50:
            self._refresh_table(filtered, [False] * len(filtered))
        else:
            rows, seps = build_display_rows(pd.DataFrame(filtered))
            self._refresh_table(rows, seps)

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "btn-back":
            if isinstance(self.app, ExcelPreviewApp):
                self.app.exit()
            else:
                self.app.pop_screen()
        elif event.button.id == "btn-next":
            self.app.push_screen(
                ImportOptionsScreen(self._df, self._conn, self._table)
            )
        elif event.button.id == "btn-save-mapping":
            self._save_mapping_dialog()

    def _save_mapping_dialog(self) -> None:
        if not isinstance(self.app, ExcelMapperApp):
            return
        mapping = getattr(self.app, "current_mapping", None)
        if not mapping:
            self.app.notify("No mapping available to save.", severity="warning")
            return

        def _on_dismiss(path: str | None) -> None:
            if not path:
                return
            try:
                save_mapping(path, mapping)
                self.app.notify(f"Mapping saved to {path}")
            except Exception as e:
                self.app.notify(str(e), severity="error")

        self.app.push_screen(FilePickerModal("Save mapping JSON"), _on_dismiss)


# ─────────────────────────────────────────────────────────────
# Screen 5: ImportOptionsScreen
# ─────────────────────────────────────────────────────────────

class ImportOptionsScreen(Screen):
    def __init__(
        self, df: pd.DataFrame, conn: sqlite3.Connection, table: str
    ) -> None:
        super().__init__()
        self._df = df
        self._conn = conn
        self._table = table
        self._columns = list(df.columns)

    def compose(self) -> ComposeResult:
        col_options = [(col, col) for col in self._columns]
        with Vertical(id="import-options-content"):
            yield Label("How do you want to import?")
            with RadioSet(id="import-mode"):
                yield RadioButton("Add  —  append rows to the table", value=True)
                yield RadioButton("Replace  —  drop and recreate the table")
                yield RadioButton("Left join on …")
            yield Label("Join column", id="join-col-label")
            yield Select(
                options=col_options,
                allow_blank=False,
                id="join-col-select",
            )
            with Horizontal(id="import-buttons"):
                yield Button("Back", id="btn-back")
                yield Button("Import", variant="primary", id="btn-import")

    def on_mount(self) -> None:
        self._set_join_visible(False)

    def _set_join_visible(self, visible: bool) -> None:
        self.query_one("#join-col-label").display = visible
        self.query_one("#join-col-select").display = visible

    def on_radio_set_changed(self, event: RadioSet.Changed) -> None:
        self._set_join_visible(event.radio_set.pressed_index == 2)

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "btn-back":
            self.app.pop_screen()
        elif event.button.id == "btn-import":
            self._do_import()

    @work
    async def _do_import(self) -> None:
        mode = self.query_one("#import-mode", RadioSet).pressed_index
        try:
            n = len(self._df)
            if mode == 0:
                await asyncio.to_thread(
                    self._df.to_sql, self._table, self._conn,
                    if_exists="append", index=False,
                )
            elif mode == 1:
                await asyncio.to_thread(
                    self._df.to_sql, self._table, self._conn,
                    if_exists="replace", index=False,
                )
            elif mode == 2:
                join_col = self.query_one("#join-col-select", Select).value
                if join_col is Select.NULL:
                    self.app.notify("Select a join column.", severity="warning")
                    return
                await asyncio.to_thread(self._upsert, str(join_col))

            if isinstance(self.app, ExcelPreviewApp):
                self.app.confirmed = True
            self.app.notify(f"Imported {n} rows into '{self._table}'.")
            self.app.push_screen(DBPreviewScreen(self._conn, self._table))
        except Exception as e:
            self.app.notify(str(e), severity="error")

    def _upsert(self, join_col: str) -> None:
        cols = self._columns
        placeholders = ", ".join(["?" for _ in cols])
        update_cols = [c for c in cols if c != join_col]
        update_set = ", ".join(f"{c}=excluded.{c}" for c in update_cols)
        sql = (
            f"INSERT INTO {self._table} ({', '.join(cols)}) "
            f"VALUES ({placeholders}) "
            f"ON CONFLICT({join_col}) DO UPDATE SET {update_set}"
        )
        cursor = self._conn.cursor()
        cursor.executemany(sql, self._df.values.tolist())
        self._conn.commit()


# ─────────────────────────────────────────────────────────────
# Screen 6: DBPreviewScreen
# ─────────────────────────────────────────────────────────────

class DBPreviewScreen(Screen):
    def __init__(self, conn: sqlite3.Connection, table: str) -> None:
        super().__init__()
        self._conn = conn
        self._table = table
        try:
            df = pd.read_sql(f'SELECT * FROM "{table}"', conn)
        except Exception:
            df = pd.DataFrame()
        self._df = df
        self._columns = list(df.columns)
        self._all_rows = df.to_dict(orient="records")
        self._display_rows, self._is_separator = build_display_rows(df)

    def compose(self) -> ComposeResult:
        yield Static(
            f"Database  ·  {self._table}  ·  {len(self._df)} rows",
            id="db-preview-title",
        )
        with Horizontal(id="search-bar"):
            for i, col in enumerate(self._columns):
                yield Input(placeholder=str(col), id=f"search-{i}")
        yield DataTable(id="preview-table")
        with Horizontal(id="bottom-bar"):
            yield Static("", id="row-count")
            yield Button("Done", variant="primary", id="btn-done")

    def on_mount(self) -> None:
        table = self.query_one("#preview-table", DataTable)
        for col in self._columns:
            table.add_column(str(col), key=str(col))
        self._refresh_table(self._display_rows, self._is_separator)

    def _refresh_table(self, rows: list[dict], is_sep: list[bool]) -> None:
        table = self.query_one("#preview-table", DataTable)
        table.clear()
        for row, sep in zip(rows, is_sep):
            if sep:
                cells = [Text("…", style="dim italic")] * len(self._columns)
            else:
                cells = [str(row.get(col, "")) for col in self._columns]
            table.add_row(*cells)
        data_count = sum(1 for s in is_sep if not s)
        total_data = sum(1 for s in self._is_separator if not s)
        self.query_one("#row-count", Static).update(
            f"Showing {data_count} / {total_data} rows"
        )

    def on_input_changed(self, event: Input.Changed) -> None:
        if not event.input.id or not event.input.id.startswith("search-"):
            return
        filters: dict[str, str] = {}
        for i, col in enumerate(self._columns):
            v = self.query_one(f"#search-{i}", Input).value.strip()
            if v:
                filters[col] = v.lower()
        if not filters:
            self._refresh_table(self._display_rows, self._is_separator)
            return
        filtered = [
            row for row in self._all_rows
            if all(filters[col] in str(row.get(col, "")).lower() for col in filters)
        ]
        if len(filtered) < 50:
            self._refresh_table(filtered, [False] * len(filtered))
        else:
            rows, seps = build_display_rows(pd.DataFrame(filtered))
            self._refresh_table(rows, seps)

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "btn-done":
            self.app.exit()


# ─────────────────────────────────────────────────────────────
# ExcelMapperApp
# ─────────────────────────────────────────────────────────────

class ExcelMapperApp(App):
    SCREENS = {
        "file": FileScreen,
        "sheet": SheetTableScreen,
        "mapping": MappingScreen,
        "preview": PreviewScreen,
        "import": ImportOptionsScreen,
        "db-preview": DBPreviewScreen,
    }
    CSS_PATH = "ExcelMapperApp.tcss"

    def __init__(self) -> None:
        super().__init__()
        self.excel_path: str = ""
        self.db_path: str = ""
        self.sheet_name: str | int = 0
        self.header_row: int = 1
        self.start_col: str = "A"
        self.target_table: str = ""
        self.mapper: ExcelMapper | None = None
        self.conn: sqlite3.Connection | None = None
        self.loaded_mapping: dict | None = None
        self.current_mapping: dict | None = None
        self.existing_table_cols: list[tuple[str, str]] = []

    def on_mount(self) -> None:
        self.push_screen(FileScreen())

    def on_exit(self) -> None:
        if self.conn:
            self.conn.close()


# ─────────────────────────────────────────────────────────────
# ExcelPreviewApp — lightweight wrapper called from CLI
# ─────────────────────────────────────────────────────────────

class ExcelPreviewApp(App):
    CSS_PATH = "ExcelMapperApp.tcss"

    def __init__(self, df: pd.DataFrame, conn: sqlite3.Connection, table: str) -> None:
        super().__init__()
        self.confirmed = False
        self._df = df
        self._conn = conn
        self._table = table

    def on_mount(self) -> None:
        self.push_screen(PreviewScreen(self._df, self._conn, self._table))
