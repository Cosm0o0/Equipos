from __future__ import annotations

import threading
from datetime import datetime, date
from pathlib import Path
from queue import Queue, Empty
from typing import Any, Iterable
import tkinter as tk
import unicodedata
import textwrap

import customtkinter as ctk
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from tkinter import filedialog, messagebox, ttk
import sys
from pathlib import Path



APP_TITLE = "Consulta de Registros"
APP_VERSION = "v2.2"
SHEET_ID = "1H-55gGDRUFl8vlymkVeuF8OwSQo_8OiFK8ghMRGk3ZI"
WORKSHEET_NAME = "BD"
ALLOWED_NAMES = [
    "ERIKA UCHUYA TROCONES",
    "JAVIER KLUIVERT CONDOR SANCHEZ",
    "Ximena Jamilet Montoya Calderon",
]
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
HIDDEN_COLUMN_INDICES = {0, 3, 7, 12, 14, 15, 16, 17, 18, 19, 20}


def resource_path(filename: str) -> Path:
    if getattr(sys, "frozen", False):
        # Ejecutándose desde el .exe
        base_dir = Path(sys.executable).parent
    else:
        # Ejecutándose desde Python
        base_dir = Path(__file__).resolve().parent

    return base_dir / filename


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split()).upper()



def safe_string(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def load_sheet_data(
    worksheet_name: str,
    *,
    filter_col_index: int | None = None,
    allowed_values: Iterable[Any] | None = None,
    date_filter_col_index: int | None = None,
    date_filter_from: date | None = None,
    date_filter_to: date | None = None,
) -> tuple[pd.DataFrame, str]:
    creds_path = resource_path("credentials.json")
    if not creds_path.exists():
        raise FileNotFoundError(
            f"No se encontró {creds_path.name}. Debe estar en la misma carpeta que main.py."
        )

    credentials = Credentials.from_service_account_file(str(creds_path), scopes=SCOPES)
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(SHEET_ID)
    worksheet = spreadsheet.worksheet(worksheet_name)

    values = worksheet.get_all_values()
    if not values:
        return pd.DataFrame(), worksheet_name

    raw_headers = values[0]
    max_cols = max(len(row) for row in values)

    headers: list[str] = []
    seen: dict[str, int] = {}
    for index in range(max_cols):
        base = safe_string(raw_headers[index]) if index < len(raw_headers) else ""
        if not base:
            base = f"Col {index + 1}"
        count = seen.get(base, 0)
        seen[base] = count + 1
        headers.append(base if count == 0 else f"{base} ({count + 1})")

    rows = [list(row) + [""] * (max_cols - len(row)) for row in values[1:]]
    data = pd.DataFrame(rows, columns=headers)

    if data.empty:
        return data, worksheet_name

    if filter_col_index is not None and allowed_values is not None and len(data.columns) > filter_col_index:
        allowed = {normalize_text(x) for x in allowed_values}
        column_values = data.iloc[:, filter_col_index].map(normalize_text)
        data = data.loc[column_values.isin(allowed)].copy()

    if (
        date_filter_col_index is not None
        and (date_filter_from is not None or date_filter_to is not None)
        and len(data.columns) > date_filter_col_index
    ):
        parsed_dates = pd.to_datetime(
            data.iloc[:, date_filter_col_index],
            errors="coerce",
            dayfirst=True,
        ).dt.date
        mask = pd.Series(True, index=data.index)
        if date_filter_from is not None:
            mask &= parsed_dates >= date_filter_from
        if date_filter_to is not None:
            mask &= parsed_dates <= date_filter_to
        data = data.loc[mask].copy()

    return data.reset_index(drop=True), worksheet_name


def prepare_visitas_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    result = df.copy()
    if len(result.columns) > 0:
        first_col = result.columns[0]
        result = result.rename(columns={first_col: "Nombre"})
        name_map = {
            "70321862": "ERIKA UCHUYA TROCONES",
            "70122639": "JAVIER KLUIVERT CONDOR SANCHEZ",
            "71406087": "Ximena Jamilet Montoya Calderon",
        }
        result["Nombre"] = result["Nombre"].map(
            lambda value: name_map.get(normalize_text(value), safe_string(value))
        )

    def split_date_time_column(frame: pd.DataFrame, column_index: int, date_name: str, time_name: str) -> pd.DataFrame:
        if len(frame.columns) <= column_index:
            return frame

        source_col = frame.columns[column_index]
        parsed = pd.to_datetime(frame[source_col], errors="coerce", dayfirst=True)
        date_values = parsed.dt.strftime("%d/%m/%Y").fillna("")
        time_values = parsed.dt.strftime("%H:%M").fillna("")

        before_cols = list(frame.columns[:column_index])
        after_cols = list(frame.columns[column_index + 1 :])

        updated = frame.drop(columns=[source_col]).copy()
        updated[date_name] = date_values
        updated[time_name] = time_values

        ordered_columns = before_cols + [date_name, time_name] + after_cols
        return updated.loc[:, ordered_columns]

    result = split_date_time_column(
        result,
        9,
        "FECHA INICIO DE VISITA",
        "HORA INICIO DE VISITA",
    )
    result = split_date_time_column(
        result,
        11,
        "FECHA FIN DE VISITA",
        "HORA FIN DE VISITA",
    )

    return result


def normalize_search_text(value: Any) -> str:
    text = normalize_text(value)
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")


def unique_display_values(series: pd.Series) -> list[str]:
    seen: dict[str, str] = {}
    for raw_value in series.tolist():
        text = safe_string(raw_value)
        if not text:
            continue
        key = normalize_text(text)
        if key and key not in seen:
            seen[key] = text
    return sorted(seen.values(), key=normalize_search_text)


_CALENDAR_CLASS: Any | None = None


def get_calendar_class() -> Any:
    global _CALENDAR_CLASS
    if _CALENDAR_CLASS is None:
        from tkcalendar import Calendar

        _CALENDAR_CLASS = Calendar
    return _CALENDAR_CLASS


class AutocompletePopup:
    def __init__(
        self,
        entry: ctk.CTkEntry,
        variable: tk.StringVar,
        *,
        max_items: int = 20,
        on_select: Any | None = None,
    ) -> None:
        self.entry = entry
        self.variable = variable
        self.max_items = max_items
        self.on_select = on_select
        self.options: list[str] = []
        self.popup: tk.Toplevel | None = None
        self.listbox: tk.Listbox | None = None
        self._update_after_id: str | None = None
        self._hide_after_id: str | None = None
        self._suspend = False

        self.entry.bind("<KeyRelease>", self._on_key_release, add="+")
        self.entry.bind("<FocusIn>", self._on_focus_in, add="+")
        self.entry.bind("<FocusOut>", self._on_focus_out, add="+")
        self.entry.bind("<Down>", self._focus_listbox, add="+")
        self.entry.bind("<Escape>", self._hide_from_event, add="+")

    def set_options(self, values: Iterable[str]) -> None:
        self.options = list(values)

    def _on_focus_in(self, _event: Any) -> None:
        if self.variable.get().strip():
            self._schedule_update()

    def _on_focus_out(self, _event: Any) -> None:
        if self._hide_after_id is not None:
            self.entry.after_cancel(self._hide_after_id)
        self._hide_after_id = self.entry.after(180, self.hide)

    def _on_key_release(self, event: Any) -> None:
        if self._suspend:
            return
        if event.keysym in {"Up", "Down", "Left", "Right", "Return", "Escape", "Tab", "Shift_L", "Shift_R"}:
            return
        self._schedule_update()

    def _schedule_update(self) -> None:
        if self._update_after_id is not None:
            self.entry.after_cancel(self._update_after_id)
        self._update_after_id = self.entry.after(120, self._update_popup)

    def _update_popup(self) -> None:
        self._update_after_id = None
        query = self.variable.get().strip()
        if not query:
            self.hide()
            return

        matches = [
            value
            for value in self.options
            if normalize_search_text(query) in normalize_search_text(value)
        ]
        if not matches:
            self.hide()
            return

        matches = matches[: self.max_items]
        if self.popup is None:
            self.popup = tk.Toplevel(self.entry.winfo_toplevel())
            self.popup.overrideredirect(True)
            self.popup.attributes("-topmost", True)
            self.popup.configure(bg="#dbeafe")
            self.popup.bind("<FocusOut>", lambda _e: self.hide(), add="+")

            frame = tk.Frame(self.popup, bg="#ffffff", bd=1, relief="solid")
            frame.pack(fill="both", expand=True)

            self.listbox = tk.Listbox(
                frame,
                activestyle="none",
                highlightthickness=0,
                selectborderwidth=0,
                relief="flat",
                bd=0,
                font=("Segoe UI", 10),
                height=min(self.max_items, 8),
                exportselection=False,
            )
            self.listbox.pack(side="left", fill="both", expand=True)

            scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
            scrollbar.pack(side="right", fill="y")
            self.listbox.configure(yscrollcommand=scrollbar.set)

            self.listbox.bind("<ButtonRelease-1>", self._accept_selection, add="+")
            self.listbox.bind("<Return>", self._accept_selection, add="+")
            self.listbox.bind("<Escape>", self._hide_from_event, add="+")

        assert self.listbox is not None
        self.listbox.delete(0, tk.END)
        for value in matches:
            self.listbox.insert(tk.END, value)
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(0)
        self.listbox.activate(0)

        self.popup.update_idletasks()
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
        width = self.entry.winfo_width()
        height = min(220, 26 + (len(matches) * 24))
        self.popup.geometry(f"{width}x{height}+{x}+{y}")
        self.popup.deiconify()
        self.popup.lift()

    def _focus_listbox(self, _event: Any) -> str | None:
        if self.listbox is None or self.popup is None:
            self._schedule_update()
            return None
        self.listbox.focus_set()
        return "break"

    def _accept_selection(self, _event: Any) -> str | None:
        if self.listbox is None:
            return None
        selection = self.listbox.curselection()
        if not selection:
            return None
        value = self.listbox.get(selection[0])
        self._suspend = True
        self.variable.set(value)
        self._suspend = False
        self.hide()
        if self.on_select is not None:
            self.on_select(value)
        return "break"

    def _hide_from_event(self, _event: Any) -> str | None:
        self.hide()
        return "break"

    def hide(self) -> None:
        if self.popup is not None:
            self.popup.withdraw()

    def destroy(self) -> None:
        if self._update_after_id is not None:
            self.entry.after_cancel(self._update_after_id)
            self._update_after_id = None
        if self._hide_after_id is not None:
            self.entry.after_cancel(self._hide_after_id)
            self._hide_after_id = None
        if self.popup is not None:
            self.popup.destroy()
            self.popup = None
            self.listbox = None


class DataPanel(ctk.CTkFrame):
    def __init__(
        self,
        master: Any,
        *,
        panel_title: str,
        worksheet_name: str,
        filter_label: str,
        filter_values: Iterable[str] | None = None,
        raw_filter_values: Iterable[str] | None = None,
        extra_filters: list[dict[str, Any]] | None = None,
        filter_col_index: int | None = None,
        date_filter_label: str | None = None,
        date_filter_col_index: int | None = None,
        hidden_column_indices: set[int] | None = None,
        show_filter_menu: bool = True,
        search_placeholder: str = "Escribe para filtrar filas cargadas...",
        initial_filter_value: str = "Todos",
        loaded_df_transform: Any | None = None,
        load_filter_on_raw_data: bool = True,
        compact_layout: bool = False,
        wrap_cells: bool = False,
        max_column_width: int | None = None,
        tree_row_height: int | None = None,
        show_actions_card: bool = True,
        show_stats_in_controls: bool = False,
    ) -> None:
        super().__init__(master, corner_radius=0, fg_color="#f8fafc")
        self.panel_title = panel_title
        self.worksheet_name = worksheet_name
        self.filter_label = filter_label
        self.filter_values = list(filter_values or [])
        self.raw_filter_values = list(raw_filter_values or self.filter_values)
        self.extra_filters = [dict(spec) for spec in (extra_filters or [])]
        self.filter_col_index = filter_col_index
        self.date_filter_label = date_filter_label
        self.date_filter_col_index = date_filter_col_index
        self.hidden_column_indices = hidden_column_indices or set()
        self.show_filter_menu = show_filter_menu
        self.search_placeholder = search_placeholder
        self.loaded_df_transform = loaded_df_transform
        self.load_filter_on_raw_data = load_filter_on_raw_data
        self.compact_layout = compact_layout
        self.wrap_cells = wrap_cells
        self.max_column_width = max_column_width
        self.tree_row_height = tree_row_height
        self.show_actions_card = show_actions_card
        self.show_stats_in_controls = show_stats_in_controls

        self.queue: Queue[tuple[str, Any]] = Queue()
        self.base_df = pd.DataFrame()
        self.current_df = pd.DataFrame()
        self.last_loaded_at: str = "-"

        self.name_var = ctk.StringVar(value=initial_filter_value)
        self.status_var = ctk.StringVar(value="Listo para cargar registros.")
        self.count_var = ctk.StringVar(value="0")
        self.visible_var = ctk.StringVar(value="0")
        self.sheet_var = ctk.StringVar(value=worksheet_name)
        self.refresh_var = ctk.StringVar(value="-")
        self.search_var = ctk.StringVar(value="")
        self.date_from_var = ctk.StringVar(value="Todas")
        self.date_to_var = ctk.StringVar(value="Todas")
        self.selected_date_from: date | None = None
        self.selected_date_to: date | None = None

        self._extra_filter_rows: list[dict[str, Any]] = []

        self._build_ui()
        self.after(100, self._poll_queue)
        self.after(250, self.refresh_data)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        table_row = 1 if self.show_stats_in_controls else 2
        self.grid_rowconfigure(table_row, weight=1)

        controls = ctk.CTkFrame(self, corner_radius=22, fg_color="#ffffff")
        controls.grid(row=0, column=0, sticky="ew", padx=20, pady=(14 if self.compact_layout else 20, 8))
        controls.grid_columnconfigure(0, weight=1)
        if self.show_actions_card:
            controls.grid_columnconfigure(1, weight=1)

        filter_card = ctk.CTkFrame(controls, corner_radius=18, fg_color="#ffffff")
        filter_card.grid(
            row=0,
            column=0,
            columnspan=2 if self.show_actions_card else 1,
            sticky="nsew",
            padx=(18, 10) if self.show_actions_card else 18,
            pady=(12 if self.compact_layout else 18),
        )
        filter_card.grid_columnconfigure(0, weight=1)
        filter_card.grid_columnconfigure(1, weight=1)
        if not self.show_actions_card:
            filter_card.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(
            filter_card,
            text=self.filter_label,
            font=ctk.CTkFont(size=12 if self.compact_layout else 14, weight="bold"),
            text_color="#1d4ed8",
        ).grid(row=0, column=0, columnspan=3 if not self.show_actions_card else 2, sticky="w", padx=14, pady=(10 if self.compact_layout else 14, 4))

        if self.show_filter_menu and self.filter_values:
            self.quick_filter_menu = ctk.CTkOptionMenu(
                filter_card,
                values=["Todos", *self.filter_values],
                variable=self.name_var,
                command=lambda _value: self._set_filter(self.name_var.get()),
                width=340 if self.compact_layout else 380,
                height=32 if self.compact_layout else 38,
            )
            self.quick_filter_menu.grid(
                row=1,
                column=0,
                columnspan=3 if not self.show_actions_card else 2,
                sticky="ew",
                padx=(14, 14),
                pady=(0, 10 if self.compact_layout else 14),
            )
        else:
            ctk.CTkLabel(
                filter_card,
                text="Columna A filtrada por los IDs autorizados.",
                font=ctk.CTkFont(size=11 if self.compact_layout else 12),
                text_color="#475569",
            ).grid(row=1, column=0, columnspan=3 if not self.show_actions_card else 2, sticky="w", padx=14, pady=(0, 10 if self.compact_layout else 14))

        if self.compact_layout and not self.show_actions_card:
            filters_row = ctk.CTkFrame(filter_card, fg_color="transparent")
            filters_row.grid(row=2, column=0, columnspan=3, sticky="ew", padx=6, pady=(0, 8))
            for i in range(3):
                filters_row.grid_columnconfigure(i, weight=1)

            client_spec = self.extra_filters[0] if len(self.extra_filters) > 0 else {}
            client_var = client_spec.setdefault("variable", ctk.StringVar(value=safe_string(client_spec.get("initial_value", ""))))
            self.client_popup = self._build_autocomplete_filter(
                filters_row,
                row=0,
                column=0,
                label=safe_string(client_spec.get("label", "Cliente")),
                variable=client_var,
                placeholder=safe_string(client_spec.get("placeholder", "Escribe para buscar cliente...")),
                on_select=lambda _value: self._apply_visible_filter(),
            )
            client_spec["widget"] = self.client_popup
            if client_spec.get("values_source") == "data":
                self._extra_filter_rows.append(client_spec)

            type_spec = self.extra_filters[1] if len(self.extra_filters) > 1 else {}
            type_var = type_spec.setdefault("variable", ctk.StringVar(value=safe_string(type_spec.get("initial_value", "Todos")) or "Todos"))
            type_widget = self._build_fixed_filter(
                filters_row,
                row=0,
                column=1,
                label=safe_string(type_spec.get("label", "Tipo Visita")),
                values=type_spec.get("values", []),
                variable=type_var,
            )
            type_spec["widget"] = type_widget

            date_container = ctk.CTkFrame(filters_row, fg_color="transparent")
            date_container.grid(row=0, column=2, sticky="ew", padx=14, pady=(0, 10))
            date_container.grid_columnconfigure(0, weight=1)
            date_container.grid_columnconfigure(1, weight=1)
            date_container.grid_columnconfigure(2, weight=0)
            date_container.grid_columnconfigure(3, weight=0)
            date_container.grid_columnconfigure(4, weight=0)
            ctk.CTkLabel(
                date_container,
                text=self.date_filter_label or "Fecha de visita",
                font=ctk.CTkFont(size=11, weight="bold"),
                text_color="#1d4ed8",
            ).grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 4))

            self.date_from_entry = ctk.CTkEntry(
                date_container,
                textvariable=self.date_from_var,
                state="readonly",
                height=32,
                width=94,
            )
            self.date_from_entry.grid(row=1, column=0, sticky="ew", padx=(0, 6))

            self.date_to_entry = ctk.CTkEntry(
                date_container,
                textvariable=self.date_to_var,
                state="readonly",
                height=32,
                width=94,
            )
            self.date_to_entry.grid(row=1, column=1, sticky="ew", padx=(0, 6))

            self.date_from_button = ctk.CTkButton(
                date_container,
                text="Desde",
                command=lambda: self._open_calendar("from"),
                corner_radius=12,
                height=32,
                width=66,
            )
            self.date_from_button.grid(row=1, column=2, sticky="e", padx=(0, 6))

            self.date_to_button = ctk.CTkButton(
                date_container,
                text="Hasta",
                command=lambda: self._open_calendar("to"),
                corner_radius=12,
                height=32,
                width=66,
            )
            self.date_to_button.grid(row=1, column=3, sticky="e", padx=(0, 6))

            self.date_clear_button = ctk.CTkButton(
                date_container,
                text="Limpiar",
                fg_color="#e2e8f0",
                hover_color="#cbd5e1",
                text_color="#0f172a",
                command=self._clear_date_filter,
                corner_radius=12,
                height=32,
                width=70,
            )
            self.date_clear_button.grid(row=1, column=4, sticky="e")

            action_row = ctk.CTkFrame(filter_card, fg_color="transparent")
            action_row.grid(row=3, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 8))
            action_row.grid_columnconfigure(0, weight=1)
            action_row.grid_columnconfigure(1, weight=0)

            self.search_entry = ctk.CTkEntry(
                action_row,
                placeholder_text=self.search_placeholder,
                textvariable=self.search_var,
                height=32,
            )
            self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
            self.search_entry.bind("<KeyRelease>", lambda _e: self._apply_visible_filter())

            buttons_row = ctk.CTkFrame(action_row, fg_color="transparent")
            buttons_row.grid(row=0, column=1, sticky="e")
            buttons_row.grid_columnconfigure(0, weight=0)
            buttons_row.grid_columnconfigure(1, weight=0)

            self.refresh_button = ctk.CTkButton(
                buttons_row,
                text="Actualizar",
                command=self.refresh_data,
                corner_radius=14,
                height=32,
                width=90,
            )
            self.refresh_button.grid(row=0, column=0, sticky="e", padx=(0, 8))

            self.clear_button = ctk.CTkButton(
                buttons_row,
                text="Limpiar búsqueda",
                fg_color="#e2e8f0",
                hover_color="#cbd5e1",
                text_color="#0f172a",
                command=self._clear_search,
                corner_radius=14,
                height=32,
                width=118,
            )
            self.clear_button.grid(row=0, column=1, sticky="e")

            self.quick_refresh_button = self.refresh_button
        else:
            extra_row = 2
            if self.extra_filters:
                extra_row = self._build_extra_filters(filter_card, start_row=2)

            if self.date_filter_label and self.date_filter_col_index is not None:
                date_row = ctk.CTkFrame(filter_card, fg_color="transparent")
                date_row.grid(row=extra_row, column=0, columnspan=3 if not self.show_actions_card else 2, sticky="ew", padx=14, pady=(0, 10 if self.compact_layout else 14))
                date_row.grid_columnconfigure(0, weight=0)
                date_row.grid_columnconfigure(1, weight=0)
                date_row.grid_columnconfigure(2, weight=0)
                date_row.grid_columnconfigure(3, weight=0)

                ctk.CTkLabel(
                    date_row,
                    text=self.date_filter_label,
                    font=ctk.CTkFont(size=11 if self.compact_layout else 12, weight="bold"),
                    text_color="#1d4ed8",
                ).grid(row=0, column=0, sticky="w", padx=(0, 10), pady=0)

                self.date_from_entry = ctk.CTkEntry(
                    date_row,
                    textvariable=self.date_from_var,
                    state="readonly",
                    width=110 if self.compact_layout else 140,
                )
                self.date_from_entry.grid(row=0, column=1, sticky="w", padx=(0, 8))

                self.date_to_entry = ctk.CTkEntry(
                    date_row,
                    textvariable=self.date_to_var,
                    state="readonly",
                    width=110 if self.compact_layout else 140,
                )
                self.date_to_entry.grid(row=0, column=2, sticky="w", padx=(0, 8))

                self.date_button = ctk.CTkButton(
                    date_row,
                    text="Desde",
                    command=lambda: self._open_calendar("from"),
                    corner_radius=14,
                    height=32 if self.compact_layout else 38,
                    width=96 if self.compact_layout else 120,
                )
                self.date_button.grid(row=0, column=3, sticky="e")

                self.date_to_button = ctk.CTkButton(
                    date_row,
                    text="Hasta",
                    command=lambda: self._open_calendar("to"),
                    corner_radius=14,
                    height=32 if self.compact_layout else 38,
                    width=96 if self.compact_layout else 120,
                )
                self.date_to_button.grid(row=0, column=4, sticky="e", padx=(0, 8))

                self.date_clear_button = ctk.CTkButton(
                    date_row,
                    text="Limpiar",
                    fg_color="#e2e8f0",
                    hover_color="#cbd5e1",
                    text_color="#0f172a",
                    command=self._clear_date_filter,
                    corner_radius=14,
                    height=32 if self.compact_layout else 38,
                    width=84 if self.compact_layout else 100,
                )
                self.date_clear_button.grid(row=0, column=5, sticky="e")

            self.quick_refresh_button = ctk.CTkButton(
                filter_card,
                text="Actualizar",
                command=self.refresh_data,
                corner_radius=14,
                height=32 if self.compact_layout else 38,
                width=90 if self.compact_layout else 118,
            )
            refresh_row = extra_row + 1 if (self.date_filter_label and self.date_filter_col_index is not None) else (extra_row if self.extra_filters else 1)
            self.quick_refresh_button.grid(row=refresh_row, column=1, sticky="e", padx=(0, 14), pady=(0, 6))

            self.search_entry = ctk.CTkEntry(
                filter_card,
                placeholder_text=self.search_placeholder,
                textvariable=self.search_var,
                height=32 if self.compact_layout else 36,
            )
            self.search_entry.grid(row=refresh_row + 1, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 6))
            self.search_entry.bind("<KeyRelease>", lambda _e: self._apply_visible_filter())

            if self.show_actions_card:
                actions_card = ctk.CTkFrame(controls, corner_radius=18, fg_color="#ffffff")
                actions_card.grid(row=0, column=1, sticky="nsew", padx=(10, 18), pady=(12 if self.compact_layout else 18))
                actions_card.grid_columnconfigure(0, weight=1)
                actions_card.grid_columnconfigure(1, weight=1)

                ctk.CTkLabel(
                    actions_card,
                    text="Búsqueda y acciones",
                    font=ctk.CTkFont(size=12 if self.compact_layout else 15, weight="bold"),
                    text_color="#1d4ed8",
                ).grid(row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(10 if self.compact_layout else 14, 6))

                self.search_entry.grid_forget()
                self.search_entry.grid(row=1, column=0, columnspan=2, padx=14, pady=(0, 10 if self.compact_layout else 12), sticky="ew")

                self.refresh_button = ctk.CTkButton(
                    actions_card,
                    text="Actualizar datos",
                    command=self.refresh_data,
                    corner_radius=14,
                    height=32 if self.compact_layout else 40,
                )
                self.refresh_button.grid(row=2, column=0, padx=(14, 6), pady=(0, 10 if self.compact_layout else 14), sticky="ew")

                self.clear_button = ctk.CTkButton(
                    actions_card,
                    text="Limpiar búsqueda",
                    fg_color="#e2e8f0",
                    hover_color="#cbd5e1",
                    text_color="#0f172a",
                    command=self._clear_search,
                    corner_radius=14,
                    height=32 if self.compact_layout else 40,
                )
                self.clear_button.grid(row=2, column=1, padx=(6, 14), pady=(0, 10 if self.compact_layout else 14), sticky="ew")
            else:
                self.refresh_button = ctk.CTkButton(
                    filter_card,
                    text="Actualizar datos",
                    command=self.refresh_data,
                    corner_radius=14,
                    height=32 if self.compact_layout else 38,
                )
                self.refresh_button.grid(row=refresh_row + 2, column=0, sticky="ew", padx=(14, 6), pady=(0, 6))

                self.clear_button = ctk.CTkButton(
                    filter_card,
                    text="Limpiar búsqueda",
                    fg_color="#e2e8f0",
                    hover_color="#cbd5e1",
                    text_color="#0f172a",
                    command=self._clear_search,
                    corner_radius=14,
                    height=32 if self.compact_layout else 38,
                )
                self.clear_button.grid(row=refresh_row + 2, column=1, sticky="ew", padx=(6, 14), pady=(0, 6))

        if self.show_stats_in_controls:
            stats = ctk.CTkFrame(filter_card, fg_color="transparent")
            stats_row = 4 if self.compact_layout and not self.show_actions_card else (refresh_row + 3)
            stats.grid(row=stats_row, column=0, columnspan=3, sticky="ew", padx=6, pady=(0, 8))
            for i in range(4):
                stats.grid_columnconfigure(i, weight=1)

            self._stat_card(stats, 0, "Registros cargados", self.count_var)
            self._stat_card(stats, 1, "Registros visibles", self.visible_var)
            self._stat_card(stats, 2, "Base consultada", self.sheet_var)
            self._stat_card(stats, 3, "Última actualización", self.refresh_var)
        else:
            stats = ctk.CTkFrame(self, corner_radius=22, fg_color="#ffffff")
            stats.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 8 if self.compact_layout else 12))
            for i in range(4):
                stats.grid_columnconfigure(i, weight=1)

            self._stat_card(stats, 0, "Registros cargados", self.count_var)
            self._stat_card(stats, 1, "Registros visibles", self.visible_var)
            self._stat_card(stats, 2, "Base consultada", self.sheet_var)
            self._stat_card(stats, 3, "Última actualización", self.refresh_var)

        table_wrap = ctk.CTkFrame(self, corner_radius=22, fg_color="#ffffff")
        table_wrap.grid(
            row=table_row,
            column=0,
            sticky="nsew",
            padx=20,
            pady=(0, 14 if self.compact_layout else 20),
        )
        table_wrap.grid_rowconfigure(1, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

        table_header = ctk.CTkFrame(table_wrap, fg_color="transparent")
        table_header.grid(row=0, column=0, sticky="ew", padx=18, pady=(12 if self.compact_layout else 18, 6))
        table_header.grid_columnconfigure(0, weight=1)
        table_header.grid_columnconfigure(1, weight=0)
        ctk.CTkLabel(
            table_header,
            text=self.panel_title,
            font=ctk.CTkFont(size=17 if self.compact_layout else 18, weight="bold"),
            text_color="#1d4ed8",
        ).grid(row=0, column=0, sticky="w")

        self.export_button = ctk.CTkButton(
            table_header,
            text="Exportar Excel",
            command=self.export_excel,
            corner_radius=14,
            height=34 if self.compact_layout else 38,
            width=140 if self.compact_layout else 150,
        )
        self.export_button.grid(row=0, column=1, sticky="e", padx=(12, 0), pady=0)

        table_frame = ctk.CTkFrame(table_wrap, corner_radius=16, fg_color="#f8fafc")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 12 if self.compact_layout else 18))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        style = ttk.Style(self)
        self.tree_style_name = f"{self.panel_title.replace(' ', '')}.Treeview"
        self.heading_style_name = f"{self.panel_title.replace(' ', '')}.Treeview.Heading"
        style.configure(
            self.tree_style_name,
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#0f172a",
            rowheight=self.tree_row_height or (56 if self.wrap_cells else 38),
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        style.map(
            self.tree_style_name,
            background=[("selected", "#dbeafe")],
            foreground=[("selected", "#0f172a")],
        )
        style.configure(
            self.heading_style_name,
            background="#dbeafe",
            foreground="#1d4ed8",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )

        self.tree = ttk.Treeview(table_frame, show="headings", style=self.tree_style_name)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns", pady=10, padx=(0, 10))
        self.tree.configure(yscrollcommand=yscroll.set)

        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        xscroll.grid(row=1, column=0, sticky="ew", padx=(10, 0), pady=(0, 10))
        self.tree.configure(xscrollcommand=xscroll.set)

        self.tree.bind("<ButtonRelease-1>", self._on_row_click)

    def _build_extra_filters(self, parent: ctk.CTkFrame, *, start_row: int) -> int:
        highest_row = start_row
        for spec in self.extra_filters:
            kind = str(spec.get("kind", "autocomplete")).lower()
            row = int(spec.get("row", start_row))
            column = int(spec.get("column", 0))
            columnspan = int(spec.get("columnspan", 1))
            highest_row = max(highest_row, row + 1)
            if kind == "fixed":
                widget = self._build_fixed_filter(
                    parent,
                    row=row,
                    column=column,
                    columnspan=columnspan,
                    label=safe_string(spec.get("label")),
                    values=spec.get("values", []),
                    variable=spec.setdefault("variable", ctk.StringVar(value=safe_string(spec.get("initial_value", "Todos")) or "Todos")),
                )
                spec["widget"] = widget
            else:
                widget = self._build_autocomplete_filter(
                    parent,
                    row=row,
                    column=column,
                    columnspan=columnspan,
                    label=safe_string(spec.get("label")),
                    placeholder=safe_string(spec.get("placeholder")),
                    variable=spec.setdefault("variable", ctk.StringVar(value=safe_string(spec.get("initial_value", "")))),
                    on_select=lambda _value: self._apply_visible_filter(),
                )
                spec["widget"] = widget
                if spec.get("values_source") == "data":
                    self._extra_filter_rows.append(spec)
        return highest_row

    def _build_autocomplete_filter(
        self,
        parent: ctk.CTkFrame,
        *,
        row: int,
        column: int,
        label: str,
        variable: tk.StringVar,
        placeholder: str,
        columnspan: int = 1,
        on_select: Any | None = None,
    ) -> AutocompletePopup:
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.grid(row=row, column=column, columnspan=columnspan, sticky="ew", padx=14, pady=(0, 10 if self.compact_layout else 12))
        container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            container,
            text=label,
            font=ctk.CTkFont(size=11 if self.compact_layout else 12, weight="bold"),
            text_color="#1d4ed8",
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        entry = ctk.CTkEntry(
            container,
            textvariable=variable,
            placeholder_text=placeholder,
            height=32 if self.compact_layout else 36,
        )
        entry.grid(row=1, column=0, sticky="ew")
        return AutocompletePopup(entry, variable, on_select=on_select)

    def _build_fixed_filter(
        self,
        parent: ctk.CTkFrame,
        *,
        row: int,
        column: int,
        label: str,
        values: Iterable[str],
        variable: tk.StringVar,
        columnspan: int = 1,
    ) -> ctk.CTkOptionMenu:
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.grid(row=row, column=column, columnspan=columnspan, sticky="ew", padx=14, pady=(0, 10 if self.compact_layout else 12))
        container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            container,
            text=label,
            font=ctk.CTkFont(size=11 if self.compact_layout else 12, weight="bold"),
            text_color="#1d4ed8",
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        menu = ctk.CTkOptionMenu(
            container,
            values=["Todos", *list(values)],
            variable=variable,
            command=lambda _value: self._apply_visible_filter(),
            height=32 if self.compact_layout else 36,
        )
        menu.grid(row=1, column=0, sticky="ew")
        return menu

    def _stat_card(self, parent: ctk.CTkFrame, column: int, title: str, variable: ctk.StringVar) -> None:
        card = ctk.CTkFrame(parent, corner_radius=18, fg_color="#eff6ff")
        card.grid(row=0, column=column, sticky="ew", padx=10 if self.compact_layout else 12, pady=8 if self.compact_layout else 12)
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=11 if self.compact_layout else 12, weight="bold"), text_color="#1d4ed8").pack(
            anchor="w", padx=14, pady=(10 if self.compact_layout else 14, 2)
        )
        ctk.CTkLabel(
            card,
            textvariable=variable,
            font=ctk.CTkFont(size=15 if self.compact_layout else 18, weight="bold"),
            text_color="#0f172a",
            wraplength=240,
            justify="left",
        ).pack(anchor="w", padx=14, pady=(0, 10 if self.compact_layout else 14))

    def _clear_search(self) -> None:
        self.search_var.set("")
        self._apply_visible_filter()

    def _clear_date_filter(self) -> None:
        self.selected_date_from = None
        self.selected_date_to = None
        self.date_from_var.set("Todas")
        self.date_to_var.set("Todas")
        self.refresh_data()

    def _clear_extra_filters(self) -> None:
        for spec in self.extra_filters:
            default_value = "Todos" if str(spec.get("kind", "autocomplete")).lower() == "fixed" else ""
            variable = spec.get("variable")
            if variable is not None:
                variable.set(default_value)

    def _open_calendar(self, target: str) -> None:
        if self.date_filter_label is None:
            return

        from tkcalendar import Calendar

        top = ctk.CTkToplevel(self)
        top.title(self.date_filter_label)
        top.geometry("360x380")
        top.resizable(False, False)
        top.grab_set()

        container = ctk.CTkFrame(top, corner_radius=18, fg_color="#ffffff")
        container.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            container,
            text=f"{self.date_filter_label} - {'Desde' if target == 'from' else 'Hasta'}",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color="#1d4ed8",
        ).pack(anchor="w", padx=14, pady=(14, 10))

        cal = Calendar(container, selectmode="day", date_pattern="dd/mm/yyyy")
        cal.pack(padx=14, pady=(0, 12), fill="both", expand=True)

        button_row = ctk.CTkFrame(container, fg_color="transparent")
        button_row.pack(fill="x", padx=14, pady=(0, 14))

        def apply_date(_event: Any | None = None) -> None:
            selected_date = cal.selection_get()
            if target == "from":
                self.selected_date_from = selected_date
                self.date_from_var.set(selected_date.strftime("%d/%m/%Y"))
            else:
                self.selected_date_to = selected_date
                self.date_to_var.set(selected_date.strftime("%d/%m/%Y"))
            top.destroy()
            self._apply_visible_filter()
            self.refresh_data()

        cal.bind("<<CalendarSelected>>", apply_date)

        ctk.CTkButton(
            button_row,
            text="Cancelar",
            fg_color="#e2e8f0",
            hover_color="#cbd5e1",
            text_color="#0f172a",
            command=top.destroy,
            corner_radius=14,
            height=38,
        ).pack(fill="x")

    def _set_filter(self, value: str) -> None:
        self.name_var.set(value)
        self._apply_visible_filter()

    def refresh_data(self) -> None:
        self.refresh_button.configure(state="disabled", text="Cargando...")
        self.status_var.set("Consultando Google Sheets...")
        self._set_busy(True)
        thread = threading.Thread(target=self._load_async, daemon=True)
        thread.start()

    def export_excel(self) -> None:
        if self.current_df.empty:
            messagebox.showinfo("Exportar Excel", "No hay datos visibles para exportar.")
            return

        default_name = f"export_{self.worksheet_name.strip().replace(' ', '_').lower() or 'datos'}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Guardar Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile=default_name,
        )
        if not path:
            return

        try:
            self.current_df.to_excel(path, index=False)
            messagebox.showinfo("Exportar Excel", f"Archivo exportado correctamente en:\n{path}")
        except Exception as exc:
            messagebox.showerror("Exportar Excel", f"No se pudo exportar el archivo.\n\n{exc}")

    def _load_async(self) -> None:
        try:
            if self.load_filter_on_raw_data and self.filter_values and self.filter_col_index is not None:
                data, sheet_name = load_sheet_data(
                    self.worksheet_name,
                    filter_col_index=self.filter_col_index,
                    allowed_values=self.raw_filter_values,
                    date_filter_col_index=self.date_filter_col_index,
                    date_filter_from=self.selected_date_from,
                    date_filter_to=self.selected_date_to,
                )
            else:
                data, sheet_name = load_sheet_data(
                    self.worksheet_name,
                    date_filter_col_index=self.date_filter_col_index,
                    date_filter_from=self.selected_date_from,
                    date_filter_to=self.selected_date_to,
                )
            if self.loaded_df_transform is not None:
                data = self.loaded_df_transform(data)
            self.queue.put(("data", data))
            self.queue.put(("sheet", sheet_name))
            self.queue.put(("ok", None))
        except Exception as exc:
            self.queue.put(("error", exc))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "data":
                    self.base_df = payload
                    self.current_df = payload
                    self.count_var.set(f"{len(self.base_df):,}".replace(",", "."))
                    self.status_var.set("Datos cargados correctamente.")
                    self._refresh_extra_filter_sources()
                elif kind == "sheet":
                    self.sheet_var.set(payload)
                elif kind == "ok":
                    self.last_loaded_at = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    self.refresh_var.set(self.last_loaded_at)
                    self.refresh_button.configure(state="normal", text="Actualizar datos")
                    self._set_busy(False)
                    self._apply_visible_filter()
                elif kind == "error":
                    self.refresh_button.configure(state="normal", text="Actualizar datos")
                    self._set_busy(False)
                    self.status_var.set("No se pudo cargar la información.")
                    messagebox.showerror("Error al consultar la hoja", str(payload))
        except Empty:
            pass
        finally:
            self.after(120, self._poll_queue)

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.quick_refresh_button.configure(state=state)
        self.refresh_button.configure(state=state)
        self.search_entry.configure(state=state)
        self.clear_button.configure(state=state)
        for spec in self.extra_filters:
            widget = spec.get("widget")
            if widget is None:
                continue
            if hasattr(widget, "entry"):
                widget.entry.configure(state=state)
            else:
                widget.configure(state=state)

    def _refresh_extra_filter_sources(self) -> None:
        for spec in self.extra_filters:
            if spec.get("values_source") != "data":
                continue
            column_index = spec.get("column_index")
            widget = spec.get("widget")
            if column_index is None or widget is None or self.base_df.empty:
                continue
            if len(self.base_df.columns) <= int(column_index):
                continue
            options = unique_display_values(self.base_df.iloc[:, int(column_index)])
            if hasattr(widget, "set_options"):
                widget.set_options(options)

    def _apply_visible_filter(self) -> None:
        if self.base_df.empty:
            self.visible_var.set("0")
            return

        df = self.base_df.copy()
        if self.filter_values and self.filter_col_index is not None:
            selected = normalize_text(self.name_var.get())
            if selected != "TODOS":
                df = df.loc[df.iloc[:, self.filter_col_index].map(normalize_text) == selected].copy()

        for spec in self.extra_filters:
            column_index = spec.get("column_index")
            if column_index is None or len(df.columns) <= int(column_index):
                continue
            kind = str(spec.get("kind", "autocomplete")).lower()
            variable = spec.get("variable")
            if variable is None:
                continue
            selected = safe_string(variable.get()).strip()
            if not selected or normalize_text(selected) == "TODOS":
                continue
            column_values = df.iloc[:, int(column_index)].map(normalize_text)
            if kind == "fixed":
                df = df.loc[column_values == normalize_text(selected)].copy()
            else:
                query = normalize_search_text(selected)
                df = df.loc[column_values.str.contains(query, na=False)].copy()

        query = self.search_var.get().strip().lower()
        if query:
            mask = df.astype(str).apply(lambda col: col.str.lower().str.contains(query, na=False))
            df = df.loc[mask.any(axis=1)].copy()

        self.current_df = df
        self.visible_var.set(f"{len(df):,}".replace(",", "."))
        self._render_table(df)

    def _render_table(self, df: pd.DataFrame) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        if df is None or df.empty:
            self.tree["columns"] = ("mensaje",)
            self.tree.heading("mensaje", text="Sin resultados")
            self.tree.column("mensaje", width=900, anchor="center", stretch=True)
            self.tree.insert("", "end", values=("No hay registros para el filtro seleccionado.",))
            return

        columns = [str(col) for idx, col in enumerate(df.columns) if idx not in self.hidden_column_indices]
        self.tree["columns"] = columns

        for col in columns:
            self.tree.heading(col, text=col, anchor="center")

        widths = self._estimate_widths(df, columns)
        for col, width in widths.items():
            self.tree.column(col, width=width, minwidth=120, anchor="center", stretch=False)

        for _, row in df.iterrows():
            values = [
                self._format_cell_value(safe_string(row[col]), widths[col])
                for col in columns
            ]
            self.tree.insert("", "end", values=values)

    def _estimate_widths(self, df: pd.DataFrame, columns: Iterable[str]) -> dict[str, int]:
        widths: dict[str, int] = {}
        sample = df.head(50)
        cap = self.max_column_width if self.max_column_width is not None else (320 if self.wrap_cells else 560)
        for col in columns:
            best = len(str(col)) * 11 + 32
            if not sample.empty:
                max_len = max([len(str(col))] + [len(safe_string(v)) for v in sample[col].tolist()])
                best = min(max(140, max_len * 9 + 32), cap)
            widths[col] = best
        return widths

    def _format_cell_value(self, value: str, width: int) -> str:
        if not self.wrap_cells or not value:
            return value

        wrap_width = max(12, min(26, max(10, width // 9)))
        parts = value.replace("\r", " ").split("\n")
        wrapped = [textwrap.fill(part, width=wrap_width, break_long_words=False, break_on_hyphens=False) for part in parts if part]
        return "\n".join(wrapped) if wrapped else value

    def _on_row_click(self, _event: Any) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0], "values")
        if not values:
            return
        snippet = " | ".join(str(v) for v in values[:5])
        self.status_var.set(f"Fila seleccionada: {snippet}")


class RegistrosPanel(ctk.CTkFrame):
    def __init__(self, master: Any) -> None:
        super().__init__(master, corner_radius=0, fg_color="#f8fafc")

        self.queue: Queue[tuple[str, Any]] = Queue()
        self.base_df = pd.DataFrame()
        self.current_df = pd.DataFrame()
        self.last_loaded_at: str = "-"

        self.client_var = ctk.StringVar(value="")
        self.brand_var = ctk.StringVar(value="")
        self.line_var = ctk.StringVar(value="")
        self.representative_var = ctk.StringVar(value="Todos")
        self.confidence_var = ctk.StringVar(value="Todos")
        self.date_from_var = ctk.StringVar(value="Todas")
        self.date_to_var = ctk.StringVar(value="Todas")
        self.search_var = ctk.StringVar(value="")
        self.status_var = ctk.StringVar(value="Listo para cargar registros.")
        self.count_var = ctk.StringVar(value="0")
        self.visible_var = ctk.StringVar(value="0")
        self.sheet_var = ctk.StringVar(value=WORKSHEET_NAME)
        self.refresh_var = ctk.StringVar(value="-")
        self.selected_date_from: date | None = None
        self.selected_date_to: date | None = None

        self._client_options: list[str] = []
        self._brand_options: list[str] = []
        self._line_options: list[str] = []
        self._representative_options: list[str] = list(ALLOWED_NAMES)
        self.extra_filters: list[dict[str, Any]] = []

        self.representative_popup: Any | None = None
        self.client_popup: AutocompletePopup | None = None
        self.brand_popup: AutocompletePopup | None = None
        self.line_popup: AutocompletePopup | None = None

        self._build_ui()
        self.after(100, self._poll_queue)
        self.after(250, self.refresh_data)

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        controls = ctk.CTkFrame(self, corner_radius=22, fg_color="#ffffff")
        controls.grid(row=0, column=0, sticky="ew", padx=20, pady=(10, 4))
        controls.grid_columnconfigure(0, weight=1)
        controls.grid_columnconfigure(1, weight=1)
        controls.grid_columnconfigure(2, weight=1)

        filter_card = ctk.CTkFrame(controls, corner_radius=18, fg_color="#ffffff")
        filter_card.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=18, pady=(8, 6))
        for i in range(3):
            filter_card.grid_columnconfigure(i, weight=1)

        ctk.CTkLabel(
            filter_card,
            text="Filtros de Registros",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#1d4ed8",
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(8, 2))

        self.representative_popup = self._build_representative_field(
            filter_card,
            row=1,
            column=0,
            label="Representante",
            variable=self.representative_var,
        )
        self.client_popup = self._build_autocomplete_field(
            filter_card,
            row=1,
            column=1,
            label="Cliente",
            variable=self.client_var,
            placeholder="Escribe para buscar cliente...",
            on_select=lambda _value: self._apply_visible_filter(),
        )
        self.brand_popup = self._build_autocomplete_field(
            filter_card,
            row=1,
            column=2,
            label="Marca",
            variable=self.brand_var,
            placeholder="Escribe para buscar marca...",
            on_select=lambda _value: self._apply_visible_filter(),
        )
        self._build_date_field(filter_card, row=2, column=0)
        self._build_confidence_field(filter_card, row=2, column=1)
        self.line_popup = self._build_autocomplete_field(
            filter_card,
            row=2,
            column=2,
            label="Línea",
            variable=self.line_var,
            placeholder="Escribe para buscar línea...",
            on_select=lambda _value: self._apply_visible_filter(),
        )
        self.search_entry = ctk.CTkEntry(
            filter_card,
            placeholder_text="Escribe para filtrar filas cargadas...",
            textvariable=self.search_var,
            height=34,
        )
        self.search_entry.grid(row=3, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 10))
        self.search_entry.bind("<KeyRelease>", lambda _e: self._apply_visible_filter())

        button_row = ctk.CTkFrame(filter_card, fg_color="transparent")
        button_row.grid(row=4, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 8))
        for i in range(4):
            button_row.grid_columnconfigure(i, weight=1)

        self.date_refresh_button = ctk.CTkButton(
            button_row,
            text="Limpiar filtros",
            fg_color="#e2e8f0",
            hover_color="#cbd5e1",
            text_color="#0f172a",
            command=self._clear_filters,
            corner_radius=14,
            height=32,
        )
        self.date_refresh_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.refresh_button = ctk.CTkButton(
            button_row,
            text="Actualizar datos",
            command=self.refresh_data,
            corner_radius=14,
            height=32,
        )
        self.refresh_button.grid(row=0, column=1, sticky="ew", padx=(6, 6))

        self.clear_button = ctk.CTkButton(
            button_row,
            text="Limpiar búsqueda",
            fg_color="#e2e8f0",
            hover_color="#cbd5e1",
            text_color="#0f172a",
            command=self._clear_search,
            corner_radius=14,
            height=32,
        )
        self.clear_button.grid(row=0, column=2, sticky="ew", padx=(6, 6))

        self.export_button = ctk.CTkButton(
            button_row,
            text="Exportar Excel",
            command=self.export_excel,
            corner_radius=14,
            height=32,
        )
        self.export_button.grid(row=0, column=3, sticky="ew", padx=(6, 0))

        stats = ctk.CTkFrame(self, corner_radius=18, fg_color="#ffffff")
        stats.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 4))
        for i in range(4):
            stats.grid_columnconfigure(i, weight=1)

        self._stat_card(stats, 0, "Registros cargados", self.count_var)
        self._stat_card(stats, 1, "Registros visibles", self.visible_var)
        self._stat_card(stats, 2, "Base consultada", self.sheet_var)
        self._stat_card(stats, 3, "Última actualización", self.refresh_var)

        table_wrap = ctk.CTkFrame(self, corner_radius=22, fg_color="#ffffff")
        table_wrap.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 10))
        table_wrap.grid_rowconfigure(1, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

        table_header = ctk.CTkFrame(table_wrap, fg_color="transparent")
        table_header.grid(row=0, column=0, sticky="ew", padx=18, pady=(8, 4))
        table_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            table_header,
            text="Resultados",
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color="#1d4ed8",
        ).grid(row=0, column=0, sticky="w")

        table_frame = ctk.CTkFrame(table_wrap, corner_radius=16, fg_color="#f8fafc")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 8))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(table_frame, show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns", pady=10, padx=(0, 10))
        self.tree.configure(yscrollcommand=yscroll.set)

        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        xscroll.grid(row=1, column=0, sticky="ew", padx=(10, 0), pady=(0, 10))
        self.tree.configure(xscrollcommand=xscroll.set)

        self.tree.bind("<ButtonRelease-1>", self._on_row_click)

    def _build_autocomplete_field(
        self,
        parent: ctk.CTkFrame,
        *,
        row: int,
        column: int,
        label: str,
        variable: tk.StringVar,
        placeholder: str,
        on_select: Any | None = None,
        columnspan: int = 1,
    ) -> AutocompletePopup:
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.grid(row=row, column=column, columnspan=columnspan, sticky="ew", padx=14, pady=(0, 12))
        container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            container,
            text=label,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1d4ed8",
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        entry = ctk.CTkEntry(
            container,
            textvariable=variable,
            placeholder_text=placeholder,
            height=36,
        )
        entry.grid(row=1, column=0, sticky="ew")
        return AutocompletePopup(entry, variable, on_select=on_select)

    def _build_representative_field(
        self,
        parent: ctk.CTkFrame,
        *,
        row: int,
        column: int,
        label: str,
        variable: ctk.StringVar,
    ) -> ctk.CTkOptionMenu:
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.grid(row=row, column=column, sticky="ew", padx=14, pady=(0, 12))
        container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            container,
            text=label,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1d4ed8",
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        menu = ctk.CTkOptionMenu(
            container,
            values=["Todos", *self._representative_options],
            variable=variable,
            command=lambda _value: self._apply_visible_filter(),
            height=36,
        )
        menu.grid(row=1, column=0, sticky="ew")
        return menu

    def _build_date_field(self, parent: ctk.CTkFrame, *, row: int, column: int) -> None:
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.grid(row=row, column=column, sticky="ew", padx=14, pady=(0, 12))
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=0)
        container.grid_columnconfigure(2, weight=1)
        container.grid_columnconfigure(3, weight=0)
        container.grid_columnconfigure(4, weight=0)

        ctk.CTkLabel(
            container,
            text="Fecha de Venta",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1d4ed8",
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        self.date_from_entry = ctk.CTkEntry(
            container,
            textvariable=self.date_from_var,
            state="readonly",
            height=36,
        )
        self.date_from_entry.grid(row=1, column=0, sticky="ew", padx=(0, 8))

        self.date_to_entry = ctk.CTkEntry(
            container,
            textvariable=self.date_to_var,
            state="readonly",
            height=36,
        )
        self.date_to_entry.grid(row=1, column=1, sticky="ew", padx=(0, 8))

        ctk.CTkButton(
            container,
            text="Desde",
            command=lambda: self._open_calendar("from"),
            corner_radius=12,
            width=96,
            height=36,
        ).grid(row=1, column=2, sticky="e", padx=(0, 8))

        ctk.CTkButton(
            container,
            text="Hasta",
            command=lambda: self._open_calendar("to"),
            corner_radius=12,
            width=96,
            height=36,
        ).grid(row=1, column=3, sticky="e", padx=(0, 8))

        ctk.CTkButton(
            container,
            text="Limpiar",
            fg_color="#e2e8f0",
            hover_color="#cbd5e1",
            text_color="#0f172a",
            command=self._clear_date_filter,
            corner_radius=12,
            width=84,
            height=36,
        ).grid(row=1, column=4, sticky="e")

    def _build_confidence_field(self, parent: ctk.CTkFrame, *, row: int, column: int) -> None:
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.grid(row=row, column=column, sticky="ew", padx=14, pady=(0, 12))
        container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            container,
            text="Confiabilidad",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1d4ed8",
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        self.confidence_menu = ctk.CTkOptionMenu(
            container,
            values=["Todos", "0%", "25%", "50%", "75%", "100%"],
            variable=self.confidence_var,
            command=lambda _value: self._apply_visible_filter(),
            height=36,
        )
        self.confidence_menu.grid(row=1, column=0, sticky="ew")

    def _stat_card(self, parent: ctk.CTkFrame, column: int, title: str, variable: ctk.StringVar) -> None:
        card = ctk.CTkFrame(parent, corner_radius=18, fg_color="#eff6ff")
        card.grid(row=0, column=column, sticky="ew", padx=12, pady=12)
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=12, weight="bold"), text_color="#1d4ed8").pack(
            anchor="w", padx=14, pady=(14, 2)
        )
        ctk.CTkLabel(
            card,
            textvariable=variable,
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#0f172a",
            wraplength=240,
            justify="left",
        ).pack(anchor="w", padx=14, pady=(0, 14))

    def _clear_search(self) -> None:
        self.search_var.set("")
        self._apply_visible_filter()

    def _clear_filters(self) -> None:
        self.representative_var.set("Todos")
        self.client_var.set("")
        self.brand_var.set("")
        self.line_var.set("")
        self.confidence_var.set("Todos")
        self.selected_date_from = None
        self.selected_date_to = None
        self.date_from_var.set("Todas")
        self.date_to_var.set("Todas")
        self._clear_extra_filters()
        self._apply_visible_filter()

    def _clear_date_filter(self) -> None:
        self.selected_date_from = None
        self.selected_date_to = None
        self.date_from_var.set("Todas")
        self.date_to_var.set("Todas")
        self._apply_visible_filter()

    def _open_calendar(self, target: str) -> None:
        Calendar = get_calendar_class()

        top = ctk.CTkToplevel(self)
        top.title("Fecha de Venta")
        top.geometry("360x380")
        top.resizable(False, False)
        top.grab_set()

        container = ctk.CTkFrame(top, corner_radius=18, fg_color="#ffffff")
        container.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            container,
            text=f"Fecha de Venta - {'Desde' if target == 'from' else 'Hasta'}",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color="#1d4ed8",
        ).pack(anchor="w", padx=14, pady=(14, 10))

        cal = Calendar(container, selectmode="day", date_pattern="dd/mm/yyyy")
        cal.pack(padx=14, pady=(0, 12), fill="both", expand=True)

        button_row = ctk.CTkFrame(container, fg_color="transparent")
        button_row.pack(fill="x", padx=14, pady=(0, 14))

        def apply_date(_event: Any | None = None) -> None:
            selected_date = cal.selection_get()
            if target == "from":
                self.selected_date_from = selected_date
                self.date_from_var.set(selected_date.strftime("%d/%m/%Y"))
            else:
                self.selected_date_to = selected_date
                self.date_to_var.set(selected_date.strftime("%d/%m/%Y"))
            top.destroy()
            self._apply_visible_filter()

        cal.bind("<<CalendarSelected>>", apply_date)

        ctk.CTkButton(
            button_row,
            text="Cancelar",
            fg_color="#e2e8f0",
            hover_color="#cbd5e1",
            text_color="#0f172a",
            command=top.destroy,
            corner_radius=14,
            height=38,
        ).pack(fill="x")

    def refresh_data(self) -> None:
        self.refresh_button.configure(state="disabled", text="Cargando...")
        self.date_refresh_button.configure(state="disabled")
        self.status_var.set("Consultando Google Sheets...")
        self._set_busy(True)
        thread = threading.Thread(target=self._load_async, daemon=True)
        thread.start()

    def export_excel(self) -> None:
        if self.current_df.empty:
            messagebox.showinfo("Exportar Excel", "No hay datos visibles para exportar.")
            return

        default_name = f"export_{WORKSHEET_NAME.strip().replace(' ', '_').lower() or 'datos'}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Guardar Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile=default_name,
        )
        if not path:
            return

        try:
            self.current_df.to_excel(path, index=False)
            messagebox.showinfo("Exportar Excel", f"Archivo exportado correctamente en:\n{path}")
        except Exception as exc:
            messagebox.showerror("Exportar Excel", f"No se pudo exportar el archivo.\n\n{exc}")

    def _load_async(self) -> None:
        try:
            data, sheet_name = load_sheet_data(
                self.sheet_var.get(),
                filter_col_index=1,
                allowed_values=ALLOWED_NAMES,
                date_filter_col_index=9,
                date_filter_from=self.selected_date_from,
                date_filter_to=self.selected_date_to,
            )
            self.queue.put(("data", data))
            self.queue.put(("sheet", sheet_name))
            self.queue.put(("ok", None))
        except Exception as exc:
            self.queue.put(("error", exc))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "data":
                    self.base_df = payload
                    self.current_df = payload
                    self.count_var.set(f"{len(self.base_df):,}".replace(",", "."))
                    self.status_var.set("Datos cargados correctamente.")
                    self._refresh_filter_sources()
                elif kind == "sheet":
                    self.sheet_var.set(payload)
                elif kind == "ok":
                    self.last_loaded_at = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    self.refresh_var.set(self.last_loaded_at)
                    self.refresh_button.configure(state="normal", text="Actualizar datos")
                    self.date_refresh_button.configure(state="normal")
                    self._set_busy(False)
                    self._apply_visible_filter()
                elif kind == "error":
                    self.refresh_button.configure(state="normal", text="Actualizar datos")
                    self.date_refresh_button.configure(state="normal")
                    self._set_busy(False)
                    self.status_var.set("No se pudo cargar la información.")
                    messagebox.showerror("Error al consultar la hoja", str(payload))
        except Empty:
            pass
        finally:
            self.after(120, self._poll_queue)

    def _refresh_filter_sources(self) -> None:
        self._client_options = []
        self._brand_options = []
        self._line_options = []

        if not self.base_df.empty:
            if len(self.base_df.columns) > 2:
                self._client_options = unique_display_values(self.base_df.iloc[:, 2])
            if len(self.base_df.columns) > 5:
                self._brand_options = unique_display_values(self.base_df.iloc[:, 5])
            if len(self.base_df.columns) > 11:
                self._line_options = unique_display_values(self.base_df.iloc[:, 11])

        if self.client_popup is not None:
            self.client_popup.set_options(self._client_options)
        if self.brand_popup is not None:
            self.brand_popup.set_options(self._brand_options)
        if self.line_popup is not None:
            self.line_popup.set_options(self._line_options)

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.search_entry.configure(state=state)
        self.clear_button.configure(state=state)
        self.export_button.configure(state=state)
        self.date_refresh_button.configure(state=state)
        self.confidence_menu.configure(state=state)
        for spec in getattr(self, "extra_filters", []):
            widget = spec.get("widget")
            if widget is None:
                continue
            if hasattr(widget, "entry"):
                widget.entry.configure(state=state)
            else:
                widget.configure(state=state)

    def _refresh_extra_filter_sources(self) -> None:
        for spec in getattr(self, "extra_filters", []):
            if spec.get("values_source") != "data":
                continue
            widget = spec.get("widget")
            column_index = spec.get("column_index")
            if widget is None or column_index is None or self.base_df.empty:
                continue
            idx = int(column_index)
            if len(self.base_df.columns) <= idx:
                continue
            options = unique_display_values(self.base_df.iloc[:, idx])
            if hasattr(widget, "set_options"):
                widget.set_options(options)

    def _apply_visible_filter(self) -> None:
        if self.base_df.empty:
            self.visible_var.set("0")
            self._render_table(self.base_df)
            return

        df = self.base_df.copy()

        # Cada filtro se apila sobre el resultado anterior; así no se pisan entre sí.
        def apply_text_filter(frame: pd.DataFrame, column_index: int, value: str) -> pd.DataFrame:
            query = normalize_search_text(value.strip())
            if not query or len(frame.columns) <= column_index:
                return frame
            column = frame.iloc[:, column_index].map(normalize_search_text)
            return frame.loc[column.str.contains(query, na=False)].copy()

        selected_representative = normalize_text(self.representative_var.get())
        if selected_representative and selected_representative != "TODOS" and len(df.columns) > 1:
            df = df.loc[df.iloc[:, 1].map(normalize_text) == selected_representative].copy()

        df = apply_text_filter(df, 2, self.client_var.get())
        df = apply_text_filter(df, 5, self.brand_var.get())
        df = apply_text_filter(df, 11, self.line_var.get())

        selected_confidence = normalize_text(self.confidence_var.get())
        if selected_confidence and selected_confidence != "TODOS" and len(df.columns) > 10:
            df = df.loc[df.iloc[:, 10].map(normalize_text) == selected_confidence].copy()

        if len(df.columns) > 9 and (self.selected_date_from is not None or self.selected_date_to is not None):
            parsed_dates = pd.to_datetime(df.iloc[:, 9], errors="coerce", dayfirst=True).dt.date
            mask = pd.Series(True, index=df.index)
            if self.selected_date_from is not None:
                mask &= parsed_dates >= self.selected_date_from
            if self.selected_date_to is not None:
                mask &= parsed_dates <= self.selected_date_to
            df = df.loc[mask].copy()

        query = self.search_var.get().strip().lower()
        if query:
            mask = df.astype(str).apply(lambda col: col.str.lower().str.contains(query, na=False))
            df = df.loc[mask.any(axis=1)].copy()

        self.current_df = df
        self.visible_var.set(f"{len(df):,}".replace(",", "."))
        self._render_table(df)

    def _render_table(self, df: pd.DataFrame) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        if df is None or df.empty:
            self.tree["columns"] = ("mensaje",)
            self.tree.heading("mensaje", text="Sin resultados")
            self.tree.column("mensaje", width=900, anchor="center", stretch=True)
            self.tree.insert("", "end", values=("No hay registros para los filtros seleccionados.",))
            return

        columns = [str(col) for idx, col in enumerate(df.columns) if idx not in HIDDEN_COLUMN_INDICES]
        self.tree["columns"] = columns

        for col in columns:
            self.tree.heading(col, text=col, anchor="center")

        widths = self._estimate_widths(df, columns)
        for col, width in widths.items():
            self.tree.column(col, width=width, minwidth=120, anchor="center", stretch=False)

        for _, row in df.iterrows():
            values = [safe_string(row[col]) for col in columns]
            self.tree.insert("", "end", values=values)

    def _estimate_widths(self, df: pd.DataFrame, columns: Iterable[str]) -> dict[str, int]:
        widths: dict[str, int] = {}
        sample = df.head(50)
        for col in columns:
            best = len(str(col)) * 11 + 32
            if not sample.empty:
                max_len = max([len(str(col))] + [len(safe_string(v)) for v in sample[col].tolist()])
                best = min(max(140, max_len * 9 + 32), 560)
            widths[col] = best
        return widths

    def _on_row_click(self, _event: Any) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0], "values")
        if not values:
            return
        snippet = " | ".join(str(v) for v in values[:5])
        self.status_var.set(f"Fila seleccionada: {snippet}")


class ConsultaApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry("1280x800")
        self.minsize(1100, 700)
        self.configure(fg_color="#f8fafc")
        self.title(f"{APP_TITLE} | {APP_VERSION}")

        self._configure_style()
        self._build_ui()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#0f172a",
            rowheight=38,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        style.map(
            "Treeview",
            background=[("selected", "#dbeafe")],
            foreground=[("selected", "#0f172a")],
        )
        style.configure(
            "Treeview.Heading",
            background="#dbeafe",
            foreground="#1d4ed8",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        top_bar = ctk.CTkFrame(self, corner_radius=18, fg_color="#ffffff")
        top_bar.grid(row=0, column=0, sticky="ew", padx=16, pady=(8, 4))
        top_bar.grid_columnconfigure(0, weight=1)

        tab_bar = ctk.CTkFrame(top_bar, fg_color="#eef2ff", corner_radius=16)
        tab_bar.grid(row=0, column=0, sticky="w", padx=10, pady=8)
        tab_bar.grid_columnconfigure(0, weight=0)
        tab_bar.grid_columnconfigure(1, weight=0)

        self.registry_tab_button = ctk.CTkButton(
            tab_bar,
            text="Registros",
            command=lambda: self._show_panel("registros"),
            height=30,
            corner_radius=12,
        )
        self.registry_tab_button.grid(row=0, column=0, sticky="w", padx=(0, 6), pady=0)

        self.visitas_tab_button = ctk.CTkButton(
            tab_bar,
            text="Visitas",
            command=lambda: self._show_panel("visitas"),
            height=30,
            corner_radius=12,
        )
        self.visitas_tab_button.grid(row=0, column=1, sticky="w", padx=(6, 0), pady=0)

        self.content = ctk.CTkFrame(self, corner_radius=22, fg_color="#ffffff")
        self.content.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))
        self.content.grid_rowconfigure(0, weight=1)
        self.content.grid_columnconfigure(0, weight=1)

        self.registros_panel = RegistrosPanel(self.content)
        self.visitas_panel = DataPanel(
            self.content,
            panel_title="Resultados de visitas",
            worksheet_name="Visita",
            filter_label="Visitas autorizadas",
            filter_values=[
                "ERIKA UCHUYA TROCONES",
                "JAVIER KLUIVERT CONDOR SANCHEZ",
                "Ximena Jamilet Montoya Calderon",
            ],
            raw_filter_values=["70321862", "70122639", "71406087"],
            extra_filters=[
                {
                    "kind": "autocomplete",
                    "label": "Cliente",
                    "placeholder": "Escribe para buscar cliente...",
                    "column_index": 2,
                    "row": 2,
                    "column": 0,
                    "values_source": "data",
                },
                {
                    "kind": "fixed",
                    "label": "Tipo Visita",
                    "column_index": 1,
                    "row": 2,
                    "column": 1,
                    "values": [
                        "CAPACITACIÓN",
                        "PROMOCIÓN",
                        "POST VENTA",
                        "SEGUIMIENTO DE OC",
                        "TENTATIVOS",
                        "PROYECTOS",
                        "VENTA",
                        "OTROS",
                        "INICIO DE LABORES",
                    ],
                },
            ],
            filter_col_index=0,
            date_filter_label="Fecha de visita",
            date_filter_col_index=9,
            hidden_column_indices=set(),
            show_filter_menu=True,
            search_placeholder="Escribe para filtrar visitas cargadas...",
            initial_filter_value="Todos",
            loaded_df_transform=prepare_visitas_dataframe,
            load_filter_on_raw_data=True,
            compact_layout=True,
            wrap_cells=True,
            max_column_width=300,
            tree_row_height=58,
            show_actions_card=False,
            show_stats_in_controls=True,
        )
        self.registros_panel.grid(row=0, column=0, sticky="nsew")
        self.visitas_panel.grid(row=0, column=0, sticky="nsew")
        self._show_panel("registros")

    def _show_panel(self, name: str) -> None:
        if name == "registros":
            self.registros_panel.tkraise()
            self.registry_tab_button.configure(fg_color="#1d4ed8", text_color="#ffffff")
            self.visitas_tab_button.configure(fg_color="#e2e8f0", text_color="#0f172a")
        else:
            self.visitas_panel.tkraise()
            self.visitas_tab_button.configure(fg_color="#1d4ed8", text_color="#ffffff")
            self.registry_tab_button.configure(fg_color="#e2e8f0", text_color="#0f172a")

if __name__ == "__main__":
    app = ConsultaApp()
    app.mainloop()
