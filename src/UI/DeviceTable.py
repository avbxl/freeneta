import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont
from typing import Callable, Dict, List, Optional, Set, Tuple

from Dataclasses.DeviceRow import DeviceRow
from UI.AutoScrollbar import AutoScrollbar
from UI.Theme import ThemeManager


class DeviceTableView(ttk.Frame):
    """Encapsulates the device treeview table, scrollbars, column management, and autosizing."""

    COLUMNS = ("name", "mac", "vendor", "ip", "ping", "netmask", "gateway", "family")

    HEADINGS = {
        "name": "Station Name",
        "mac": "MAC",
        "vendor": "Vendor",
        "ip": "IP",
        "ping": "Ping",
        "netmask": "Netmask",
        "gateway": "Gateway",
        "family": "Family",
    }

    STRETCHABLE_COLUMNS = {"name", "vendor", "family"}

    def __init__(
        self,
        parent: tk.Widget,
        theme_manager: ThemeManager,
        on_selection_changed: Optional[Callable[[Optional[DeviceRow]], None]] = None,
        *args,
        **kwargs,
    ):
        super().__init__(parent, *args, **kwargs)
        self.theme_manager = theme_manager
        self.on_selection_changed = on_selection_changed
        self.devices: List[DeviceRow] = []

        self.user_resized_columns: Set[str] = set()
        self._column_widths_before_drag: Dict[str, int] = {}

        self.column_vars = {
            "name": tk.BooleanVar(value=True),
            "mac": tk.BooleanVar(value=True),
            "vendor": tk.BooleanVar(value=True),
            "ip": tk.BooleanVar(value=True),
            "netmask": tk.BooleanVar(value=True),
            "gateway": tk.BooleanVar(value=True),
            "family": tk.BooleanVar(value=True),
        }
        self.ping_column_enabled = False

        self._build_tree()

    def _build_tree(self) -> None:
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            self,
            columns=self.COLUMNS,
            show="headings",
            height=10,
        )

        widths = {
            "name": self.theme_manager.scaled(210),
            "mac": self.theme_manager.scaled(150),
            "vendor": self.theme_manager.scaled(220),
            "ip": self.theme_manager.scaled(110),
            "ping": self.theme_manager.scaled(110),
            "netmask": self.theme_manager.scaled(125),
            "gateway": self.theme_manager.scaled(125),
            "family": self.theme_manager.scaled(200),
        }

        for col in self.COLUMNS:
            self.tree.heading(col, text=self.HEADINGS[col])
            self.tree.column(
                col,
                width=widths[col],
                minwidth=self.theme_manager.scaled(90),
                anchor="w",
                stretch=col in self.STRETCHABLE_COLUMNS,
            )

        self.tree_scroll_y = AutoScrollbar(
            self, orient="vertical", command=self.tree.yview
        )
        self.tree_scroll_x = AutoScrollbar(
            self, orient="horizontal", command=self.tree.xview
        )
        self.tree.configure(
            yscrollcommand=self.tree_scroll_y.set,
            xscrollcommand=self.tree_scroll_x.set,
        )

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree_scroll_y.grid(row=0, column=1, sticky="ns")
        self.tree_scroll_x.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind(
            "<ButtonPress-1>", self._remember_column_widths_before_drag, add="+"
        )
        self.tree.bind("<ButtonRelease-1>", self._detect_user_column_resize, add="+")
        self.tree.bind(
            "<Double-1>", self._autosize_column_from_header_doubleclick, add="+"
        )

        self.update_tree_columns()

    def build_columns_menu(
        self, menu: tk.Menu, on_columns_changed: Optional[Callable[[], None]] = None
    ) -> None:
        ordered_columns = (
            "name",
            "mac",
            "vendor",
            "ip",
            "netmask",
            "gateway",
            "family",
        )
        for col in ordered_columns:
            menu.add_checkbutton(
                label=self.HEADINGS[col],
                variable=self.column_vars[col],
                command=lambda c=col: self._toggle_column(c, on_columns_changed),
            )

    def _toggle_column(
        self, column_key: str, callback: Optional[Callable[[], None]]
    ) -> None:
        enabled = [key for key, var in self.column_vars.items() if var.get()]
        if not enabled:
            self.column_vars[column_key].set(True)
            return
        self.update_tree_columns()
        if callback:
            callback()

    def set_ping_monitor_enabled(self, enabled: bool) -> None:
        self.ping_column_enabled = enabled
        self.update_tree_columns()

    def update_tree_columns(self) -> None:
        display_columns = []
        for col in self.COLUMNS:
            if col == "ping":
                if self.ping_column_enabled:
                    display_columns.append(col)
            elif self.column_vars.get(col, tk.BooleanVar(value=True)).get():
                display_columns.append(col)

        if not display_columns:
            display_columns = ["name"]
            self.column_vars["name"].set(True)

        self.tree.configure(displaycolumns=tuple(display_columns))
        self.after_idle(self.autosize_tree_columns)

    def _tree_display_columns(self) -> Tuple[str, ...]:
        display_cols = self.tree.cget("displaycolumns")
        if display_cols == "#all":
            return tuple(self.tree["columns"])
        if isinstance(display_cols, str):
            return tuple(col for col in display_cols.split() if col)
        return tuple(display_cols)

    def _remember_column_widths_before_drag(self, _event=None) -> None:
        self._column_widths_before_drag = {
            col: self.tree.column(col, "width") for col in self._tree_display_columns()
        }

    def _detect_user_column_resize(self, _event=None) -> None:
        if not self._column_widths_before_drag:
            return
        for col in self._tree_display_columns():
            before = self._column_widths_before_drag.get(col)
            after = self.tree.column(col, "width")
            if before is not None and after != before:
                self.user_resized_columns.add(col)
        self._column_widths_before_drag = {}

    def _column_key_from_event(self, event) -> Optional[str]:
        region = self.tree.identify_region(event.x, event.y)
        if region not in ("separator", "heading"):
            return None
        col_id = self.tree.identify_column(event.x)
        if not col_id or col_id == "#0":
            return None
        try:
            idx = int(col_id[1:]) - 1
        except Exception:
            return None
        display_cols = self._tree_display_columns()
        if 0 <= idx < len(display_cols):
            return display_cols[idx]
        return None

    def _autosize_column_from_header_doubleclick(self, event) -> None:
        col = self._column_key_from_event(event)
        if not col:
            return
        self.user_resized_columns.discard(col)
        self.autosize_tree_columns(columns=[col])

    def autosize_tree_columns(self, only_visible: bool = True, columns=None) -> None:
        if columns is None:
            display_cols = (
                self._tree_display_columns()
                if only_visible
                else tuple(self.tree["columns"])
            )
        else:
            display_cols = tuple(columns)

        if not display_cols:
            return

        heading_font = tkfont.nametofont("TkHeadingFont")
        body_font = tkfont.nametofont("TkDefaultFont")
        padding = self.theme_manager.scaled(24)
        min_col_width = self.theme_manager.scaled(90)
        max_col_width = self.theme_manager.scaled(480)
        tree_columns = tuple(self.tree["columns"])

        for col in display_cols:
            if col in self.user_resized_columns:
                continue

            header_text = self.HEADINGS.get(col, col)
            best_width = heading_font.measure(header_text) + padding

            try:
                value_idx = tree_columns.index(col)
            except ValueError:
                continue

            for item in self.tree.get_children():
                values = self.tree.item(item, "values")
                cell_text = str(values[value_idx]) if value_idx < len(values) else ""
                best_width = max(best_width, body_font.measure(cell_text) + padding)

            best_width = max(min_col_width, min(best_width, max_col_width))
            stretch = bool(self.tree.column(col, "stretch"))
            self.tree.column(
                col, width=best_width, minwidth=min_col_width, stretch=stretch
            )

    def _device_values(self, dev: DeviceRow) -> tuple:
        return (
            dev.name_of_station,
            dev.mac,
            dev.vendor,
            dev.ip,
            dev.ping_display,
            dev.netmask,
            dev.gateway,
            dev.family,
        )

    def _device_row_tag(self, dev: DeviceRow) -> str:
        access = dev.dcp_access_normalized
        if access == "read_only":
            return "dcp_readonly"
        if access == "read_write":
            return "dcp_readwrite"
        return "dcp_unknown"

    def set_devices(self, devices: List[DeviceRow]) -> None:
        self.devices = devices
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, dev in enumerate(devices):
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=self._device_values(dev),
                tags=(self._device_row_tag(dev),),
            )
        self.update_tree_columns()
        self.autosize_tree_columns()

    def update_device_row(self, idx: int) -> None:
        if 0 <= idx < len(self.devices) and self.tree.exists(str(idx)):
            dev = self.devices[idx]
            self.tree.item(
                str(idx),
                values=self._device_values(dev),
                tags=(self._device_row_tag(dev),),
            )

    def get_selected_device(self) -> Optional[DeviceRow]:
        sel = self.tree.selection()
        if not sel:
            return None
        idx = int(sel[0])
        if 0 <= idx < len(self.devices):
            return self.devices[idx]
        return None

    def get_selected_index(self) -> Optional[int]:
        sel = self.tree.selection()
        if not sel:
            return None
        return int(sel[0])

    def select_device_by_index(self, idx: int) -> None:
        if 0 <= idx < len(self.devices):
            iid = str(idx)
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)
            if self.on_selection_changed:
                self.on_selection_changed(self.devices[idx])

    def _on_tree_select(self, _event=None) -> None:
        if self.on_selection_changed:
            self.on_selection_changed(self.get_selected_device())
