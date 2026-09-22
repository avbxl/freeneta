import tkinter as tk
from tkinter import ttk
from typing import Callable, Dict, List, Optional, Tuple

from Dataclasses.DeviceRow import DeviceRow
from Services.Constants import APP_NOTES
from UI.DeviceTable import DeviceTableView
from UI.HorizontalScrollableFrame import HorizontalScrollableFrame
from UI.ScrollableFrame import ScrollableFrame
from UI.Theme import ThemeManager
from UI.TopologyView import TopologyView


class MainWindow:
    """FreeNeta Main Window presentation layout and user interface controls."""

    def __init__(
        self,
        root: tk.Tk,
        theme_manager: ThemeManager,
        callbacks: Dict[str, Callable],
    ):
        self.root = root
        self.theme_manager = theme_manager
        self.callbacks = callbacks

        self.root.title("FreeNeta")
        try:
            self.root.iconbitmap("app.ico")
        except Exception:
            pass
        self.root.geometry(self.theme_manager.scaled_geometry(1440, 820))
        self.root.minsize(
            self.theme_manager.scaled(1100), self.theme_manager.scaled(700)
        )

        self.dark_mode_var = tk.BooleanVar(value=False)
        self.ping_monitor_var = tk.BooleanVar(value=False)
        self.show_topology_var = tk.BooleanVar(value=True)
        self.show_notes_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Idle.")
        self.host_interface_var = tk.StringVar()
        self.host_ip_var = tk.StringVar()

        self.right_panel_visible = True
        self.last_sash_fraction = 0.64
        self.quick_actions = []

        self._build_layout()

    def _build_layout(self) -> None:
        self.outer = ScrollableFrame(self.root)
        self.outer.pack(fill="both", expand=True)

        main = self.outer.content
        main.configure(padding=self.theme_manager.scaled(16))
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        # -------------------------------------------------------------
        # Top toolbar
        # -------------------------------------------------------------
        top = ttk.Frame(main)
        top.grid(row=0, column=0, sticky="ew", pady=(0, self.theme_manager.scaled(14)))
        top.grid_columnconfigure(0, weight=1)
        top.grid_columnconfigure(1, weight=0)

        self.top_scroller = HorizontalScrollableFrame(top)
        self.top_scroller.grid(row=0, column=0, sticky="ew")
        top_bar = self.top_scroller.content

        ttk.Label(top_bar, text="Host interface").grid(
            row=0, column=0, sticky="w", pady=4
        )

        self.interface_combo = ttk.Combobox(
            top_bar,
            textvariable=self.host_interface_var,
            state="readonly",
            width=34,
        )
        self.interface_combo.grid(row=0, column=1, sticky="w", padx=(8, 8), pady=4)
        self.interface_combo.bind(
            "<<ComboboxSelected>>",
            lambda _e: self.callbacks.get("on_interface_selected", lambda: None)(),
        )

        self.refresh_interfaces_btn = ttk.Button(
            top_bar,
            text="Refresh interfaces",
            command=self.callbacks.get("refresh_interfaces", lambda: None),
        )
        self.refresh_interfaces_btn.grid(
            row=0, column=2, sticky="w", padx=(0, 16), pady=4
        )

        self.scan_btn = ttk.Button(
            top_bar,
            text="Refresh Devices",
            command=self.callbacks.get("scan_devices", lambda: None),
        )
        self.scan_btn.grid(row=0, column=3, sticky="w", pady=4)

        self.set_ip_btn = ttk.Button(
            top_bar,
            text="Set IP",
            state="disabled",
            command=self.callbacks.get("set_ip", lambda: None),
        )
        self.set_ip_btn.grid(row=0, column=4, sticky="w", padx=(18, 0), pady=4)

        self.set_name_btn = ttk.Button(
            top_bar,
            text="Set Name",
            state="disabled",
            command=self.callbacks.get("set_name", lambda: None),
        )
        self.set_name_btn.grid(row=0, column=5, sticky="w", padx=(8, 0), pady=4)

        self.reset_btn = ttk.Button(
            top_bar,
            text="Reset to Factory",
            state="disabled",
            command=self.callbacks.get("reset_comm", lambda: None),
        )
        self.reset_btn.grid(row=0, column=6, sticky="w", padx=(8, 0), pady=4)

        self.monitor_chk = ttk.Checkbutton(
            top_bar,
            text="Ping monitor",
            variable=self.ping_monitor_var,
            command=self._on_toggle_ping_monitor,
        )
        self.monitor_chk.grid(row=0, column=7, sticky="w", padx=(18, 0), pady=4)

        self.view_button = ttk.Menubutton(top_bar, text="View")
        self.view_button.grid(row=0, column=8, sticky="w", padx=(18, 0), pady=4)

        self.view_menu = tk.Menu(self.view_button, tearoff=False)
        self.view_menu.add_checkbutton(
            label="Show topology",
            variable=self.show_topology_var,
            command=self.update_view_visibility,
        )
        self.view_menu.add_checkbutton(
            label="Show notes",
            variable=self.show_notes_var,
            command=self.update_view_visibility,
        )
        self.view_menu.add_separator()
        self.columns_menu = tk.Menu(self.view_menu, tearoff=False)
        self.view_menu.add_cascade(label="Columns", menu=self.columns_menu)
        self.view_menu.add_separator()
        self.view_menu.add_checkbutton(
            label="Dark mode",
            variable=self.dark_mode_var,
            command=self._on_toggle_dark_mode,
        )
        self.view_button["menu"] = self.view_menu

        ttk.Label(top, textvariable=self.status_var).grid(
            row=0, column=1, sticky="e", padx=(12, 0), pady=4
        )

        # -------------------------------------------------------------
        # Body PanedWindow
        # -------------------------------------------------------------
        body = ttk.PanedWindow(main, orient="horizontal")
        body.grid(row=1, column=0, sticky="nsew")
        self.body_pane = body

        self.left_panel = ttk.Frame(body, padding=(0, 0, 10, 0))
        self.right_panel = ttk.Frame(body)
        self.left_panel.grid_rowconfigure(0, weight=1)
        self.left_panel.grid_columnconfigure(0, weight=1)
        self.right_panel.grid_rowconfigure(1, weight=3)
        self.right_panel.grid_rowconfigure(4, weight=1)
        self.right_panel.grid_columnconfigure(0, weight=1)
        body.add(self.left_panel, weight=4)
        body.add(self.right_panel, weight=3)

        # -------------------------------------------------------------
        # Left Panel: Device Table & Action Row
        # -------------------------------------------------------------
        self.table_view = DeviceTableView(
            self.left_panel,
            theme_manager=self.theme_manager,
            on_selection_changed=self._on_table_selection_changed,
        )
        self.table_view.grid(row=0, column=0, sticky="nsew")
        self.table_view.build_columns_menu(
            self.columns_menu,
            on_columns_changed=self._on_columns_changed,
        )

        action_row = ttk.Frame(self.left_panel)
        action_row.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(
            action_row,
            text="Export CSV",
            command=self.callbacks.get("export_csv", lambda: None),
        ).pack(side="left")
        self.show_details_btn = ttk.Button(
            action_row,
            text="Show Selected Details",
            state="disabled",
            command=self.callbacks.get("show_details", lambda: None),
        )
        self.show_details_btn.pack(side="left", padx=(8, 0))

        self.quick_menu_button = ttk.Menubutton(
            action_row, text="Quick connect", state="disabled"
        )
        self.quick_menu_button.pack(side="left", padx=(8, 0))
        self.quick_menu = tk.Menu(self.quick_menu_button, tearoff=False)
        self.quick_menu_button["menu"] = self.quick_menu

        self.blink_btn = ttk.Button(
            action_row,
            text="Blink Device",
            state="disabled",
            command=self.callbacks.get("blink", lambda: None),
        )
        self.blink_btn.pack(side="left", padx=(8, 0))

        # -------------------------------------------------------------
        # Right Panel: Topology & Notes
        # -------------------------------------------------------------
        self.topology_title = ttk.Label(
            self.right_panel, text="Topology View", font="TkHeadingFont"
        )
        self.topology_title.grid(row=0, column=0, sticky="w")

        self.topology_view = TopologyView(
            self.right_panel,
            theme_manager=self.theme_manager,
            on_device_clicked=self._on_topology_node_clicked,
        )
        self.topology_view.grid(row=1, column=0, sticky="nsew", pady=(6, 0))

        self.notes_title = ttk.Label(
            self.right_panel, text="Notes", font="TkHeadingFont"
        )
        self.notes_title.grid(row=3, column=0, sticky="w", pady=(14, 6))
        self.notes = tk.Text(
            self.right_panel,
            height=8,
            wrap="word",
            relief="solid",
            borderwidth=1,
            font="TkTextFont",
            padx=self.theme_manager.scaled(8),
            pady=self.theme_manager.scaled(8),
            spacing1=self.theme_manager.scaled(2),
            spacing3=self.theme_manager.scaled(2),
        )
        self.notes.insert("1.0", APP_NOTES)
        self.notes.configure(state="disabled")
        self.notes.grid(row=4, column=0, sticky="nsew")

        # Window resize and sash position handling
        self.root.bind("<Configure>", self._on_root_resize, add="+")
        self.body_pane.bind(
            "<ButtonRelease-1>", lambda _e: self._save_current_sash_fraction(), add="+"
        )
        self.root.after_idle(self.apply_initial_layout)

    def apply_initial_layout(self) -> None:
        self.show_topology_var.set(True)
        self.show_notes_var.set(True)
        self._ensure_right_panel_visible()
        self.update_view_visibility()
        self.top_scroller._update_scrollbar_visibility()
        self._restore_sash_fraction()
        self.outer._update_scrollbar_visibility()

    def _on_root_resize(self, event=None) -> None:
        if event is not None and event.widget is not self.root:
            return
        if hasattr(self, "outer"):
            self.outer.after_idle(self.outer._update_scrollbar_visibility)
        if hasattr(self, "top_scroller"):
            self.top_scroller.after_idle(self.top_scroller._update_scrollbar_visibility)
        if self.show_topology_var.get():
            self.root.after_idle(self.topology_view.draw)

    def _on_columns_changed(self) -> None:
        if hasattr(self, "top_scroller"):
            self.top_scroller.after_idle(self.top_scroller._update_scrollbar_visibility)

    def set_device_selected_state(self, is_selected: bool) -> None:
        """Enables or disables action buttons based on whether a device is selected."""
        button_state = "normal" if is_selected else "disabled"
        self.set_ip_btn.configure(state=button_state)
        self.set_name_btn.configure(state=button_state)
        self.reset_btn.configure(state=button_state)
        self.blink_btn.configure(state=button_state)
        if hasattr(self, "show_details_btn"):
            self.show_details_btn.configure(state=button_state)

    def _on_table_selection_changed(self, device: Optional[DeviceRow]) -> None:
        self.set_device_selected_state(device is not None)
        if self.show_topology_var.get():
            self.topology_view.set_selected_index(self.table_view.get_selected_index())
        on_select = self.callbacks.get("on_device_selected")
        if on_select:
            on_select(device)

    def _on_topology_node_clicked(self, index: int) -> None:
        self.table_view.select_device_by_index(index)

    def _on_toggle_ping_monitor(self) -> None:
        enabled = self.ping_monitor_var.get()
        self.table_view.set_ping_monitor_enabled(enabled)
        callback = self.callbacks.get("on_toggle_ping_monitor")
        if callback:
            callback(enabled)

    def _on_toggle_dark_mode(self) -> None:
        self.apply_theme()
        self.update_view_visibility()

    def update_view_visibility(self) -> None:
        topology_visible = self.show_topology_var.get()
        notes_visible = self.show_notes_var.get()
        right_should_show = topology_visible or notes_visible

        if topology_visible:
            self.topology_title.grid()
            self.topology_view.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
            self.right_panel.grid_rowconfigure(1, weight=3)
            self.root.after_idle(self.topology_view.draw)
        else:
            self.topology_title.grid_remove()
            self.topology_view.grid_remove()
            self.right_panel.grid_rowconfigure(1, weight=0)

        if notes_visible:
            self.notes_title.grid()
            self.notes.grid()
            self.right_panel.grid_rowconfigure(4, weight=1)
        else:
            self.notes_title.grid_remove()
            self.notes.grid_remove()
            self.right_panel.grid_rowconfigure(4, weight=0)

        if right_should_show:
            self._ensure_right_panel_visible()
        else:
            self._hide_right_panel()

        if topology_visible and notes_visible:
            self.status_var.set("Topology and notes shown.")
        elif topology_visible:
            self.status_var.set("Topology shown. Notes hidden.")
        elif notes_visible:
            self.status_var.set("Notes shown. Topology hidden.")
        else:
            self.status_var.set("Topology and notes hidden.")

        if hasattr(self, "outer"):
            self.outer.after_idle(self.outer._update_scrollbar_visibility)

    def _save_current_sash_fraction(self) -> None:
        if not getattr(self, "right_panel_visible", False):
            return
        try:
            panes = tuple(self.body_pane.panes())
            if len(panes) < 2:
                return
            total_width = max(self.body_pane.winfo_width(), 1)
            sash_x = self.body_pane.sashpos(0)
            self.last_sash_fraction = min(max(sash_x / total_width, 0.25), 0.85)
        except tk.TclError:
            pass

    def _restore_sash_fraction(self) -> None:
        if not getattr(self, "right_panel_visible", False):
            return
        try:
            total_width = max(self.body_pane.winfo_width(), 1)
            sash_x = int(total_width * self.last_sash_fraction)
            self.body_pane.sashpos(0, sash_x)
        except tk.TclError:
            pass

    def _hide_right_panel(self) -> None:
        if not getattr(self, "right_panel_visible", False):
            return
        self._save_current_sash_fraction()
        try:
            self.body_pane.forget(self.right_panel)
        except tk.TclError:
            pass
        self.right_panel_visible = False

    def _ensure_right_panel_visible(self) -> None:
        if getattr(self, "right_panel_visible", False):
            self.root.after_idle(self._restore_sash_fraction)
            return
        try:
            self.body_pane.add(self.right_panel, weight=2)
        except tk.TclError:
            return
        self.right_panel_visible = True
        self.root.after_idle(self._restore_sash_fraction)

    def apply_theme(self) -> None:
        colors = self.theme_manager.apply_theme(
            dark_mode=self.dark_mode_var.get(),
            tree=self.table_view.tree,
        )
        c = colors

        if hasattr(self, "outer"):
            self.outer.configure(style="TFrame")
            self.outer.canvas.configure(bg=c["bg"])
            self.outer.content.configure(style="TFrame")
        if hasattr(self, "top_scroller"):
            self.top_scroller.configure(style="TFrame")
            self.top_scroller.canvas.configure(bg=c["bg"])
            self.top_scroller.content.configure(style="TFrame")

        self.topology_view.set_colors(bg=c["canvas_bg"], border=c["canvas_border"])
        self.notes.configure(
            bg=c["note_bg"],
            fg=c["text"],
            insertbackground=c["text"],
            highlightbackground=c["note_border"],
            highlightcolor=c["note_border"],
        )
        if self.show_topology_var.get():
            self.topology_view.draw()

    def set_quick_actions(
        self, actions: List[Tuple[str, Callable]], message: str = "Quick connect"
    ) -> None:
        self.quick_actions = actions
        self.quick_menu.delete(0, "end")

        if actions:
            for label, command in actions:
                self.quick_menu.add_command(label=label, command=command)
            self.quick_menu_button.configure(
                text=f"Quick connect ({len(actions)})", state="normal"
            )
        else:
            self.quick_menu.add_command(label=message, state="disabled")
            self.quick_menu_button.configure(text=message, state="disabled")
