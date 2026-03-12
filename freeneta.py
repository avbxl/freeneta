import json
import platform
import shutil
import socket
import subprocess
import threading
import time
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
import socket
import psutil
from dataclasses import dataclass, field
from tkinter import ttk, messagebox, simpledialog
from typing import Dict, List, Optional

try:
    from pnio_dcp import DCP
except Exception:
    DCP = None


class AutoScrollbar(ttk.Scrollbar):
    def set(self, first, last):
        first = float(first)
        last = float(last)
        if first <= 0.0 and last >= 1.0:
            if self.winfo_ismapped():
                self.grid_remove()
        else:
            if not self.winfo_ismapped():
                self.grid()
        super().set(first, last)


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.v_scrollbar = AutoScrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set)

        self.content = ttk.Frame(self.canvas)
        self.window_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.v_scrollbar.grid_remove()
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.content.bind("<Configure>", self._on_content_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.canvas.bind_all("<Button-4>", self._on_linux_scroll_up, add="+")
        self.canvas.bind_all("<Button-5>", self._on_linux_scroll_down, add="+")

    def _on_content_configure(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.after_idle(self._update_scrollbar_visibility)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)
        self.update_idletasks()
        content_height = self.content.winfo_reqheight()
        self.canvas.itemconfigure(self.window_id, height=max(event.height, content_height))
        self.after_idle(self._update_scrollbar_visibility)

    def _update_scrollbar_visibility(self):
        bbox = self.canvas.bbox("all")
        if not bbox:
            self.v_scrollbar.grid_remove()
            return
        _, _, _, content_height = bbox
        canvas_height = max(self.canvas.winfo_height(), 1)
        if content_height <= canvas_height:
            self.v_scrollbar.grid_remove()
            self.canvas.yview_moveto(0)
        else:
            self.v_scrollbar.grid()

    def _pointer_inside(self):
        widget = self.winfo_containing(self.winfo_pointerx(), self.winfo_pointery())
        while widget is not None:
            if widget == self.canvas:
                return True
            widget = widget.master
        return False

    def _on_mousewheel(self, event):
        if not self._pointer_inside():
            return
        if event.delta:
            self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _on_linux_scroll_up(self, _event):
        if self._pointer_inside():
            self.canvas.yview_scroll(-1, "units")

    def _on_linux_scroll_down(self, _event):
        if self._pointer_inside():
            self.canvas.yview_scroll(1, "units")


class HorizontalScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, height=42)
        self.h_scrollbar = AutoScrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set)

        self.content = ttk.Frame(self.canvas)
        self.window_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")

        self.canvas.grid(row=0, column=0, sticky="ew")
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.h_scrollbar.grid_remove()
        self.grid_columnconfigure(0, weight=1)

        self.content.bind("<Configure>", self._on_content_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel, add="+")

    def _on_content_configure(self, _event=None):
        req_width = max(self.content.winfo_reqwidth(), 1)
        self.canvas.configure(scrollregion=(0, 0, req_width, self.canvas.winfo_height()))
        self.after_idle(self._update_scrollbar_visibility)

    def _on_canvas_configure(self, event):
        req_width = max(self.content.winfo_reqwidth(), 1)
        overflow = req_width > max(event.width, 1)
        self.canvas.itemconfigure(self.window_id, height=event.height)
        self.canvas.itemconfigure(self.window_id, width=req_width if overflow else event.width)
        self.canvas.configure(scrollregion=(0, 0, req_width, event.height))
        self.after_idle(self._update_scrollbar_visibility)

    def _update_scrollbar_visibility(self):
        req_width = max(self.content.winfo_reqwidth(), 1)
        canvas_width = max(self.canvas.winfo_width(), 1)
        overflow_threshold = 8
        if req_width <= canvas_width + overflow_threshold:
            self.h_scrollbar.grid_remove()
            self.canvas.xview_moveto(0)
            self.canvas.itemconfigure(self.window_id, width=canvas_width)
            self.canvas.configure(scrollregion=(0, 0, canvas_width, self.canvas.winfo_height()))
        else:
            self.h_scrollbar.grid()
            self.canvas.itemconfigure(self.window_id, width=req_width)
            self.canvas.configure(scrollregion=(0, 0, req_width, self.canvas.winfo_height()))

    def _pointer_inside(self):
        widget = self.winfo_containing(self.winfo_pointerx(), self.winfo_pointery())
        while widget is not None:
            if widget == self.canvas:
                return True
            widget = widget.master
        return False

    def _on_shift_mousewheel(self, event):
        if self._pointer_inside() and event.delta:
            self.canvas.xview_scroll(int(-event.delta / 120), "units")


@dataclass
class DeviceRow:
    name_of_station: str
    mac: str
    ip: str
    netmask: str
    gateway: str
    family: str
    vendor: str = "Looking up..."
    ping_status: str = "Unknown"
    ping_ms: str = ""


class Freeneta:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("FreeNeta")
        self.root.geometry("1380x760")

        self.devices: List[DeviceRow] = []
        self.dark_mode_var = tk.BooleanVar(value=False)
        self.ping_monitor_var = tk.BooleanVar(value=False)
        self.canvas_item_to_index = {}
        self.colors = {}
        self.vendor_cache: Dict[str, str] = {}
        self.ping_thread: Optional[threading.Thread] = None
        self.ping_monitor_stop = threading.Event()
        self.vendor_lookup_lock = threading.Lock()
        self.port_scan_token = 0
        self.quick_actions = []
        self.quick_menu_button = None
        self.quick_menu = None
        self.show_topology_var = tk.BooleanVar(value=True)
        self.show_notes_var = tk.BooleanVar(value=True)
        self.column_vars = {
            "name": tk.BooleanVar(value=True),
            "mac": tk.BooleanVar(value=True),
            "vendor": tk.BooleanVar(value=True),
            "ip": tk.BooleanVar(value=True),
            "netmask": tk.BooleanVar(value=True),
            "gateway": tk.BooleanVar(value=True),
            "family": tk.BooleanVar(value=True),
        }

        self._build_ui()
        self.apply_theme()

    def _build_ui(self) -> None:
        self.root.minsize(980, 620)

        self.outer = ScrollableFrame(self.root)
        self.outer.pack(fill="both", expand=True)

        main = self.outer.content
        main.configure(padding=12)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        top = ttk.Frame(main)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        top.grid_columnconfigure(0, weight=1)
        top.grid_columnconfigure(1, weight=0)

        self.top_scroller = HorizontalScrollableFrame(top)
        self.top_scroller.grid(row=0, column=0, sticky="ew")
        top_bar = self.top_scroller.content

        ttk.Label(top_bar, text="Host interface").grid(row=0, column=0, sticky="w")

        self.host_ip_var = tk.StringVar()
        self.host_interface_var = tk.StringVar()
        self.host_interfaces = self.get_host_interfaces()

        interface_values = [f"{iface} ({ip})" for iface, ip in self.host_interfaces]

        self.interface_combo = ttk.Combobox(
            top_bar,
            textvariable=self.host_interface_var,
            values=interface_values,
            state="readonly",
            width=34,
        )
        self.interface_combo.grid(row=0, column=1, sticky="w", padx=(8, 6))
        self.interface_combo.bind("<<ComboboxSelected>>", self.on_interface_selected)

        self.refresh_interfaces_btn = ttk.Button(top_bar, text="Refresh interfaces", command=self.refresh_interfaces_only)
        self.refresh_interfaces_btn.grid(row=0, column=2, sticky="w", padx=(0, 14))

        self.refresh_host_interfaces(preserve_selection=False)

        self.scan_btn = ttk.Button(top_bar, text="Scan", command=self.scan_devices)
        self.scan_btn.grid(row=0, column=3, sticky="w")

        self.refresh_btn = ttk.Button(top_bar, text="Refresh", command=self.scan_devices)
        self.refresh_btn.grid(row=0, column=4, sticky="w", padx=(8, 0))

        self.set_ip_btn = ttk.Button(top_bar, text="Set IP", command=self.set_ip_for_selected)
        self.set_ip_btn.grid(row=0, column=5, sticky="w", padx=(18, 0))

        self.set_name_btn = ttk.Button(top_bar, text="Set Name", command=self.set_name_for_selected)
        self.set_name_btn.grid(row=0, column=6, sticky="w", padx=(8, 0))

        self.reset_btn = ttk.Button(top_bar, text="Reset Comm", command=self.reset_selected)
        self.reset_btn.grid(row=0, column=7, sticky="w", padx=(8, 0))

        self.monitor_chk = ttk.Checkbutton(
            top_bar,
            text="Ping monitor",
            variable=self.ping_monitor_var,
            command=self.toggle_ping_monitor,
        )
        self.monitor_chk.grid(row=0, column=8, sticky="w", padx=(18, 0))

        self.view_button = ttk.Menubutton(top_bar, text="View")
        self.view_button.grid(row=0, column=9, sticky="w", padx=(18, 0))
        self.view_menu = tk.Menu(self.view_button, tearoff=False)
        self.view_menu.add_checkbutton(label="Show topology", variable=self.show_topology_var, command=self.update_view_visibility)
        self.view_menu.add_checkbutton(label="Show notes", variable=self.show_notes_var, command=self.update_view_visibility)
        self.view_menu.add_separator()
        self.columns_menu = tk.Menu(self.view_menu, tearoff=False)
        self.view_menu.add_cascade(label="Columns", menu=self.columns_menu)
        self.view_menu.add_separator()
        self.view_menu.add_checkbutton(label="Dark mode", variable=self.dark_mode_var, command=self.toggle_dark_mode)
        self.view_button["menu"] = self.view_menu

        self.status_var = tk.StringVar(value="Idle.")
        ttk.Label(top, textvariable=self.status_var).grid(row=0, column=1, sticky="e", padx=(12, 0))

        body = ttk.PanedWindow(main, orient="horizontal")
        body.grid(row=1, column=0, sticky="nsew")
        self.body_pane = body
        self.right_panel_visible = True
        self.last_sash_fraction = 0.60

        left = ttk.Frame(body, padding=(0, 0, 8, 0))
        right = ttk.Frame(body)
        left.grid_rowconfigure(0, weight=1)
        left.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(2, weight=3)
        right.grid_rowconfigure(4, weight=1)
        right.grid_columnconfigure(0, weight=1)
        body.add(left, weight=3)
        body.add(right, weight=2)

        columns = ("name", "mac", "vendor", "ip", "ping", "netmask", "gateway", "family")
        tree_wrap = ttk.Frame(left)
        tree_wrap.grid(row=0, column=0, sticky="nsew")

        self.tree = ttk.Treeview(tree_wrap, columns=columns, show="headings", height=10)
        headings = {
            "name": "Station Name",
            "mac": "MAC",
            "vendor": "Vendor",
            "ip": "IP",
            "ping": "Ping",
            "netmask": "Netmask",
            "gateway": "Gateway",
            "family": "Family",
        }
        self.column_labels = headings.copy()
        widths = {
            "name": 180,
            "mac": 150,
            "vendor": 180,
            "ip": 110,
            "ping": 95,
            "netmask": 110,
            "gateway": 110,
            "family": 170,
        }
        stretchable_columns = {"name", "vendor", "family"}
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], minwidth=90, anchor="w", stretch=col in stretchable_columns)

        self.tree_scroll_y = AutoScrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        self.tree_scroll_x = AutoScrollbar(tree_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree_scroll_y.grid(row=0, column=1, sticky="ns")
        self.tree_scroll_x.grid(row=1, column=0, sticky="ew")
        tree_wrap.grid_rowconfigure(0, weight=1)
        tree_wrap.grid_columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_selection_changed)

        action_row = ttk.Frame(left)
        action_row.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(action_row, text="Export CSV", command=self.export_csv).pack(side="left")
        ttk.Button(action_row, text="Show Selected Details", command=self.show_selected_details).pack(side="left", padx=(8, 0))

        self.quick_menu_button = ttk.Menubutton(action_row, text="Quick connect", state="disabled")
        self.quick_menu_button.pack(side="left", padx=(8, 0))
        self.quick_menu = tk.Menu(self.quick_menu_button, tearoff=False)
        self.quick_menu_button["menu"] = self.quick_menu

        self.left_panel = left
        self.right_panel = right

        self.topology_title = ttk.Label(right, text="Topology View", font=("Segoe UI", 12, "bold"))
        self.topology_title.grid(row=0, column=0, sticky="w")
        self.topology_desc = ttk.Label(
            right,
            text="This is a visual summary, not real cable topology. DCP does discovery and commissioning; it does not know physical links.",
            wraplength=360,
            justify="left",
        )
        self.topology_desc.grid(row=1, column=0, sticky="ew", pady=(4, 8))

        self.canvas = tk.Canvas(right, highlightthickness=1, cursor="hand2", height=260)
        self.canvas.grid(row=2, column=0, sticky="nsew")

        self.notes_title = ttk.Label(right, text="Notes", font=("Segoe UI", 11, "bold"))
        self.notes_title.grid(row=3, column=0, sticky="w", pady=(12, 4))
        self.notes = tk.Text(right, height=7, wrap="word", relief="solid", borderwidth=1)
        self.notes.insert(
            "1.0",
            "Freeneta – v1.2\n\n"
            "Features:\n"
            "- Discover PROFINET devices using DCP\n"
            "- Show station name, MAC, vendor, IP, subnet, and gateway\n"
            "- Identify device vendor via MAC OUI lookup\n"
            "- Optional ping monitor to check device reachability\n"
            "- Visual topology summary of discovered devices\n"
            "- Set device IP address\n"
            "- Set station name\n"
            "- Reset communication parameters\n"
            "- Quick-connect actions based on detected ports\n"
            "- Open device Web GUI via HTTP\n"
            "- Open device Web GUI via HTTPS\n"
            "- Start an SSH session to the device\n\n"
            "Vendor lookup note:\n"
            "- Uses local cache and an online OUI lookup fallback\n"
            "- If the machine has no internet connection, vendor may remain Unknown\n\n"
        )
        self.notes.configure(state="disabled")
        self.notes.grid(row=4, column=0, sticky="nsew")

        self._build_columns_menu()
        self._update_tree_columns()
        self.update_view_visibility()
        self.root.bind("<Configure>", self._on_root_resize, add="+")
        self.body_pane.bind("<ButtonRelease-1>", lambda _e: self._save_current_sash_fraction(), add="+")
        self.root.after_idle(self.top_scroller._update_scrollbar_visibility)
        self.root.after_idle(self._restore_sash_fraction)
        self.root.after_idle(self.outer._update_scrollbar_visibility)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _on_root_resize(self, event=None) -> None:
        if event is not None and event.widget is not self.root:
            return
        if hasattr(self, "outer"):
            self.outer.after_idle(self.outer._update_scrollbar_visibility)
        if hasattr(self, "top_scroller"):
            self.top_scroller.after_idle(self.top_scroller._update_scrollbar_visibility)
        if self.show_topology_var.get():
            self.root.after_idle(self.draw_topology)

    def update_view_visibility(self) -> None:
        topology_visible = self.show_topology_var.get()
        notes_visible = self.show_notes_var.get()
        right_should_show = topology_visible or notes_visible

        if topology_visible:
            self.topology_title.grid()
            self.topology_desc.grid()
            self.canvas.grid()
            self.topology_desc.configure(wraplength=max(self.right_panel.winfo_width() - 20, 220))
            self.right_panel.grid_rowconfigure(2, weight=3)
            self.root.after_idle(self.draw_topology)
        else:
            self.topology_title.grid_remove()
            self.topology_desc.grid_remove()
            self.canvas.grid_remove()
            self.right_panel.grid_rowconfigure(2, weight=0)

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

    def _theme_palette(self):
        if self.dark_mode_var.get():
            return {
                "bg": "#111827",
                "panel": "#1f2937",
                "text": "#f3f4f6",
                "muted": "#9ca3af",
                "canvas_bg": "#0f172a",
                "canvas_border": "#334155",
                "pc_fill": "#1d4ed8",
                "pc_outline": "#60a5fa",
                "pc_text": "#eff6ff",
                "node_fill": "#0f766e",
                "node_outline": "#5eead4",
                "node_selected_fill": "#166534",
                "node_selected_outline": "#86efac",
                "line": "#64748b",
                "note_bg": "#111827",
                "note_border": "#374151",
            }
        return {
            "bg": "#f8fafc",
            "panel": "#ffffff",
            "text": "#111827",
            "muted": "#555555",
            "canvas_bg": "#ffffff",
            "canvas_border": "#cccccc",
            "pc_fill": "#eef2ff",
            "pc_outline": "#8aa0ff",
            "pc_text": "#111827",
            "node_fill": "#ecfeff",
            "node_outline": "#67e8f9",
            "node_selected_fill": "#dcfce7",
            "node_selected_outline": "#22c55e",
            "line": "#888888",
            "note_bg": "#ffffff",
            "note_border": "#d1d5db",
        }

    def apply_theme(self) -> None:
        self.colors = self._theme_palette()
        c = self.colors

        try:
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("TFrame", background=c["bg"])
            style.configure("TPanedwindow", background=c["bg"])
            style.configure("TLabel", background=c["bg"], foreground=c["text"])
            style.configure("TButton", padding=6)
            style.configure("TMenubutton", padding=6, background=c["panel"], foreground=c["text"])
            style.map("TMenubutton", background=[("active", c["panel"])], foreground=[("active", c["text"])])
            style.configure("TCheckbutton", background=c["bg"], foreground=c["text"])
            style.map("TCheckbutton", background=[("active", c["bg"])], foreground=[("active", c["text"])])
            style.configure(
                "Treeview",
                background=c["panel"],
                foreground=c["text"],
                fieldbackground=c["panel"],
                rowheight=24,
                bordercolor=c["canvas_border"],
                lightcolor=c["canvas_border"],
                darkcolor=c["canvas_border"],
            )
            style.configure(
                "Treeview.Heading",
                background=c["panel"],
                foreground=c["text"],
                relief="flat",
            )
            style.map(
                "Treeview",
                background=[("selected", "#2563eb" if self.dark_mode_var.get() else "#bfdbfe")],
                foreground=[("selected", "#ffffff" if self.dark_mode_var.get() else "#111827")],
            )
        except Exception:
            pass

        self.root.configure(bg=c["bg"])
        if hasattr(self, "outer"):
            self.outer.configure(style="TFrame")
            self.outer.canvas.configure(bg=c["bg"])
            self.outer.content.configure(style="TFrame")
        if hasattr(self, "top_scroller"):
            self.top_scroller.configure(style="TFrame")
            self.top_scroller.canvas.configure(bg=c["bg"])
            self.top_scroller.content.configure(style="TFrame")
        self.canvas.configure(bg=c["canvas_bg"], highlightbackground=c["canvas_border"])
        self.notes.configure(bg=c["note_bg"], fg=c["text"], insertbackground=c["text"], highlightbackground=c["note_border"])
        if self.show_topology_var.get():
            self.draw_topology()

    def toggle_dark_mode(self) -> None:
        self.apply_theme()
        self.update_view_visibility()

    def _build_columns_menu(self) -> None:
        ordered_columns = ("name", "mac", "vendor", "ip", "netmask", "gateway", "family")
        for col in ordered_columns:
            self.columns_menu.add_checkbutton(
                label=self.column_labels[col],
                variable=self.column_vars[col],
                command=lambda c=col: self.toggle_column(c),
            )

    def toggle_column(self, column_key: str) -> None:
        enabled = [key for key, var in self.column_vars.items() if var.get()]
        if not enabled:
            self.column_vars[column_key].set(True)
            self.status_var.set("At least one column must stay visible.")
            return
        self._update_tree_columns()

    def _update_tree_columns(self) -> None:
        ordered_columns = ("name", "mac", "vendor", "ip", "ping", "netmask", "gateway", "family")
        display_columns = []
        for col in ordered_columns:
            if col == "ping":
                if self.ping_monitor_var.get():
                    display_columns.append(col)
            elif self.column_vars[col].get():
                display_columns.append(col)
        if not display_columns:
            display_columns = ["name"]
            self.column_vars["name"].set(True)
        self.tree.configure(displaycolumns=tuple(display_columns))
        if hasattr(self, "top_scroller"):
            self.top_scroller.after_idle(self.top_scroller._update_scrollbar_visibility)

    def get_host_interfaces(self):
        interfaces = []
        addrs = psutil.net_if_addrs()

        for iface_name, iface_addrs in addrs.items():
            for addr in iface_addrs:
                if addr.family == socket.AF_INET:
                    ip = addr.address
                    if ip and not ip.startswith("127."):
                        interfaces.append((iface_name, ip))

        def sort_key(item):
            name = item[0].lower()
            ethernet_score = 0 if any(x in name for x in ["ethernet", "eth", "en"]) else 1
            virtual_score = 1 if any(x in name for x in ["vmware", "virtual", "vbox", "hyper-v", "loopback", "bluetooth", "wlan", "wi-fi"]) else 0
            return (ethernet_score, virtual_score, name)

        interfaces.sort(key=sort_key)
        return interfaces

    def on_interface_selected(self, event=None):
        selected = self.host_interface_var.get()
        for iface_name, ip in self.host_interfaces:
            label = f"{iface_name} ({ip})"
            if selected == label:
                self.host_ip_var.set(ip)
                break

    def refresh_host_interfaces(self, preserve_selection: bool = True) -> None:
        previous_selection = self.host_interface_var.get() if preserve_selection else ""
        previous_iface_name = previous_selection.split(" (", 1)[0] if previous_selection else ""
        previous_ip = self.host_ip_var.get().strip()

        self.host_interfaces = self.get_host_interfaces()
        interface_values = [f"{iface} ({ip})" for iface, ip in self.host_interfaces]
        self.interface_combo["values"] = interface_values

        selected_label = ""
        if preserve_selection and previous_selection in interface_values:
            selected_label = previous_selection
        elif preserve_selection and previous_iface_name:
            for iface_name, ip in self.host_interfaces:
                if iface_name == previous_iface_name:
                    selected_label = f"{iface_name} ({ip})"
                    break
        elif preserve_selection and previous_ip:
            for iface_name, ip in self.host_interfaces:
                if ip == previous_ip:
                    selected_label = f"{iface_name} ({ip})"
                    break
        elif interface_values:
            selected_label = interface_values[0]

        if selected_label:
            self.host_interface_var.set(selected_label)
            self.on_interface_selected()
        else:
            self.host_interface_var.set("")
            self.host_ip_var.set("")

    def refresh_interfaces_only(self) -> None:
        previous_ip = self.host_ip_var.get().strip()
        self.refresh_host_interfaces(preserve_selection=True)
        current_ip = self.host_ip_var.get().strip()
        if current_ip:
            if current_ip != previous_ip:
                self.status_var.set(f"Host interface refreshed. Using {current_ip}.")
            else:
                self.status_var.set(f"Host interface list refreshed. Still using {current_ip}.")
        else:
            self.status_var.set("No usable IPv4 host interface found.")

    def _get_dcp(self):
        if DCP is None:
            raise RuntimeError("pnio_dcp is not installed in this Python environment.")
        host_ip = self.host_ip_var.get().strip()
        if not host_ip:
            raise RuntimeError("Host IP is empty.")
        return DCP(host_ip)

    def scan_devices(self) -> None:
        self.refresh_host_interfaces()
        self.scan_btn.configure(state="disabled")
        self.refresh_btn.configure(state="disabled")
        self.status_var.set("Scanning for PROFINET devices...")
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self) -> None:
        try:
            dcp = self._get_dcp()
            found = dcp.identify_all()
            rows = []
            existing_by_mac = {dev.mac.upper(): dev for dev in self.devices if dev.mac}
            for dev in found:
                mac = str(getattr(dev, "MAC", ""))
                existing = existing_by_mac.get(mac.upper())
                rows.append(
                    DeviceRow(
                        name_of_station=str(getattr(dev, "name_of_station", "")),
                        mac=mac,
                        ip=str(getattr(dev, "IP", "")),
                        netmask=str(getattr(dev, "netmask", "")),
                        gateway=str(getattr(dev, "gateway", "")),
                        family=str(getattr(dev, "family", "")),
                        vendor=(existing.vendor if existing else self.vendor_cache.get(self._mac_prefix(mac), "Looking up...")),
                        ping_status=existing.ping_status if existing else "Unknown",
                        ping_ms=existing.ping_ms if existing else "",
                    )
                )
            self.root.after(0, lambda: self._load_scan_results(rows))
        except Exception as exc:
            self.root.after(0, lambda: self._scan_failed(exc))

    def _load_scan_results(self, rows: List[DeviceRow]) -> None:
        self.devices = rows
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, dev in enumerate(rows):
            self.tree.insert("", "end", iid=str(idx), values=self._device_values(dev))
        self.scan_btn.configure(state="normal")
        self.refresh_btn.configure(state="normal")
        self.status_var.set(f"Found {len(rows)} device(s).")
        self._update_tree_columns()
        if self.show_topology_var.get():
            self.draw_topology()
        self._start_vendor_lookup_for_unknowns()
        if self.ping_monitor_var.get():
            self._ensure_ping_monitor_running()

    def _scan_failed(self, exc: Exception) -> None:
        self.scan_btn.configure(state="normal")
        self.refresh_btn.configure(state="normal")
        self.status_var.set("Scan failed.")
        messagebox.showerror("Scan failed", str(exc))

    def _device_values(self, dev: DeviceRow):
        ping_display = dev.ping_status
        if dev.ping_ms and dev.ping_status.lower().startswith("online"):
            ping_display = f"{dev.ping_status} ({dev.ping_ms})"
        return (
            dev.name_of_station,
            dev.mac,
            dev.vendor,
            dev.ip,
            ping_display,
            dev.netmask,
            dev.gateway,
            dev.family,
        )

    def on_tree_selection_changed(self, _event=None) -> None:
        if self.show_topology_var.get():
            self.draw_topology()
        self._start_quick_action_scan_for_selected()

    def _selected_device(self) -> Optional[DeviceRow]:
        sel = self.tree.selection()
        if not sel:
            return None
        idx = int(sel[0])
        return self.devices[idx]

    def _select_device_by_index(self, idx: int) -> None:
        if idx < 0 or idx >= len(self.devices):
            return
        iid = str(idx)
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self.tree.see(iid)
        if self.show_topology_var.get():
            self.draw_topology()

    def on_canvas_click(self, event) -> None:
        item = self.canvas.find_withtag("current")
        if not item:
            return
        idx = self.canvas_item_to_index.get(item[0])
        if idx is not None:
            self._select_device_by_index(idx)

    def set_ip_for_selected(self) -> None:
        dev = self._selected_device()
        if not dev:
            messagebox.showinfo("No selection", "Pick a device first.")
            return

        values = self._ask_ip_config(
            ip=dev.ip if dev.ip != "0.0.0.0" else "192.168.0.10",
            netmask=dev.netmask if dev.netmask and dev.netmask != "0.0.0.0" else "255.255.255.0",
            gateway=dev.gateway if dev.gateway else "0.0.0.0",
            mac=dev.mac,
        )
        if not values:
            return

        ip, netmask, gateway = values
        try:
            dcp = self._get_dcp()
            dcp.set_ip_address(dev.mac, [ip, netmask, gateway])
            self.status_var.set(f"Assigned {ip} to {dev.mac}")
            self.scan_devices()
        except Exception as exc:
            messagebox.showerror("Set IP failed", str(exc))

    def _ask_ip_config(self, ip: str, netmask: str, gateway: str, mac: str):
        dialog = tk.Toplevel(self.root)
        dialog.title("Set IP configuration")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        try:
            dialog.iconbitmap(self.root.iconbitmap())
        except Exception:
            pass

        frame = ttk.Frame(dialog, padding=14)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=f"Assign IP settings for {mac}").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(frame, text="IP address").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Label(frame, text="Subnet mask").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Label(frame, text="Gateway").grid(row=3, column=0, sticky="w", pady=4)

        ip_var = tk.StringVar(value=ip)
        netmask_var = tk.StringVar(value=netmask)
        gateway_var = tk.StringVar(value=gateway)

        ip_entry = ttk.Entry(frame, textvariable=ip_var, width=22)
        netmask_entry = ttk.Entry(frame, textvariable=netmask_var, width=22)
        gateway_entry = ttk.Entry(frame, textvariable=gateway_var, width=22)
        ip_entry.grid(row=1, column=1, sticky="ew", padx=(12, 0), pady=4)
        netmask_entry.grid(row=2, column=1, sticky="ew", padx=(12, 0), pady=4)
        gateway_entry.grid(row=3, column=1, sticky="ew", padx=(12, 0), pady=4)

        result = {"value": None}

        def submit(event=None):
            result["value"] = (ip_var.get().strip(), netmask_var.get().strip(), gateway_var.get().strip())
            dialog.destroy()

        def cancel(event=None):
            dialog.destroy()

        btns = ttk.Frame(frame)
        btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=cancel).pack(side="right")
        ttk.Button(btns, text="Apply", command=submit).pack(side="right", padx=(0, 8))

        frame.columnconfigure(1, weight=1)
        dialog.bind("<Return>", submit)
        dialog.bind("<Escape>", cancel)
        ip_entry.focus_set()
        dialog.update_idletasks()

        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()
        dialog_w = dialog.winfo_width()
        dialog_h = dialog.winfo_height()
        pos_x = root_x + max((root_w - dialog_w) // 2, 0)
        pos_y = root_y + max((root_h - dialog_h) // 2, 0)
        dialog.geometry(f"+{pos_x}+{pos_y}")
        dialog.lift()
        dialog.focus_force()

        self.root.wait_window(dialog)
        return result["value"]

    def set_name_for_selected(self) -> None:
        dev = self._selected_device()
        if not dev:
            messagebox.showinfo("No selection", "Pick a device first.")
            return
        name = simpledialog.askstring("Set station name", "New PROFINET station name:", initialvalue=dev.name_of_station or "device-01")
        if not name:
            return
        try:
            dcp = self._get_dcp()
            dcp.set_name_of_station(dev.mac, name)
            self.status_var.set(f"Assigned station name '{name}' to {dev.mac}")
            self.scan_devices()
        except Exception as exc:
            messagebox.showerror("Set name failed", str(exc))

    def reset_selected(self) -> None:
        dev = self._selected_device()
        if not dev:
            messagebox.showinfo("No selection", "Pick a device first.")
            return
        if not messagebox.askyesno("Reset communication", f"Reset communication parameters for {dev.mac}?\n\n"):
            return
        try:
            dcp = self._get_dcp()
            if hasattr(dcp, "reset_to_factory"):
                dcp.reset_to_factory(dev.mac)
            elif hasattr(dcp, "reset"):
                dcp.reset(dev.mac)
            else:
                raise RuntimeError("This pnio_dcp version does not expose a reset method.")
            self.status_var.set(f"Reset requested for {dev.mac}")
            self.scan_devices()
        except Exception as exc:
            messagebox.showerror("Reset failed", str(exc))

    def show_selected_details(self) -> None:
        dev = self._selected_device()
        if not dev:
            messagebox.showinfo("No selection", "Pick a device first.")
            return
        ping_line = dev.ping_status
        if dev.ping_ms:
            ping_line = f"{ping_line} ({dev.ping_ms})"
        messagebox.showinfo(
            "Device details",
            f"Station name: {dev.name_of_station or '(empty)'}\n"
            f"MAC: {dev.mac}\n"
            f"Vendor: {dev.vendor}\n"
            f"IP: {dev.ip}\n"
            f"Ping: {ping_line}\n"
            f"Netmask: {dev.netmask}\n"
            f"Gateway: {dev.gateway}\n"
            f"Family: {dev.family}"
        )

    def export_csv(self) -> None:
        import csv
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["station_name", "mac", "vendor", "ip", "ping_status", "ping_ms", "netmask", "gateway", "family"])
            for dev in self.devices:
                writer.writerow([dev.name_of_station, dev.mac, dev.vendor, dev.ip, dev.ping_status, dev.ping_ms, dev.netmask, dev.gateway, dev.family])
        self.status_var.set(f"Exported CSV to {path}")

    def draw_topology(self) -> None:
        if not self.colors:
            self.colors = self._theme_palette()

        self.canvas.delete("all")
        self.canvas_item_to_index = {}
        w = max(self.canvas.winfo_width(), 300)
        h = max(self.canvas.winfo_height(), 200)
        c = self.colors

        pc_rect = self.canvas.create_rectangle(
            w / 2 - 70, 30, w / 2 + 70, 80,
            fill=c["pc_fill"], outline=c["pc_outline"], width=2,
        )
        pc_text = self.canvas.create_text(
            w / 2, 55,
            text="This PC\n(DCP host)",
            font=("Segoe UI", 10, "bold"),
            fill=c["pc_text"],
        )
        self.canvas.tag_bind(pc_rect, "<Button-1>", self.on_canvas_click)
        self.canvas.tag_bind(pc_text, "<Button-1>", self.on_canvas_click)

        if not self.devices:
            self.canvas.create_text(w / 2, h / 2, text="No devices scanned yet.", font=("Segoe UI", 11), fill=c["text"])
            return

        n = len(self.devices)
        spacing = w / (n + 1)
        selected = self.tree.selection()
        selected_idx = int(selected[0]) if selected else None

        for idx, dev in enumerate(self.devices, start=1):
            x = spacing * idx
            y = 180
            is_selected = selected_idx == idx - 1
            fill = c["node_selected_fill"] if is_selected else c["node_fill"]
            outline = c["node_selected_outline"] if is_selected else c["node_outline"]
            status_color = self._status_color(dev.ping_status)

            line_id = self.canvas.create_line(w / 2, 80, x, y - 35, fill=c["line"], dash=(4, 3), width=2)
            oval_id = self.canvas.create_oval(x - 60, y - 35, x + 60, y + 35, fill=fill, outline=outline, width=2)
            status_dot_id = self.canvas.create_oval(
                x - 52, y - 18, x - 40, y - 6,
                fill=status_color, outline=status_color, width=1
            )
            text_id = self.canvas.create_text(
                x + 6,
                y,
                text=f"{dev.family or 'PROFINET device'}\n{dev.ip or '0.0.0.0'}\n{dev.mac}",
                justify="center",
                font=("Segoe UI", 9),
                fill=c["text"],
            )

            self.canvas_item_to_index[oval_id] = idx - 1
            self.canvas_item_to_index[status_dot_id] = idx - 1
            self.canvas_item_to_index[text_id] = idx - 1
            self.canvas_item_to_index[line_id] = idx - 1
            self.canvas.tag_bind(oval_id, "<Button-1>", self.on_canvas_click)
            self.canvas.tag_bind(status_dot_id, "<Button-1>", self.on_canvas_click)
            self.canvas.tag_bind(text_id, "<Button-1>", self.on_canvas_click)
            self.canvas.tag_bind(line_id, "<Button-1>", self.on_canvas_click)

        self.canvas.create_text(
            w / 2,
            h - 24,
            text="Visualized as host-to-device discovery. Physical switch ports and link paths require LLDP/SNMP/MAC-table data.",
            font=("Segoe UI", 9),
            fill=c["muted"],
        )

    def _mac_prefix(self, mac: str) -> str:
        cleaned = "".join(ch for ch in mac.upper() if ch.isalnum())
        return cleaned[:6]

    def _status_color(self, ping_status: str) -> str:
        status = (ping_status or "").lower()
        if status.startswith("online"):
            return "#22c55e"
        if status == "offline":
            return "#ef4444"
        return "#9ca3af"

    def _start_vendor_lookup_for_unknowns(self) -> None:
        threading.Thread(target=self._vendor_lookup_worker, daemon=True).start()

    def _vendor_lookup_worker(self) -> None:
        for idx, dev in enumerate(list(self.devices)):
            prefix = self._mac_prefix(dev.mac)
            if not prefix:
                continue
            with self.vendor_lookup_lock:
                cached = self.vendor_cache.get(prefix)
            if cached:
                if dev.vendor != cached:
                    self.root.after(0, lambda i=idx, v=cached: self._update_device_vendor(i, v))
                continue
            vendor = self.lookup_mac_vendor(dev.mac)
            with self.vendor_lookup_lock:
                self.vendor_cache[prefix] = vendor
            self.root.after(0, lambda i=idx, v=vendor: self._update_device_vendor(i, v))

    def _update_device_vendor(self, idx: int, vendor: str) -> None:
        if idx < 0 or idx >= len(self.devices):
            return
        self.devices[idx].vendor = vendor
        if self.tree.exists(str(idx)):
            self.tree.item(str(idx), values=self._device_values(self.devices[idx]))

    def lookup_mac_vendor(self, mac: str) -> str:
        prefix = self._mac_prefix(mac)
        if not prefix:
            return "Unknown"

        local_fallbacks = {
            "000ECF": "Hirschmann",
            "001B1B": "Siemens AG",
            "080006": "Siemens AG",
            "000AF7": "Phoenix Contact",
            "00000C": "Cisco Systems",
            "3C39E7": "Rockwell Automation",
        }
        if prefix in local_fallbacks:
            return local_fallbacks[prefix]

        urls = [
            f"https://api.macvendors.com/{mac}",
            f"https://api.macvendors.com/{prefix}",
        ]
        headers = {"User-Agent": "Poor-Mans-PRONETA/1.1"}
        for url in urls:
            try:
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=3) as response:
                    vendor = response.read().decode("utf-8", errors="ignore").strip()
                    if vendor:
                        return vendor
            except (urllib.error.URLError, TimeoutError, ValueError):
                continue
            except Exception:
                continue
        return "Unknown"

    def toggle_ping_monitor(self) -> None:
        self._update_tree_columns()
        if self.ping_monitor_var.get():
            self._ensure_ping_monitor_running()
            self.status_var.set("Ping monitor enabled.")
        else:
            self.ping_monitor_stop.set()
            self.status_var.set("Ping monitor disabled.")

    def _ensure_ping_monitor_running(self) -> None:
        if self.ping_thread and self.ping_thread.is_alive():
            return
        self.ping_monitor_stop.clear()
        self.ping_thread = threading.Thread(target=self._ping_monitor_worker, daemon=True)
        self.ping_thread.start()

    def _ping_monitor_worker(self) -> None:
        while not self.ping_monitor_stop.is_set():
            snapshot = [(idx, dev.ip) for idx, dev in enumerate(self.devices)]
            for idx, ip in snapshot:
                if self.ping_monitor_stop.is_set():
                    break
                if not ip or ip == "0.0.0.0":
                    self.root.after(0, lambda i=idx: self._update_ping_status(i, "No IP", ""))
                    continue
                status, latency = self._ping_once(ip)
                self.root.after(0, lambda i=idx, s=status, l=latency: self._update_ping_status(i, s, l))
            self.ping_monitor_stop.wait(5.0)

    def _update_ping_status(self, idx: int, status: str, ping_ms: str) -> None:
        if idx < 0 or idx >= len(self.devices):
            return
        self.devices[idx].ping_status = status
        self.devices[idx].ping_ms = ping_ms
        if self.tree.exists(str(idx)):
            self.tree.item(str(idx), values=self._device_values(self.devices[idx]))
        if self.show_topology_var.get():
            self.draw_topology()

    def _ping_once(self, ip: str):
        is_windows = platform.system().lower().startswith("win")
        if is_windows:
            cmd = ["ping", "-n", "1", "-w", "1000", ip]
        else:
            cmd = ["ping", "-c", "1", "-W", "1", ip]

        try:
            run_kwargs = {
                "capture_output": True,
                "text": True,
                "timeout": 3,
            }

            if is_windows:
                run_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

            completed = subprocess.run(cmd, **run_kwargs)
            output = f"{completed.stdout}\n{completed.stderr}"
            latency = self._extract_ping_ms(output)

            if completed.returncode == 0:
                return "Online", latency
            return "Offline", latency

        except subprocess.TimeoutExpired:
            return "Timeout", ""
        except Exception:
            return "Error", ""

    def _extract_ping_ms(self, output: str) -> str:
        lowered = output.lower()
        markers = ["time=", "time<", "tempo=", "temps="]
        for marker in markers:
            pos = lowered.find(marker)
            if pos == -1:
                continue
            rest = output[pos + len(marker):]
            value = []
            for ch in rest:
                if ch.isdigit() or ch in ".<":
                    value.append(ch)
                elif value:
                    break
            if value:
                return "".join(value).replace("<", "<") + " ms"
        return ""

    def _start_quick_action_scan_for_selected(self) -> None:
        dev = self._selected_device()
        self.port_scan_token += 1
        token = self.port_scan_token
        if not dev or not dev.ip or dev.ip == "0.0.0.0":
            self.root.after(0, lambda: self._set_quick_actions([], message="Quick connect"))
            return

        self._set_quick_actions([], message="Checking ports...")
        threading.Thread(target=self._port_scan_worker, args=(token, dev.ip), daemon=True).start()

    def _port_scan_worker(self, token: int, ip: str) -> None:
        port_labels = {80: "Open web UI (HTTP)", 443: "Open web UI (HTTPS)", 22: "Open SSH session"}
        open_ports = []
        for port in (80, 443, 22):
            if self._is_port_open(ip, port):
                open_ports.append((port, port_labels[port]))
        self.root.after(0, lambda: self._apply_quick_actions(token, ip, open_ports))

    def _apply_quick_actions(self, token: int, ip: str, open_ports) -> None:
        if token != self.port_scan_token:
            return
        actions = []
        for port, label in open_ports:
            if port == 80:
                actions.append((label, lambda target=ip: self._open_url(f"http://{target}")))
            elif port == 443:
                actions.append((label, lambda target=ip: self._open_url(f"https://{target}")))
            elif port == 22:
                actions.append((label, lambda target=ip: self._open_ssh(target)))
        self._set_quick_actions(actions)

    def _set_quick_actions(self, actions, message: str = "Quick connect") -> None:
        self.quick_actions = actions
        self.quick_menu.delete(0, "end")

        if actions:
            for label, command in actions:
                self.quick_menu.add_command(label=label, command=command)
            self.quick_menu_button.configure(text=f"Quick connect ({len(actions)})", state="normal")
        else:
            self.quick_menu.add_command(label=message, state="disabled")
            self.quick_menu_button.configure(text=message, state="disabled")

    def _is_port_open(self, ip: str, port: int, timeout: float = 0.6) -> bool:
        try:
            with socket.create_connection((ip, port), timeout=timeout):
                return True
        except Exception:
            return False

    def _open_url(self, url: str) -> None:
        try:
            webbrowser.open(url, new=2)
            self.status_var.set(f"Opening {url}")
        except Exception as exc:
            messagebox.showerror("Open failed", str(exc))

    def _open_ssh(self, ip: str) -> None:
        username = simpledialog.askstring(
            "SSH Username",
            f"""Enter SSH username for {ip}:

Leave blank to open a plain ssh prompt.""",
            parent=self.root,
        )
        if username is None:
            self.status_var.set("SSH launch cancelled.")
            return

        username = username.strip()
        target = f"{username}@{ip}" if username else ip
        system = platform.system().lower()
        try:
            if system.startswith("win"):
                if shutil.which("wt"):
                    subprocess.Popen(["wt", "new-tab", "ssh", target])
                elif shutil.which("ssh"):
                    subprocess.Popen(["cmd", "/c", "start", "", "cmd", "/k", f"ssh {target}"])
                elif shutil.which("putty"):
                    subprocess.Popen(["putty", "-ssh", target])
                else:
                    raise RuntimeError("No SSH client found. Install OpenSSH or PuTTY.")
            else:
                terminal_cmds = [
                    ["x-terminal-emulator", "-e", f"ssh {target}"],
                    ["gnome-terminal", "--", "ssh", target],
                    ["konsole", "-e", "ssh", target],
                    ["xterm", "-e", f"ssh {target}"],
                ]
                launched = False
                for cmd in terminal_cmds:
                    if shutil.which(cmd[0]):
                        subprocess.Popen(cmd)
                        launched = True
                        break
                if not launched:
                    raise RuntimeError("No supported terminal emulator found to launch SSH.")
            self.status_var.set(f"Launching SSH to {target}")
        except Exception as exc:
            messagebox.showerror("SSH launch failed", str(exc))


    def on_close(self) -> None:
        self.ping_monitor_stop.set()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    app = Freeneta(root)
    root.mainloop()
