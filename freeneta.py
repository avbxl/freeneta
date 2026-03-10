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

        self._build_ui()
        self.apply_theme()

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        top = ttk.Frame(main)
        top.pack(fill="x", pady=(0, 10))

        ttk.Label(top, text="Host IP").pack(side="left")

        self.host_ip_var = tk.StringVar()
        self.host_interface_var = tk.StringVar()
        self.host_interfaces = self.get_host_interfaces()

        self.host_ip_entry = ttk.Entry(top, textvariable=self.host_ip_var, width=18)
        self.host_ip_entry.pack(side="left", padx=(8, 6))

        interface_values = [f"{iface} ({ip})" for iface, ip in self.host_interfaces]

        self.interface_combo = ttk.Combobox(
            top,
            textvariable=self.host_interface_var,
            values=interface_values,
            state="readonly",
            width=30,
        )
        self.interface_combo.pack(side="left", padx=(0, 14))
        self.interface_combo.bind("<<ComboboxSelected>>", self.on_interface_selected)

        if self.host_interfaces:
            first_iface, first_ip = self.host_interfaces[0]
            self.host_interface_var.set(f"{first_iface} ({first_ip})")
            self.host_ip_var.set(first_ip)
        else:
            self.host_ip_var.set("")

        self.scan_btn = ttk.Button(top, text="Scan", command=self.scan_devices)
        self.scan_btn.pack(side="left")

        self.refresh_btn = ttk.Button(top, text="Refresh", command=self.scan_devices)
        self.refresh_btn.pack(side="left", padx=(8, 0))

        self.set_ip_btn = ttk.Button(top, text="Set IP", command=self.set_ip_for_selected)
        self.set_ip_btn.pack(side="left", padx=(18, 0))

        self.set_name_btn = ttk.Button(top, text="Set Name", command=self.set_name_for_selected)
        self.set_name_btn.pack(side="left", padx=(8, 0))

        self.reset_btn = ttk.Button(top, text="Reset Comm", command=self.reset_selected)
        self.reset_btn.pack(side="left", padx=(8, 0))

        self.monitor_chk = ttk.Checkbutton(
            top,
            text="Ping monitor",
            variable=self.ping_monitor_var,
            command=self.toggle_ping_monitor,
        )
        self.monitor_chk.pack(side="left", padx=(18, 0))

        self.dark_mode_chk = ttk.Checkbutton(
            top,
            text="Dark mode",
            variable=self.dark_mode_var,
            command=self.toggle_dark_mode,
        )
        self.dark_mode_chk.pack(side="left", padx=(18, 0))

        self.status_var = tk.StringVar(value="Idle.")
        ttk.Label(top, textvariable=self.status_var).pack(side="right")

        body = ttk.PanedWindow(main, orient="horizontal")
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body, padding=(0, 0, 8, 0))
        right = ttk.Frame(body)
        body.add(left, weight=3)
        body.add(right, weight=2)

        columns = ("name", "mac", "vendor", "ip", "ping", "netmask", "gateway", "family")
        tree_wrap = ttk.Frame(left)
        tree_wrap.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_wrap, columns=columns, show="headings", height=18)
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
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="w", stretch=False)

        self.tree_scroll_y = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        self.tree_scroll_x = ttk.Scrollbar(tree_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree_scroll_y.grid(row=0, column=1, sticky="ns")
        self.tree_scroll_x.grid(row=1, column=0, sticky="ew")
        tree_wrap.grid_rowconfigure(0, weight=1)
        tree_wrap.grid_columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_selection_changed)

        action_row = ttk.Frame(left)
        action_row.pack(fill="x", pady=(8, 0))
        ttk.Button(action_row, text="Export CSV", command=self.export_csv).pack(side="left")
        ttk.Button(action_row, text="Show Selected Details", command=self.show_selected_details).pack(side="left", padx=(8, 0))

        self.quick_menu_button = ttk.Menubutton(action_row, text="Quick connect", state="disabled")
        self.quick_menu_button.pack(side="left", padx=(8, 0))
        self.quick_menu = tk.Menu(self.quick_menu_button, tearoff=False)
        self.quick_menu_button["menu"] = self.quick_menu

        self.topology_title = ttk.Label(right, text="Topology View", font=("Segoe UI", 12, "bold"))
        self.topology_title.pack(anchor="w")
        self.topology_desc = ttk.Label(
            right,
            text="This is a visual summary, not real cable topology. DCP does discovery and commissioning; it does not know physical links.",
            wraplength=360,
            justify="left",
        )
        self.topology_desc.pack(anchor="w", pady=(4, 8))

        self.canvas = tk.Canvas(right, height=520, highlightthickness=1, cursor="hand2")
        self.canvas.pack(fill="both", expand=True)

        self.notes_title = ttk.Label(right, text="Notes", font=("Segoe UI", 11, "bold"))
        self.notes_title.pack(anchor="w", pady=(12, 4))
        self.notes = tk.Text(right, height=9, wrap="word", relief="solid", borderwidth=1)
        self.notes.insert(
            "1.0",
            "Freeneta – v1.5\n\n"
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
        self.notes.pack(fill="x")

        self._update_tree_columns()
        self.draw_topology()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

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
        self.canvas.configure(bg=c["canvas_bg"], highlightbackground=c["canvas_border"])
        self.notes.configure(bg=c["note_bg"], fg=c["text"], insertbackground=c["text"], highlightbackground=c["note_border"])
        self.draw_topology()

    def toggle_dark_mode(self) -> None:
        self.apply_theme()

    def _update_tree_columns(self) -> None:
        if self.ping_monitor_var.get():
            self.tree.configure(displaycolumns=("name", "mac", "vendor", "ip", "ping", "netmask", "gateway", "family"))
        else:
            self.tree.configure(displaycolumns=("name", "mac", "vendor", "ip", "netmask", "gateway", "family"))

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

    def _get_dcp(self):
        if DCP is None:
            raise RuntimeError("pnio_dcp is not installed in this Python environment.")
        host_ip = self.host_ip_var.get().strip()
        if not host_ip:
            raise RuntimeError("Host IP is empty.")
        return DCP(host_ip)

    def scan_devices(self) -> None:
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
            messagebox.showinfo("No selection", "Pick a device first. Telepathy support is still pending.")
            return
        ip = simpledialog.askstring("Set IP", "New IP address:", initialvalue=dev.ip if dev.ip != "0.0.0.0" else "192.168.0.10")
        if not ip:
            return
        netmask = simpledialog.askstring("Set Netmask", "Netmask:", initialvalue=dev.netmask if dev.netmask and dev.netmask != "0.0.0.0" else "255.255.255.0")
        if not netmask:
            return
        gateway = simpledialog.askstring("Set Gateway", "Gateway:", initialvalue=dev.gateway if dev.gateway else "0.0.0.0")
        if gateway is None:
            return
        try:
            dcp = self._get_dcp()
            dcp.set_ip_address(dev.mac, [ip, netmask, gateway])
            self.status_var.set(f"Assigned {ip} to {dev.mac}")
            self.scan_devices()
        except Exception as exc:
            messagebox.showerror("Set IP failed", str(exc))

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
        if not messagebox.askyesno("Reset communication", f"Reset communication parameters for {dev.mac}?\n\nThis is the part where regret often begins."):
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
            text="Visualized as host-to-device discovery. Physical switch ports and link paths require LLDP/SNMP/MAC-table data, because DCP isn't wizardry.",
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
            self.status_var.set("Ping monitor enabled. Because staring at static data wasn't enough.")
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
        self.draw_topology()

    def _ping_once(self, ip: str):
        is_windows = platform.system().lower().startswith("win")
        if is_windows:
            cmd = ["ping", "-n", "1", "-w", "1000", ip]
        else:
            cmd = ["ping", "-c", "1", "-W", "1", ip]
        try:
            completed = subprocess.run(cmd, capture_output=True, text=True, timeout=3)
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
