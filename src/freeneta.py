import threading
import tkinter as tk
from tkinter import filedialog
from typing import List, Optional

from Dataclasses import DeviceRow, Device
from Services import (
    ConnectionService,
    ExportService,
    NetworkService,
    PingMonitor,
    PingService,
    ProfinetService,
    VendorService,
)
from UI import Dialogs, MainWindow, ThemeManager


class Freeneta:
    """Main FreeNeta Application Controller."""

    def __init__(self, root: tk.Tk):
        self.root = root

        # Initialize Services
        self.profinet_service = ProfinetService()
        self.network_service = NetworkService()
        self.vendor_service = VendorService()
        self.ping_service = PingService()
        self.ping_monitor = PingMonitor(self.ping_service)
        self.connection_service = ConnectionService()
        self.export_service = ExportService()

        # Application State
        self.devices: List[DeviceRow] = []
        self.host_interfaces = []
        self.port_scan_token = 0

        # UI & Theme Initialization
        self.theme_manager = ThemeManager(self.root)
        callbacks = {
            "on_interface_selected": self.on_interface_selected,
            "refresh_interfaces": self.refresh_interfaces_only,
            "scan_devices": self.scan_devices,
            "set_ip": self.set_ip_for_selected,
            "set_name": self.set_name_for_selected,
            "blink": self.blink_selected,
            "reset_comm": self.reset_selected,
            "export_csv": self.export_csv,
            "show_details": self.show_selected_details,
            "on_device_selected": self.on_device_selected,
            "on_toggle_ping_monitor": self.on_toggle_ping_monitor,
        }
        self.window = MainWindow(self.root, self.theme_manager, callbacks)
        self.window.apply_theme()

        # Initial Network Interface Setup
        self.refresh_host_interfaces(preserve_selection=False)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # -------------------------------------------------------------------------
    # Network Interface Management
    # -------------------------------------------------------------------------

    def refresh_host_interfaces(self, preserve_selection: bool = True) -> None:
        previous_selection = (
            self.window.host_interface_var.get() if preserve_selection else ""
        )
        previous_iface_name = (
            previous_selection.split(" (", 1)[0] if previous_selection else ""
        )
        previous_ip = self.window.host_ip_var.get().strip()

        self.host_interfaces = self.network_service.get_host_interfaces()
        interface_values = [f"{iface} ({ip})" for iface, ip in self.host_interfaces]
        self.window.interface_combo["values"] = interface_values

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
            selected_label = self.network_service.get_preferred_interface(
                self.host_interfaces
            )

        if selected_label:
            self.window.host_interface_var.set(selected_label)
            self.on_interface_selected()
        else:
            self.window.host_interface_var.set("")
            self.window.host_ip_var.set("")

    def on_interface_selected(self) -> None:
        selected = self.window.host_interface_var.get()
        for iface_name, ip in self.host_interfaces:
            label = f"{iface_name} ({ip})"
            if selected == label:
                self.window.host_ip_var.set(ip)
                break

    def refresh_interfaces_only(self) -> None:
        previous_ip = self.window.host_ip_var.get().strip()
        self.refresh_host_interfaces(preserve_selection=True)
        current_ip = self.window.host_ip_var.get().strip()
        current_interface = self.window.host_interface_var.get()
        current_iface_name = (
            current_interface.split(" (", 1)[0] if current_interface else ""
        )
        if current_ip:
            if current_ip != previous_ip:
                self.window.status_var.set(
                    f"Host interface refreshed. \r\nUsing {current_interface}."
                )
            else:
                self.window.status_var.set(
                    f"Host interface list refreshed. \r\nStill using {current_iface_name}."
                )
        else:
            self.window.status_var.set("No usable IPv4 host interface found.")

    # -------------------------------------------------------------------------
    # PROFINET Discovery & Device Operations
    # -------------------------------------------------------------------------

    def scan_devices(self) -> None:
        self.refresh_host_interfaces()
        self.window.scan_btn.configure(state="disabled")
        self.window.status_var.set("Scanning for PROFINET devices...")
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self) -> None:
        try:
            host_ip = self.window.host_ip_var.get().strip()
            found = self.profinet_service.discover_devices(
                host_ip=host_ip,
                existing_devices=self.devices,
            )
            self.root.after(0, lambda: self._load_scan_results(found))
        except Exception as exc:
            self.root.after(0, lambda: self._scan_failed(exc))

    def _load_scan_results(self, rows: List[DeviceRow]) -> None:
        self.devices = rows
        self.window.table_view.set_devices(rows)
        self.window.set_device_selected_state(False)
        self.window.scan_btn.configure(state="normal")

        read_only_count = sum(
            1 for dev in rows if dev.dcp_access_normalized == "read_only"
        )
        read_write_count = sum(
            1 for dev in rows if dev.dcp_access_normalized == "read_write"
        )
        self.window.status_var.set(
            f"Found {len(rows)} device(s). RW: {read_write_count}  RO: {read_only_count}"
        )

        if self.window.show_topology_var.get():
            self.window.topology_view.set_devices(rows)

        self._start_vendor_lookup_for_unknowns()

        if self.window.ping_monitor_var.get():
            self._ensure_ping_monitor_running()

    def _scan_failed(self, exc: Exception) -> None:
        self.window.scan_btn.configure(state="normal")
        self.window.status_var.set("Scan failed.")
        Dialogs.show_error("Scan failed", str(exc))

    def set_ip_for_selected(self) -> None:
        dev = self.window.table_view.get_selected_device()
        if not dev:
            Dialogs.show_info("No selection", "Pick a device first.")
            return

        values = Dialogs.ask_ip_config(
            parent=self.root,
            mac=dev.mac,
            initial_ip=dev.ip if dev.ip != "0.0.0.0" else "192.168.0.10",
            initial_netmask=(
                dev.netmask
                if dev.netmask and dev.netmask != "0.0.0.0"
                else "255.255.255.0"
            ),
            initial_gateway=dev.gateway if dev.gateway else "0.0.0.0",
            initial_persistent=False,
        )
        if not values:
            return

        ip, netmask, gateway, persistent = values
        try:
            host_ip = self.window.host_ip_var.get().strip()
            self.profinet_service.set_device_ip(
                host_ip, dev.mac, ip, netmask, gateway, persistent
            )
            self.window.status_var.set(
                f"Assigned {ip} to {dev.mac}"
                if persistent
                else f"Assigned {ip} to {dev.mac} (temporary)"
            )
            self.scan_devices()
        except Exception as exc:
            Dialogs.show_error("Set IP failed", str(exc))

    def blink_selected(self) -> None:
        dev = self.window.table_view.get_selected_device()
        if not dev:
            Dialogs.show_info("No selection", "Pick a device first.")
            return

        try:
            host_ip = self.window.host_ip_var.get().strip()
            self.profinet_service.blink_device(host_ip, dev.mac)
            self.window.status_var.set(f"Blink command sent to {dev.mac}")
        except Exception as exc:
            Dialogs.show_error("Blink failed", str(exc))

    def set_name_for_selected(self) -> None:
        dev = self.window.table_view.get_selected_device()
        if not dev:
            Dialogs.show_info("No selection", "Pick a device first.")
            return

        name = Dialogs.ask_station_name(self.root, dev.name_of_station)
        if not name:
            return

        try:
            host_ip = self.window.host_ip_var.get().strip()
            self.profinet_service.set_device_name(host_ip, dev.mac, name)
            self.window.status_var.set(f"Assigned station name '{name}' to {dev.mac}")
            self.scan_devices()
        except Exception as exc:
            Dialogs.show_error("Set name failed", str(exc))

    def reset_selected(self) -> None:
        dev = self.window.table_view.get_selected_device()
        if not dev:
            Dialogs.show_info("No selection", "Pick a device first.")
            return

        if not Dialogs.confirm_reset(dev.mac):
            return

        try:
            host_ip = self.window.host_ip_var.get().strip()
            self.profinet_service.reset_device_communication(host_ip, dev.mac)
            self.window.status_var.set(
                f"Factory reset command sent to {dev.mac}. Rescanning..."
            )
            self.scan_devices()
        except Exception as exc:
            Dialogs.show_error("Factory reset failed", str(exc))

    def show_selected_details(self) -> None:
        dev = self.window.table_view.get_selected_device()
        if not dev:
            Dialogs.show_info("No selection", "Pick a device first.")
            return
        Dialogs.show_device_details(dev)

    def export_csv(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if not path:
            return
        self.export_service.export_to_csv(path, self.devices)
        self.window.status_var.set(f"Exported CSV to {path}")

    # -------------------------------------------------------------------------
    # Vendor Resolution
    # -------------------------------------------------------------------------

    def _start_vendor_lookup_for_unknowns(self) -> None:
        def on_resolved(idx: int, vendor: str):
            self.root.after(0, lambda i=idx, v=vendor: self._update_device_vendor(i, v))

        self.vendor_service.resolve_vendors_async(self.devices, on_resolved)

    def _update_device_vendor(self, idx: int, vendor: str) -> None:
        if 0 <= idx < len(self.devices):
            self.devices[idx].vendor = vendor
            self.window.table_view.update_device_row(idx)
            self.window.table_view.autosize_tree_columns(columns=["vendor"])

    # -------------------------------------------------------------------------
    # Ping Monitoring
    # -------------------------------------------------------------------------

    def on_toggle_ping_monitor(self, enabled: bool) -> None:
        if enabled:
            self._ensure_ping_monitor_running()
            self.window.status_var.set("Ping monitor enabled.")
        else:
            self.ping_monitor.stop()
            self.window.status_var.set("Ping monitor disabled.")

    def _ensure_ping_monitor_running(self) -> None:
        def on_ping_updated(idx: int, status: str, latency: str):
            self.root.after(
                0, lambda i=idx, s=status, l=latency: self._update_ping_status(i, s, l)
            )

        self.ping_monitor.start(
            get_devices_callback=lambda: self.devices,
            on_ping_updated=on_ping_updated,
        )

    def _update_ping_status(self, idx: int, status: str, ping_ms: str) -> None:
        if 0 <= idx < len(self.devices):
            self.devices[idx].ping_status = status
            self.devices[idx].ping_ms = ping_ms
            self.window.table_view.update_device_row(idx)
            if self.window.ping_monitor_var.get():
                self.window.table_view.autosize_tree_columns(columns=["ping"])
            if self.window.show_topology_var.get():
                self.window.topology_view.draw()

    # -------------------------------------------------------------------------
    # Quick Actions & Port Scanning
    # -------------------------------------------------------------------------

    def on_device_selected(self, device: Optional[DeviceRow]) -> None:
        self.port_scan_token += 1
        token = self.port_scan_token

        if not device or not device.ip or device.ip == "0.0.0.0":
            self.window.set_quick_actions([], message="Quick connect")
            return

        self.window.set_quick_actions([], message="Checking ports...")
        threading.Thread(
            target=self._port_scan_worker,
            args=(token, device.ip),
            daemon=True,
        ).start()

    def _port_scan_worker(self, token: int, ip: str) -> None:
        open_ports = self.connection_service.scan_standard_ports(ip)
        self.root.after(0, lambda: self._apply_quick_actions(token, ip, open_ports))

    def _apply_quick_actions(self, token: int, ip: str, open_ports) -> None:
        if token != self.port_scan_token:
            return

        actions = []
        for port, label in open_ports:
            if port == 80:
                actions.append(
                    (label, lambda target=ip: self._open_url(f"http://{target}"))
                )
            elif port == 443:
                actions.append(
                    (label, lambda target=ip: self._open_url(f"https://{target}"))
                )
            elif port == 22:
                actions.append((label, lambda target=ip: self._open_ssh(target)))
        self.window.set_quick_actions(actions)

    def _open_url(self, url: str) -> None:
        try:
            self.connection_service.open_url(url)
            self.window.status_var.set(f"Opening {url}")
        except Exception as exc:
            Dialogs.show_error("Open failed", str(exc))

    def _open_ssh(self, ip: str) -> None:
        username = Dialogs.ask_ssh_username(self.root, ip)
        if username is None:
            self.window.status_var.set("SSH launch cancelled.")
            return

        try:
            self.connection_service.open_ssh(ip, username)
            target = f"{username.strip()}@{ip}" if username.strip() else ip
            self.window.status_var.set(f"Launching SSH to {target}")
        except Exception as exc:
            Dialogs.show_error("SSH launch failed", str(exc))

    # -------------------------------------------------------------------------
    # Lifecycle
    # -------------------------------------------------------------------------

    def on_close(self) -> None:
        self.ping_monitor.stop()
        self.root.destroy()


# Alias for backward compatibility
FreenetaApplication = Freeneta

if __name__ == "__main__":
    root = tk.Tk()
    try:
        from tkinter import ttk

        ttk.Style().theme_use("clam")
    except Exception:
        pass
    app = Freeneta(root)
    root.mainloop()
