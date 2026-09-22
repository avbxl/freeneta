import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Tuple
from Dataclasses.DeviceRow import DeviceRow


class Dialogs:
    """Helper class for user input dialogs and message boxes."""

    @staticmethod
    def _apply_icon(dialog: tk.Toplevel, parent: tk.Tk) -> None:
        """Applies the application iconbitmap to a dialog window."""
        try:
            dialog.iconbitmap("app.ico")
        except Exception:
            try:
                dialog.iconbitmap(parent.iconbitmap())
            except Exception:
                pass

    @classmethod
    def ask_string(
        cls,
        parent: tk.Tk,
        title: str,
        prompt: str,
        initialvalue: str = "",
    ) -> Optional[str]:
        """Displays a modal string input prompt with proper iconbitmap and styling."""
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.transient(parent)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        cls._apply_icon(dialog, parent)

        frame = ttk.Frame(dialog, padding=14)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=prompt, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10)
        )

        var = tk.StringVar(value=initialvalue)
        entry = ttk.Entry(frame, textvariable=var, width=32)
        entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 12))

        result = {"value": None}

        def submit(event=None):
            result["value"] = var.get().strip()
            dialog.destroy()

        def cancel(event=None):
            dialog.destroy()

        btns = ttk.Frame(frame)
        btns.grid(row=2, column=0, columnspan=2, sticky="e")
        ttk.Button(btns, text="Cancel", command=cancel).pack(side="right")
        ttk.Button(btns, text="OK", command=submit).pack(side="right", padx=(0, 8))

        frame.columnconfigure(0, weight=1)
        dialog.bind("<Return>", submit)
        dialog.bind("<Escape>", cancel)
        entry.focus_set()
        if initialvalue:
            entry.selection_range(0, tk.END)
        dialog.update_idletasks()

        root_x = parent.winfo_rootx()
        root_y = parent.winfo_rooty()
        root_w = parent.winfo_width()
        root_h = parent.winfo_height()
        dialog_w = dialog.winfo_width()
        dialog_h = dialog.winfo_height()
        pos_x = root_x + max((root_w - dialog_w) // 2, 0)
        pos_y = root_y + max((root_h - dialog_h) // 2, 0)
        dialog.geometry(f"+{pos_x}+{pos_y}")
        dialog.lift()
        dialog.focus_force()

        parent.wait_window(dialog)
        return result["value"]

    @classmethod
    def ask_ip_config(
        cls,
        parent: tk.Tk,
        mac: str,
        initial_ip: str,
        initial_netmask: str,
        initial_gateway: str,
        initial_persistent: bool,
    ) -> Optional[Tuple[str, str, str]]:
        """Displays a modal dialog to configure IP address, subnet mask, and gateway."""
        dialog = tk.Toplevel(parent)
        dialog.title("Set IP configuration")
        dialog.transient(parent)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        cls._apply_icon(dialog, parent)

        frame = ttk.Frame(dialog, padding=14)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=f"Assign IP settings for {mac}").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10)
        )
        ttk.Label(frame, text="IP address").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Label(frame, text="Subnet mask").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Label(frame, text="Gateway").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Label(frame, text="Persistent").grid(row=4, column=0, sticky="w", pady=4)

        ip_var = tk.StringVar(value=initial_ip)
        netmask_var = tk.StringVar(value=initial_netmask)
        gateway_var = tk.StringVar(value=initial_gateway)
        persistent_var = tk.BooleanVar(value=initial_persistent)

        ip_entry = ttk.Entry(frame, textvariable=ip_var, width=22)
        netmask_entry = ttk.Entry(frame, textvariable=netmask_var, width=22)
        gateway_entry = ttk.Entry(frame, textvariable=gateway_var, width=22)
        persistent_check = ttk.Checkbutton(
            frame, text="Store permanently", variable=persistent_var
        )

        ip_entry.grid(row=1, column=1, sticky="ew", padx=(12, 0), pady=4)
        netmask_entry.grid(row=2, column=1, sticky="ew", padx=(12, 0), pady=4)
        gateway_entry.grid(row=3, column=1, sticky="ew", padx=(12, 0), pady=4)
        persistent_check.grid(row=4, column=1, sticky="ew", padx=(12, 0), pady=4)

        result = {"value": None}

        def submit(event=None):
            result["value"] = (
                ip_var.get().strip(),
                netmask_var.get().strip(),
                gateway_var.get().strip(),
                persistent_var.get(),
            )
            dialog.destroy()

        def cancel(event=None):
            dialog.destroy()

        btns = ttk.Frame(frame)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=cancel).pack(side="right")
        ttk.Button(btns, text="Apply", command=submit).pack(side="right", padx=(0, 8))

        frame.columnconfigure(1, weight=1)
        dialog.bind("<Return>", submit)
        dialog.bind("<Escape>", cancel)
        ip_entry.focus_set()
        dialog.update_idletasks()

        root_x = parent.winfo_rootx()
        root_y = parent.winfo_rooty()
        root_w = parent.winfo_width()
        root_h = parent.winfo_height()
        dialog_w = dialog.winfo_width()
        dialog_h = dialog.winfo_height()
        pos_x = root_x + max((root_w - dialog_w) // 2, 0)
        pos_y = root_y + max((root_h - dialog_h) // 2, 0)
        dialog.geometry(f"+{pos_x}+{pos_y}")
        dialog.lift()
        dialog.focus_force()

        parent.wait_window(dialog)
        return result["value"]

    @classmethod
    def ask_station_name(cls, parent: tk.Tk, current_name: str) -> Optional[str]:
        """Displays an input prompt to change the PROFINET station name."""
        return cls.ask_string(
            parent=parent,
            title="Set station name",
            prompt="New PROFINET station name:",
            initialvalue=current_name or "device-01",
        )

    @classmethod
    def ask_ssh_username(cls, parent: tk.Tk, ip: str) -> Optional[str]:
        """Displays an input prompt for SSH username."""
        return cls.ask_string(
            parent=parent,
            title="SSH Username",
            prompt=f"Enter SSH username for {ip}:\n\nLeave blank to open a plain ssh prompt.",
            initialvalue="",
        )

    @staticmethod
    def confirm_reset(mac: str) -> bool:
        """Prompts the user with explicit details before issuing a PROFINET DCP factory reset."""
        return messagebox.askyesno(
            "Factory Reset Device (DCP)",
            f"Reset network & communication settings for {mac} to factory defaults?\n\n"
            f"• IP Address, Netmask, and Gateway will be reset (0.0.0.0)\n"
            f"• PROFINET Station Name will be cleared/reset\n"
            f"• Active controller connections will be closed\n\n"
            f"Do you want to proceed?",
        )

    @staticmethod
    def show_device_details(dev: DeviceRow) -> None:
        """Displays full device details in a modal message dialog."""
        ping_line = dev.ping_display
        messagebox.showinfo(
            "Device details",
            f"Station name: {dev.name_of_station or '(empty)'}\n"
            f"MAC: {dev.mac}\n"
            f"Vendor: {dev.vendor}\n"
            f"IP: {dev.ip}\n"
            f"Ping: {ping_line}\n"
            f"Netmask: {dev.netmask}\n"
            f"Gateway: {dev.gateway}\n"
            f"Family: {dev.family}",
        )

    @staticmethod
    def show_error(title: str, message: str) -> None:
        messagebox.showerror(title, message)

    @staticmethod
    def show_info(title: str, message: str) -> None:
        messagebox.showinfo(title, message)
