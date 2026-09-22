import platform
import shutil
import socket
import subprocess
import webbrowser
from typing import List, Tuple


class ConnectionService:
    """Service for port scanning and launching device connection tools (Web GUI, SSH)."""

    PORT_LABELS = {
        80: "Open web UI (HTTP)",
        443: "Open web UI (HTTPS)",
        22: "Open SSH session",
    }

    @staticmethod
    def is_port_open(ip: str, port: int, timeout: float = 0.6) -> bool:
        """Checks if a specific TCP port is open on the target IP."""
        try:
            with socket.create_connection((ip, port), timeout=timeout):
                return True
        except Exception:
            return False

    @classmethod
    def scan_standard_ports(
        cls, ip: str, timeout: float = 0.6
    ) -> List[Tuple[int, str]]:
        """Scans standard device management ports (80, 443, 22)."""
        open_ports = []
        for port in (80, 443, 22):
            if cls.is_port_open(ip, port, timeout=timeout):
                open_ports.append((port, cls.PORT_LABELS[port]))
        return open_ports

    @staticmethod
    def open_url(url: str) -> None:
        """Opens the specified URL in the default web browser."""
        webbrowser.open(url, new=2)

    @staticmethod
    def open_ssh(ip: str, username: str = "") -> None:
        """Launches an SSH session using the preferred OS terminal or SSH client."""
        username = username.strip()
        target = f"{username}@{ip}" if username else ip
        system = platform.system().lower()

        if system.startswith("win"):
            if shutil.which("wt"):
                subprocess.Popen(["wt", "new-tab", "ssh", target])
            elif shutil.which("ssh"):
                subprocess.Popen(
                    ["cmd", "/c", "start", "", "cmd", "/k", f"ssh {target}"]
                )
            elif shutil.which("putty"):
                subprocess.Popen(["putty", "-ssh", target])
            else:
                raise RuntimeError(
                    "No SSH client found. Install OpenSSH, Windows Terminal, or PuTTY."
                )
        else:
            terminal_cmds = [
                ["x-terminal-emulator", "-e", f"ssh {target}"],
                ["gnome-terminal", "--", "ssh", target],
                ["konsole", "-e", "ssh", target],
                ["xterm", "-e", f"ssh {target}"],
            ]
            for cmd in terminal_cmds:
                if shutil.which(cmd[0]):
                    subprocess.Popen(cmd)
                    return
            raise RuntimeError("No supported terminal emulator found to launch SSH.")
