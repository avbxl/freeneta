import platform
import subprocess
import threading
from typing import Callable, List, Optional, Tuple

from Dataclasses.DeviceRow import DeviceRow


class PingService:
    """Service for checking network reachability via ICMP ping."""

    @staticmethod
    def extract_ping_ms(output: str) -> str:
        """Parses ICMP ping stdout/stderr to extract latency in ms."""
        lowered = output.lower()
        markers = ["time=", "time<", "tempo=", "temps="]
        for marker in markers:
            pos = lowered.find(marker)
            if pos == -1:
                continue
            rest = output[pos + len(marker) :]
            value = []
            for ch in rest:
                if ch.isdigit() or ch in ".<":
                    value.append(ch)
                elif value:
                    break
            if value:
                return "".join(value).replace("<", "<") + " ms"
        return ""

    @classmethod
    def ping_once(cls, ip: str) -> Tuple[str, str]:
        """
        Executes a single ICMP ping against the target IP.
        Returns a tuple of (status, latency_ms), e.g. ("Online", "2.4 ms").
        """
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
            latency = cls.extract_ping_ms(output)

            if completed.returncode == 0:
                return "Online", latency
            return "Offline", latency

        except subprocess.TimeoutExpired:
            return "Timeout", ""
        except Exception:
            return "Error", ""


class PingMonitor:
    """Background manager for continuously polling device reachability."""

    def __init__(self, ping_service: Optional[PingService] = None):
        self.ping_service = ping_service or PingService()
        self.stop_event = threading.Event()
        self.thread: Optional[threading.Thread] = None

    def is_running(self) -> bool:
        return self.thread is not None and self.thread.is_alive()

    def start(
        self,
        get_devices_callback: Callable[[], List[DeviceRow]],
        on_ping_updated: Callable[[int, str, str], None],
    ) -> None:
        """Starts the background ping worker loop if not already running."""
        if self.is_running():
            return
        self.stop_event.clear()

        def worker():
            while not self.stop_event.is_set():
                devices = get_devices_callback()
                snapshot = [(idx, dev.ip) for idx, dev in enumerate(devices)]
                for idx, ip in snapshot:
                    if self.stop_event.is_set():
                        break
                    if not ip or ip == "0.0.0.0":
                        on_ping_updated(idx, "No IP", "")
                        continue
                    status, latency = self.ping_service.ping_once(ip)
                    on_ping_updated(idx, status, latency)
                self.stop_event.wait(5.0)

        self.thread = threading.Thread(target=worker, daemon=True)
        self.thread.start()

    def stop(self) -> None:
        """Signals the background worker to stop."""
        self.stop_event.set()
