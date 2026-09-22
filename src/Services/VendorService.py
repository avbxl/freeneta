import threading
import urllib.error
import urllib.request
from typing import Callable, Dict, List, Optional

from Dataclasses.DeviceRow import DeviceRow
from Services.Constants import LOCAL_OUI_FALLBACKS


class VendorService:
    """Thread-safe MAC OUI vendor resolution service with local cache and HTTP fallback."""

    def __init__(self, user_agent: str = "FreeNeta/1.5"):
        self.vendor_cache: Dict[str, str] = {}
        self.lock = threading.Lock()
        self.user_agent = user_agent

    @staticmethod
    def extract_mac_prefix(mac: str) -> str:
        """Extracts the first 6 hexadecimal characters of a MAC address (OUI)."""
        cleaned = "".join(ch for ch in mac.upper() if ch.isalnum())
        return cleaned[:6]

    def get_cached_vendor(self, mac: str) -> Optional[str]:
        prefix = self.extract_mac_prefix(mac)
        if not prefix:
            return None
        with self.lock:
            return self.vendor_cache.get(prefix)

    def lookup_mac_vendor(self, mac: str) -> str:
        """Looks up the vendor name for a given MAC address (synchronous)."""
        prefix = self.extract_mac_prefix(mac)
        if not prefix:
            return "Unknown"

        if prefix in LOCAL_OUI_FALLBACKS:
            return LOCAL_OUI_FALLBACKS[prefix]

        urls = [
            f"https://api.macvendors.com/{mac}",
            f"https://api.macvendors.com/{prefix}",
        ]
        headers = {"User-Agent": self.user_agent}

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

    def resolve_vendors_async(
        self,
        devices: List[DeviceRow],
        on_vendor_resolved: Callable[[int, str], None],
    ) -> threading.Thread:
        """
        Starts a background thread to resolve vendor names for any devices with unknown vendors.
        Invokes on_vendor_resolved(device_index, vendor_name) on each resolution.
        """

        def worker():
            for idx, dev in enumerate(list(devices)):
                prefix = self.extract_mac_prefix(dev.mac)
                if not prefix:
                    continue

                with self.lock:
                    cached = self.vendor_cache.get(prefix)

                if cached:
                    if dev.vendor != cached:
                        on_vendor_resolved(idx, cached)
                    continue

                vendor = self.lookup_mac_vendor(dev.mac)
                with self.lock:
                    self.vendor_cache[prefix] = vendor

                on_vendor_resolved(idx, vendor)

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()
        return thread
