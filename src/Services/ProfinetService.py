import threading
import time
from typing import Dict, List, Optional, Tuple

from Dataclasses.DeviceRow import DeviceRow
from Services.Constants import (
    DCP_IDENTIFY_RESPONSE_FRAME_ID,
    DCP_RESPONSE,
    DCP_SERVICE_ID_IDENTIFY,
    PROFINET_ETHERTYPE,
    SNIFF_EXTRA_SECONDS,
)

try:
    from scapy.all import Ether, conf, load_contrib, sniff

    load_contrib("pnio")
    load_contrib("pnio_dcp")
    from scapy.contrib.pnio import ProfinetIO
    from scapy.contrib.pnio_dcp import (
        DCPDeviceOptionsBlock,
        DCPNameOfStationBlock,
        DCPIPBlock,
        ProfinetDCP,
    )

    SCAPY_DCP_AVAILABLE = True
except Exception:
    sniff = conf = Ether = ProfinetIO = ProfinetDCP = DCPIPBlock = (
        DCPNameOfStationBlock
    ) = DCPDeviceOptionsBlock = None
    SCAPY_DCP_AVAILABLE = False

try:
    from pnio_dcp import DCP
except Exception:
    DCP = None


class ProfinetService:
    """Service for PROFINET DCP discovery, access level analysis, and configuration."""

    @staticmethod
    def is_dcp_available() -> bool:
        return DCP is not None

    @staticmethod
    def is_scapy_dcp_available() -> bool:
        return SCAPY_DCP_AVAILABLE

    @staticmethod
    def _get_dcp_client(host_ip: str):
        if DCP is None:
            raise RuntimeError("pnio_dcp is not installed in this Python environment.")
        if not host_ip or not host_ip.strip():
            raise RuntimeError(
                "Host IP is empty. Please select a valid host interface."
            )
        return DCP(host_ip.strip())

    @staticmethod
    def _find_scapy_iface_by_ip(target_ip: str) -> Optional[str]:
        if not SCAPY_DCP_AVAILABLE or conf is None:
            return None
        try:
            for name, iface in conf.ifaces.items():
                iface_ip = getattr(iface, "ip", None)
                if iface_ip == target_ip:
                    return name
        except Exception:
            return None
        return None

    @staticmethod
    def _parse_raw_dcp_response(pkt) -> Optional[dict]:
        if not SCAPY_DCP_AVAILABLE:
            return None
        try:
            if Ether not in pkt or pkt[Ether].type != PROFINET_ETHERTYPE:
                return None
            if (
                ProfinetIO not in pkt
                or pkt[ProfinetIO].frameID != DCP_IDENTIFY_RESPONSE_FRAME_ID
            ):
                return None
            if ProfinetDCP not in pkt:
                return None

            dcp = pkt[ProfinetDCP]
            if (
                dcp.service_id != DCP_SERVICE_ID_IDENTIFY
                or dcp.service_type != DCP_RESPONSE
            ):
                return None

            result = {
                "mac": pkt[Ether].src.lower(),
                "name_writable": None,
                "ip_writable": None,
                "options_writable": None,
            }

            for block in getattr(dcp, "dcp_blocks", []):
                if isinstance(block, DCPNameOfStationBlock):
                    result["name_writable"] = bool(block.block_info & 0x0001)
                elif isinstance(block, DCPIPBlock):
                    result["ip_writable"] = bool(block.block_info & 0x0001)
                elif isinstance(block, DCPDeviceOptionsBlock):
                    opts = [(o.option, o.sub_option) for o in block.device_options]
                    result["options_writable"] = any(
                        o in {(2, 2), (1, 2)} for o in opts
                    )
            return result
        except Exception:
            return None

    @staticmethod
    def _determine_dcp_access_from_raw(raw: Optional[dict]) -> str:
        if raw is None:
            return "unknown"

        name_rw = raw.get("name_writable")
        ip_rw = raw.get("ip_writable")
        opts_rw = raw.get("options_writable")

        if name_rw is True or ip_rw is True:
            return "read_write"
        if name_rw is False and ip_rw is False:
            return "read_only"
        if name_rw is False or ip_rw is False:
            return "read_only"
        if opts_rw is True:
            return "read_write"
        if opts_rw is False:
            return "read_only"
        return "unknown"

    def discover_devices(
        self,
        host_ip: str,
        existing_devices: Optional[List[DeviceRow]] = None,
    ) -> List[DeviceRow]:
        """
        Discovers PROFINET devices on the specified host interface.
        Performs raw Scapy packet sniffing if available to detect read/write access.
        """
        dcp = self._get_dcp_client(host_ip)
        access_by_mac: Dict[str, str] = {}

        if SCAPY_DCP_AVAILABLE:
            iface = self._find_scapy_iface_by_ip(host_ip.strip())
            raw_by_mac: Dict[str, dict] = {}
            sniffer_done = threading.Event()

            def handle(pkt):
                parsed = self._parse_raw_dcp_response(pkt)
                if parsed:
                    raw_by_mac.setdefault(parsed["mac"], parsed)

            def do_sniff():
                timeout = 3 + SNIFF_EXTRA_SECONDS
                try:
                    if iface:
                        sniff(iface=iface, prn=handle, timeout=timeout, store=False)
                    else:
                        sniff(prn=handle, timeout=timeout, store=False)
                finally:
                    sniffer_done.set()

            sniffer_thread = threading.Thread(target=do_sniff, daemon=True)
            sniffer_thread.start()
            time.sleep(0.3)
            found = dcp.identify_all()
            sniffer_done.wait(timeout=3 + SNIFF_EXTRA_SECONDS + 1)

            for dev in found:
                mac = str(getattr(dev, "MAC", "")).lower()
                access_by_mac[mac] = self._determine_dcp_access_from_raw(
                    raw_by_mac.get(mac)
                )
        else:
            found = dcp.identify_all()

        existing_by_mac = {d.mac.upper(): d for d in (existing_devices or []) if d.mac}

        devices: List[DeviceRow] = []
        for dev in found:
            mac = str(getattr(dev, "MAC", ""))
            existing = existing_by_mac.get(mac.upper())
            devices.append(
                DeviceRow(
                    name_of_station=str(getattr(dev, "name_of_station", "")),
                    mac=mac,
                    ip=str(getattr(dev, "IP", "")),
                    netmask=str(getattr(dev, "netmask", "")),
                    gateway=str(getattr(dev, "gateway", "")),
                    family=str(getattr(dev, "family", "")),
                    dcp_access=access_by_mac.get(mac.lower(), "unknown"),
                    vendor=existing.vendor if existing else "Looking up...",
                    ping_status=existing.ping_status if existing else "Unknown",
                    ping_ms=existing.ping_ms if existing else "",
                )
            )

        return devices

    def set_device_ip(
        self,
        host_ip: str,
        mac: str,
        ip: str,
        netmask: str,
        gateway: str,
        persistent: bool,
    ) -> None:
        """Sets device IP configuration over DCP."""
        dcp = self._get_dcp_client(host_ip)
        dcp.set_ip_address(mac, [ip, netmask, gateway], store_permanent=persistent)

    def set_device_name(self, host_ip: str, mac: str, name: str) -> None:
        """Sets PROFINET station name over DCP."""
        dcp = self._get_dcp_client(host_ip)
        dcp.set_name_of_station(mac, name)

    def blink_device(self, host_ip: str, mac: str) -> None:
        """Blinks the device's LED for identification."""
        dcp = self._get_dcp_client(host_ip)
        dcp.blink(mac)

    def reset_device_communication(self, host_ip: str, mac: str) -> None:
        """Resets device communication parameters to factory defaults."""
        dcp = self._get_dcp_client(host_ip)
        if hasattr(dcp, "reset_to_factory"):
            dcp.reset_to_factory(mac)
        elif hasattr(dcp, "reset"):
            dcp.reset(mac)
        else:
            raise RuntimeError("This pnio_dcp version does not expose a reset method.")
