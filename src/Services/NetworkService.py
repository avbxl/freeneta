import socket
from typing import List, Tuple
import psutil


class NetworkService:
    """Service for discovering and managing host network interfaces."""

    @staticmethod
    def get_host_interfaces() -> List[Tuple[str, str]]:
        """
        Enumerates host IPv4 network interfaces and returns a sorted list of (interface_name, ip).
        Interfaces are prioritized: Physical Ethernet > Wi-Fi > Virtual/VPN.
        """
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
            ethernet_score = (
                0
                if any(x in name for x in ["ethernet", "eth", "enp", "eno", "ens"])
                else 1
            )
            wireless_score = (
                1 if any(x in name for x in ["wlan", "wi-fi", "wifi", "wl"]) else 0
            )
            virtual_score = (
                1
                if any(
                    x in name
                    for x in [
                        "vmware",
                        "virtual",
                        "vbox",
                        "hyper-v",
                        "loopback",
                        "bluetooth",
                        "tailscale",
                        "tun",
                        "tap",
                        "docker",
                        "br-",
                        "virbr",
                        "veth",
                        "wg",
                        "zt",
                    ]
                )
                else 0
            )
            return (virtual_score, wireless_score, ethernet_score, name)

        interfaces.sort(key=sort_key)
        return interfaces

    @staticmethod
    def get_preferred_interface(interfaces: List[Tuple[str, str]]) -> str:
        """Finds the label of the most preferred network interface from a list."""
        if not interfaces:
            return ""

        interface_values = [f"{iface} ({ip})" for iface, ip in interfaces]
        selected_label = interface_values[0]

        for iface_name, ip in interfaces:
            lowered = iface_name.lower()
            virtual_or_wifi = any(
                x in lowered
                for x in [
                    "tailscale",
                    "tun",
                    "tap",
                    "docker",
                    "br-",
                    "virbr",
                    "veth",
                    "wg",
                    "zt",
                    "wlan",
                    "wi-fi",
                    "wifi",
                    "wl",
                ]
            )
            if not virtual_or_wifi:
                selected_label = f"{iface_name} ({ip})"
                break

        return selected_label
