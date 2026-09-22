# PROFINET DCP protocol constants, IDs, and fallback dictionaries

DCP_IDENTIFY_RESPONSE_FRAME_ID = 0xFEFF
DCP_SERVICE_ID_IDENTIFY = 0x05
DCP_RESPONSE = 0x01
PROFINET_ETHERTYPE = 0x8892
SNIFF_EXTRA_SECONDS = 2

LOCAL_OUI_FALLBACKS = {
    "000ECF": "Hirschmann",
    "001B1B": "Siemens AG",
    "080006": "Siemens AG",
    "000AF7": "Phoenix Contact",
    "00000C": "Cisco Systems",
    "3C39E7": "Rockwell Automation",
}

APP_NOTES = """FreeNeta – v1.6

Features:
- Discover PROFINET devices using DCP
- Show station name, MAC, vendor, IP, subnet, and gateway
- Identify device vendor via MAC OUI lookup
- Detect DCP access level (read-only vs read-write) via raw DCP analysis
- Optional ping monitor to check device reachability
- Visual topology summary of discovered devices
- Set device IP address
- Set station name
- Reset communication parameters to factory defaults (PROFINET DCP Reset)
- Quick‑connect actions based on detected ports
- Open device Web GUI via HTTP/HTTPS
- Start an SSH session to the device
- Dark mode support
- Adjustable UI layout with hideable topology and notes panels
- Export device list to CSV
- Show selected device details

Refactoring:
- Extensive internal code cleanup and modularization
- Improved performance and test coverage

Vendor lookup note:
- Uses local cache and an online OUI lookup fallback
- If the machine has no internet connection, vendor may remain Unknown
"""
