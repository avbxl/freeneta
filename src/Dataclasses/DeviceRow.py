from dataclasses import dataclass


@dataclass
class DeviceRow:
    """Domain model representing a PROFINET device row."""

    name_of_station: str
    mac: str
    ip: str
    netmask: str
    gateway: str
    family: str
    dcp_access: str = "unknown"
    vendor: str = "Looking up..."
    ping_status: str = "Unknown"
    ping_ms: str = ""

    @property
    def ping_display(self) -> str:
        if self.ping_ms and self.ping_status.lower().startswith("online"):
            return f"{self.ping_status} ({self.ping_ms})"
        return self.ping_status

    @property
    def dcp_access_normalized(self) -> str:
        if not self.dcp_access:
            return "unknown"
        access = (
            str(self.dcp_access).strip().lower().replace("-", "_").replace(" ", "_")
        )
        if access in {"read_only", "readonly", "ro"}:
            return "read_only"
        if access in {"read_write", "readwrite", "rw"}:
            return "read_write"
        return "unknown"


# Alias Device to DeviceRow for flexibility
Device = DeviceRow
