import csv
from typing import List
from Dataclasses.DeviceRow import DeviceRow


class ExportService:
    """Service for exporting device discovery data to files."""

    @staticmethod
    def export_to_csv(file_path: str, devices: List[DeviceRow]) -> None:
        """Exports the list of devices to a CSV file."""
        with open(file_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(
                [
                    "station_name",
                    "mac",
                    "vendor",
                    "ip",
                    "ping_status",
                    "ping_ms",
                    "netmask",
                    "gateway",
                    "family",
                ]
            )
            for dev in devices:
                writer.writerow(
                    [
                        dev.name_of_station,
                        dev.mac,
                        dev.vendor,
                        dev.ip,
                        dev.ping_status,
                        dev.ping_ms,
                        dev.netmask,
                        dev.gateway,
                        dev.family,
                    ]
                )
