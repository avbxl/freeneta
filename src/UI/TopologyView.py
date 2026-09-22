import tkinter as tk
from tkinter import ttk
from typing import Callable, Dict, List, Optional

from Dataclasses.DeviceRow import DeviceRow
from UI.Theme import ThemeManager


class TopologyView(ttk.Frame):
    """Component for rendering an interactive topology graph of discovered PROFINET devices."""

    def __init__(
        self,
        parent: tk.Widget,
        theme_manager: ThemeManager,
        on_device_clicked: Optional[Callable[[int], None]] = None,
        *args,
        **kwargs,
    ):
        super().__init__(parent, *args, **kwargs)
        self.theme_manager = theme_manager
        self.on_device_clicked = on_device_clicked
        self.devices: List[DeviceRow] = []
        self.selected_index: Optional[int] = None
        self.canvas_item_to_index: Dict[int, int] = {}

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            self,
            highlightthickness=1,
            cursor="hand2",
            height=self.theme_manager.scaled(380),
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas.bind("<Configure>", lambda _e: self.draw())

    def set_colors(self, bg: str, border: str) -> None:
        self.canvas.configure(bg=bg, highlightbackground=border)

    def set_devices(
        self, devices: List[DeviceRow], selected_index: Optional[int] = None
    ) -> None:
        self.devices = devices
        self.selected_index = selected_index
        self.draw()

    def set_selected_index(self, selected_index: Optional[int]) -> None:
        self.selected_index = selected_index
        self.draw()

    def draw(self) -> None:
        c = self.theme_manager.colors
        if not c:
            c = self.theme_manager.get_palette(False)

        self.canvas.delete("all")
        self.canvas_item_to_index = {}

        w = max(self.canvas.winfo_width(), 300)
        h = max(self.canvas.winfo_height(), 200)

        pc_half_w = self.theme_manager.scaled(90)
        pc_top = self.theme_manager.scaled(28)
        pc_bottom = pc_top + self.theme_manager.scaled(64)

        pc_rect = self.canvas.create_rectangle(
            w / 2 - pc_half_w,
            pc_top,
            w / 2 + pc_half_w,
            pc_bottom,
            fill=c["pc_fill"],
            outline=c["pc_outline"],
            width=max(2, self.theme_manager.scaled(2)),
        )
        pc_text = self.canvas.create_text(
            w / 2,
            (pc_top + pc_bottom) / 2,
            text="This PC\n(DCP host)",
            font="TkHeadingFont",
            fill=c["pc_text"],
        )
        self.canvas.tag_bind(pc_rect, "<Button-1>", self._on_canvas_click)
        self.canvas.tag_bind(pc_text, "<Button-1>", self._on_canvas_click)

        if not self.devices:
            self.canvas.create_text(
                w / 2,
                h / 2,
                text="No devices scanned yet.",
                font="TkDefaultFont",
                fill=c["text"],
            )
            return

        n = len(self.devices)
        spacing = w / (n + 1)

        node_half_w = self.theme_manager.scaled(76)
        node_half_h = self.theme_manager.scaled(42)
        status_dot = self.theme_manager.scaled(8)
        top_anchor_y = pc_bottom
        y = min(
            max(self.theme_manager.scaled(210), h * 0.42),
            h - self.theme_manager.scaled(90),
        )

        for idx, dev in enumerate(self.devices, start=1):
            x = spacing * idx
            is_selected = self.selected_index == idx - 1
            fill = c["node_selected_fill"] if is_selected else c["node_fill"]
            outline = c["node_selected_outline"] if is_selected else c["node_outline"]
            status_color = self.theme_manager.get_status_color(dev.ping_status)

            line_id = self.canvas.create_line(
                w / 2,
                top_anchor_y,
                x,
                y - node_half_h,
                fill=c["line"],
                dash=(4, 3),
                width=max(2, self.theme_manager.scaled(2)),
            )
            oval_id = self.canvas.create_oval(
                x - node_half_w,
                y - node_half_h,
                x + node_half_w,
                y + node_half_h,
                fill=fill,
                outline=outline,
                width=max(2, self.theme_manager.scaled(2)),
            )
            status_dot_id = self.canvas.create_oval(
                x - node_half_w + self.theme_manager.scaled(10),
                y - self.theme_manager.scaled(10),
                x - node_half_w + self.theme_manager.scaled(10) + status_dot * 2,
                y - self.theme_manager.scaled(10) + status_dot * 2,
                fill=status_color,
                outline=status_color,
                width=1,
            )
            text_id = self.canvas.create_text(
                x + 6,
                y,
                text=f"{dev.family or 'PROFINET device'}\n{dev.ip or '0.0.0.0'}\n{dev.mac}",
                justify="center",
                font="TkDefaultFont",
                fill=c["text"],
            )

            device_index = idx - 1
            for item_id in (oval_id, status_dot_id, text_id, line_id):
                self.canvas_item_to_index[item_id] = device_index
                self.canvas.tag_bind(item_id, "<Button-1>", self._on_canvas_click)

    def _on_canvas_click(self, event) -> None:
        item = self.canvas.find_withtag("current")
        if not item:
            return
        idx = self.canvas_item_to_index.get(item[0])
        if idx is not None and self.on_device_clicked:
            self.on_device_clicked(idx)
