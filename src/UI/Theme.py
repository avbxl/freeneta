import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont
from typing import Dict


class ThemeManager:
    """Manages high-DPI scaling, font sizing, color palettes, and TTK styling."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.ui_scale = self.detect_ui_scale()
        self.configure_tk_scaling()
        self.default_font = None
        self.text_font = None
        self.heading_font = None
        self.small_font = None
        self.colors: Dict[str, str] = {}
        self.init_fonts()

    def detect_ui_scale(self) -> float:
        screen_w = max(self.root.winfo_screenwidth(), 1)
        screen_h = max(self.root.winfo_screenheight(), 1)
        scale = min(screen_w / 1920.0, screen_h / 1080.0)
        return max(1.0, min(scale, 1.8))

    def configure_tk_scaling(self) -> None:
        try:
            self.root.tk.call("tk", "scaling", max(1.0, self.ui_scale * 1.15))
        except Exception:
            pass

    def scaled(self, value: int) -> int:
        return max(1, int(round(value * self.ui_scale)))

    def scaled_geometry(self, width: int, height: int) -> str:
        return f"{self.scaled(width)}x{self.scaled(height)}"

    def init_fonts(self) -> None:
        self.default_font = tkfont.nametofont("TkDefaultFont")
        self.default_font.configure(size=10)

        self.text_font = tkfont.nametofont("TkTextFont")
        self.text_font.configure(size=10)

        self.heading_font = tkfont.nametofont("TkHeadingFont")
        self.heading_font.configure(size=11, weight="bold")

        self.small_font = self.default_font.copy()
        self.small_font.configure(size=9)

        for font_name in (
            "TkMenuFont",
            "TkCaptionFont",
            "TkSmallCaptionFont",
            "TkTooltipFont",
        ):
            try:
                tkfont.nametofont(font_name).configure(size=10)
            except tk.TclError:
                pass

    @staticmethod
    def get_palette(dark_mode: bool) -> Dict[str, str]:
        if dark_mode:
            return {
                "bg": "#111827",
                "panel": "#1f2937",
                "text": "#f3f4f6",
                "muted": "#9ca3af",
                "canvas_bg": "#0f172a",
                "canvas_border": "#334155",
                "pc_fill": "#1d4ed8",
                "pc_outline": "#60a5fa",
                "pc_text": "#eff6ff",
                "node_fill": "#0f766e",
                "node_outline": "#5eead4",
                "node_selected_fill": "#166534",
                "node_selected_outline": "#86efac",
                "line": "#64748b",
                "note_bg": "#111827",
                "note_border": "#374151",
            }
        return {
            "bg": "#f8fafc",
            "panel": "#ffffff",
            "text": "#111827",
            "muted": "#555555",
            "canvas_bg": "#ffffff",
            "canvas_border": "#cccccc",
            "pc_fill": "#eef2ff",
            "pc_outline": "#8aa0ff",
            "pc_text": "#111827",
            "node_fill": "#ecfeff",
            "node_outline": "#67e8f9",
            "node_selected_fill": "#dcfce7",
            "node_selected_outline": "#22c55e",
            "line": "#888888",
            "note_bg": "#ffffff",
            "note_border": "#d1d5db",
        }

    @staticmethod
    def get_status_color(ping_status: str) -> str:
        status = (ping_status or "").lower()
        if status.startswith("online"):
            return "#22c55e"
        if status == "offline":
            return "#ef4444"
        return "#9ca3af"

    def apply_theme(self, dark_mode: bool, tree: ttk.Treeview) -> Dict[str, str]:
        self.colors = self.get_palette(dark_mode)
        c = self.colors

        try:
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("TFrame", background=c["bg"])
            style.configure("TPanedwindow", background=c["bg"])
            style.configure(
                "TLabel", background=c["bg"], foreground=c["text"], font="TkDefaultFont"
            )
            style.configure("TButton", padding=(10, 6), font="TkDefaultFont")
            style.configure(
                "TMenubutton",
                padding=(10, 6),
                background=c["panel"],
                foreground=c["text"],
                font="TkDefaultFont",
            )
            style.map(
                "TMenubutton",
                background=[("active", c["panel"])],
                foreground=[("active", c["text"])],
            )
            style.configure(
                "TCheckbutton",
                background=c["bg"],
                foreground=c["text"],
                font="TkDefaultFont",
            )
            style.map(
                "TCheckbutton",
                background=[("active", c["bg"])],
                foreground=[("active", c["text"])],
            )
            style.configure("TCombobox", padding=self.scaled(6))
            style.configure(
                "Treeview",
                background=c["panel"],
                foreground=c["text"],
                fieldbackground=c["panel"],
                rowheight=self.scaled(28),
                borderwidth=1,
                relief="solid",
                bordercolor=c["canvas_border"],
                lightcolor=c["canvas_border"],
                darkcolor=c["canvas_border"],
            )
            style.configure(
                "Treeview.Heading",
                background=c["panel"],
                foreground=c["text"],
                relief="raised",
                borderwidth=1,
                padding=(self.scaled(10), self.scaled(8)),
                font="TkHeadingFont",
            )
            style.map(
                "Treeview.Heading", relief=[("active", "raised"), ("pressed", "sunken")]
            )
            style.map(
                "Treeview",
                background=[("selected", "#2563eb" if dark_mode else "#bfdbfe")],
                foreground=[("selected", "#ffffff" if dark_mode else "#111827")],
            )

            if tree is not None:
                tree.tag_configure(
                    "dcp_readonly",
                    background="#f8d7da" if not dark_mode else "#5b1f24",
                    foreground="#111827" if not dark_mode else "#f9fafb",
                )
                tree.tag_configure(
                    "dcp_readwrite",
                    background="#d4edda" if not dark_mode else "#1c4532",
                    foreground="#111827" if not dark_mode else "#f9fafb",
                )
                tree.tag_configure(
                    "dcp_unknown",
                    background=c["panel"],
                    foreground=c["text"],
                )

            style.configure("Vertical.TScrollbar", arrowsize=14)
            style.configure("Horizontal.TScrollbar", arrowsize=14)
        except Exception:
            pass

        self.root.configure(bg=c["bg"])
        return self.colors
