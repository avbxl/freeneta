import tkinter as tk
from tkinter import ttk

from UI.AutoScrollbar import AutoScrollbar


class HorizontalScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.h_scrollbar = AutoScrollbar(
            self, orient="horizontal", command=self.canvas.xview
        )
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set)

        self.content = ttk.Frame(self.canvas)
        self.window_id = self.canvas.create_window(
            (0, 0), window=self.content, anchor="nw"
        )

        self.canvas.grid(row=0, column=0, sticky="ew")
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.h_scrollbar.grid_remove()
        self.grid_columnconfigure(0, weight=1)

        self.content.bind("<Configure>", self._on_content_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel, add="+")

    def _on_content_configure(self, _event=None):
        req_width = max(self.content.winfo_reqwidth(), 1)
        req_height = max(self.content.winfo_reqheight(), 1)
        self.canvas.configure(
            height=req_height, scrollregion=(0, 0, req_width, req_height)
        )
        self.after_idle(self._update_scrollbar_visibility)

    def _on_canvas_configure(self, event):
        req_width = max(self.content.winfo_reqwidth(), 1)
        req_height = max(self.content.winfo_reqheight(), 1)
        overflow = req_width > max(event.width, 1)
        self.canvas.itemconfigure(self.window_id, height=req_height)
        self.canvas.itemconfigure(
            self.window_id, width=req_width if overflow else event.width
        )
        self.canvas.configure(scrollregion=(0, 0, req_width, req_height))
        self.after_idle(self._update_scrollbar_visibility)

    def _update_scrollbar_visibility(self):
        req_width = max(self.content.winfo_reqwidth(), 1)
        canvas_width = max(self.canvas.winfo_width(), 1)
        overflow_threshold = 8
        if req_width <= canvas_width + overflow_threshold:
            self.h_scrollbar.grid_remove()
            self.canvas.xview_moveto(0)
            self.canvas.itemconfigure(self.window_id, width=canvas_width)
            self.canvas.configure(
                scrollregion=(
                    0,
                    0,
                    canvas_width,
                    max(self.content.winfo_reqheight(), 1),
                )
            )
        else:
            self.h_scrollbar.grid()
            self.canvas.itemconfigure(self.window_id, width=req_width)
            self.canvas.configure(
                scrollregion=(0, 0, req_width, max(self.content.winfo_reqheight(), 1))
            )

    def _pointer_inside(self):
        widget = self.winfo_containing(self.winfo_pointerx(), self.winfo_pointery())
        while widget is not None:
            if widget == self.canvas:
                return True
            widget = widget.master
        return False

    def _on_shift_mousewheel(self, event):
        if self._pointer_inside() and event.delta:
            self.canvas.xview_scroll(int(-event.delta / 120), "units")
