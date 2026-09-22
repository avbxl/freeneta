import tkinter as tk
from tkinter import ttk

from UI.AutoScrollbar import AutoScrollbar


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.v_scrollbar = AutoScrollbar(
            self, orient="vertical", command=self.canvas.yview
        )
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set)

        self.content = ttk.Frame(self.canvas)
        self.window_id = self.canvas.create_window(
            (0, 0), window=self.content, anchor="nw"
        )

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.v_scrollbar.grid_remove()
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.content.bind("<Configure>", self._on_content_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.canvas.bind_all("<Button-4>", self._on_linux_scroll_up, add="+")
        self.canvas.bind_all("<Button-5>", self._on_linux_scroll_down, add="+")

    def _on_content_configure(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.after_idle(self._update_scrollbar_visibility)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)
        self.update_idletasks()
        content_height = self.content.winfo_reqheight()
        self.canvas.itemconfigure(
            self.window_id, height=max(event.height, content_height)
        )
        self.after_idle(self._update_scrollbar_visibility)

    def _update_scrollbar_visibility(self):
        bbox = self.canvas.bbox("all")
        if not bbox:
            self.v_scrollbar.grid_remove()
            return
        _, _, _, content_height = bbox
        canvas_height = max(self.canvas.winfo_height(), 1)
        if content_height <= canvas_height:
            self.v_scrollbar.grid_remove()
            self.canvas.yview_moveto(0)
        else:
            self.v_scrollbar.grid()

    def _pointer_inside(self):
        widget = self.winfo_containing(self.winfo_pointerx(), self.winfo_pointery())
        while widget is not None:
            if widget == self.canvas:
                return True
            widget = widget.master
        return False

    def _on_mousewheel(self, event):
        if not self._pointer_inside():
            return
        if event.delta:
            self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _on_linux_scroll_up(self, _event):
        if self._pointer_inside():
            self.canvas.yview_scroll(-1, "units")

    def _on_linux_scroll_down(self, _event):
        if self._pointer_inside():
            self.canvas.yview_scroll(1, "units")
