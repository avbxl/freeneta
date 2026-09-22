from tkinter import ttk


class AutoScrollbar(ttk.Scrollbar):
    def set(self, first, last):
        first = float(first)
        last = float(last)
        if first <= 0.0 and last >= 1.0:
            if self.winfo_ismapped():
                self.grid_remove()
        else:
            if not self.winfo_ismapped():
                self.grid()
        super().set(first, last)
