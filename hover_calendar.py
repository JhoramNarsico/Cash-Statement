from tkcalendar import Calendar

class HoverCalendar(Calendar):
    """Custom Calendar class to enable hovering over month and year for navigation."""
    def __init__(self, master=None, **kw):
        # Increase font size and add zoom for larger cells
        kw.setdefault('font', ("Arial", 14))
        kw.setdefault('zoom', 1)  # Scale calendar grid by 1.5x
        super().__init__(master, **kw)
        self._setup_hover_navigation()

    def _setup_hover_navigation(self):
        """Add hover bindings to month and year labels."""
        self._header_month_label = self._calendar[0][2]  # Month label widget
        self._header_month_label.bind("<Enter>", self._on_month_hover)
        self._header_month_label.bind("<Leave>", self._on_month_leave)
        self._header_year_label = self._calendar[0][4]  # Year label widget
        self._header_year_label.bind("<Enter>", self._on_year_hover)
        self._header_year_label.bind("<Leave>", self._on_year_leave)

    def _on_month_hover(self, event):
        self._header_month_label.configure(fg="blue")
        self._calendar[0][3].event_generate("<Button-1>")

    def _on_month_leave(self, event):
        self._header_month_label.configure(fg=self["foreground"])

    def _on_year_hover(self, event):
        self._header_year_label.configure(fg="blue")
        self._date = self._date.replace(year=self._date.year + 1)
        self._setup_calendar()

    def _on_year_leave(self, event):
        self._header_year_label.configure(fg=self["foreground"])