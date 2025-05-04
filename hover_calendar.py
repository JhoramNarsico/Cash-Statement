from tkcalendar import Calendar
import tkinter as tk
import calendar

class HoverCalendar(Calendar):
    """Custom Calendar class styled to resemble Google Calendar with hover navigation."""
    def __init__(self, master=None, **kw):
        # Set default Google Calendar-like styles
        kw.setdefault('font', ("Roboto", 12))  # Use Roboto for consistency with app
        kw.setdefault('background', '#FFFFFF')  # White background
        kw.setdefault('foreground', '#202124')  # Dark gray text
        kw.setdefault('bordercolor', '#DADCE0')  # Light gray borders
        kw.setdefault('selectbackground', '#1A73E8')  # Google blue for selected day
        kw.setdefault('selectforeground', '#FFFFFF')  # White text on selected day
        kw.setdefault('normalbackground', '#FFFFFF')  # White for days
        kw.setdefault('normalforeground', '#202124')  # Dark gray for day numbers
        kw.setdefault('weekendbackground', '#FFFFFF')  # Same as normal days
        kw.setdefault('weekendforeground', '#202124')
        kw.setdefault('headersbackground', '#FFFFFF')  # White header background
        kw.setdefault('headersforeground', '#202124')  # Dark gray header text
        kw.setdefault('cursor', 'hand2')  # Hand cursor for interactivity
        kw.setdefault('zoom', 1.2)  # Slightly larger cells
        kw.setdefault('showweeknumbers', False)  # Hide week numbers
        kw.setdefault('showothermonthdays', False)  # Hide days from other months

        super().__init__(master, **kw)
        self._setup_google_calendar_style()
        self._setup_hover_navigation()

    def _setup_google_calendar_style(self):
        """Apply Google Calendar-like styling to the calendar."""
        # Style day cells
        for i in range(6):  # Rows
            for j in range(7):  # Columns
                if (i, j) in self._calendar:
                    cell = self._calendar[i, j]
                    cell.configure(
                        bg='#FFFFFF',
                        fg='#202124',
                        borderwidth=1,
                        relief='solid',
                        highlightthickness=0,
                        font=("Roboto", 12)
                    )
                    # Add hover effect for days
                    cell.bind("<Enter>", self._on_day_hover)
                    cell.bind("<Leave>", self._on_day_leave)

        # Style weekday headers (Mon, Tue, etc.)
        for widget in self.winfo_children():
            if isinstance(widget, tk.Frame):  # Headers are typically in a frame
                for child in widget.winfo_children():
                    if isinstance(child, tk.Label) and child.cget("text") in ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]:
                        child.configure(
                            bg='#FFFFFF',
                            fg='#5F6368',  # Lighter gray for headers
                            font=("Roboto", 10, "bold")
                        )

        # Style navigation arrows
        for widget in self.winfo_children():
            if isinstance(widget, tk.Frame):  # Navigation buttons are in the header frame
                for child in widget.winfo_children():
                    if isinstance(child, tk.Button) and child.cget("text") in ["<", ">"]:
                        child.configure(
                            bg='#FFFFFF',
                            fg='#1A73E8',
                            activebackground='#E8F0FE',
                            font=("Roboto", 12)
                        )

        # Style month and year labels
        for widget in self.winfo_children():
            if isinstance(widget, tk.Frame):  # Header frame
                for child in widget.winfo_children():
                    if isinstance(child, tk.Label):
                        text = child.cget("text")
                        # Check for month names
                        if text in calendar.month_name[1:]:  # January to December
                            self._header_month_label = child
                            child.configure(
                                bg='#FFFFFF',
                                fg='#1A73E8',
                                font=("Roboto", 14, "bold")
                            )
                        # Check for year (numeric, e.g., "2025")
                        elif text.isdigit() and len(text) == 4:
                            self._header_year_label = child
                            child.configure(
                                bg='#FFFFFF',
                                fg='#1A73E8',
                                font=("Roboto", 14, "bold")
                            )

    def _on_day_hover(self, event):
        """Highlight day cell on hover."""
        widget = event.widget
        if widget.cget("state") != 'disabled':
            widget.configure(bg='#E8F0FE')  # Google blue hover background

    def _on_day_leave(self, event):
        """Reset day cell on leave."""
        widget = event.widget
        if widget.cget("state") != 'disabled':
            widget.configure(bg='#FFFFFF')

    def _setup_hover_navigation(self):
        """Add hover bindings to month and year labels with Google Calendar style."""
        # Bind hover events to month and year labels (already set in _setup_google_calendar_style)
        if hasattr(self, '_header_month_label'):
            self._header_month_label.bind("<Enter>", self._on_month_hover)
            self._header_month_label.bind("<Leave>", self._on_month_leave)
        if hasattr(self, '_header_year_label'):
            self._header_year_label.bind("<Enter>", self._on_year_hover)
            self._header_year_label.bind("<Leave>", self._on_year_leave)

    def _on_month_hover(self, event):
        self._header_month_label.configure(fg='#174EA6')  # Darker blue on hover
        # Simulate click on right arrow to change month
        for widget in self.winfo_children():
            if isinstance(widget, tk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, tk.Button) and child.cget("text") == ">":
                        child.invoke()

    def _on_month_leave(self, event):
        self._header_month_label.configure(fg='#1A73E8')  # Restore Google blue

    def _on_year_hover(self, event):
        self._header_year_label.configure(fg='#174EA6')  # Darker blue on hover
        self._date = self._date.replace(year=self._date.year + 1)
        self._setup_calendar()

    def _on_year_leave(self, event):
        self._header_year_label.configure(fg='#1A73E8')  # Restore Google blue