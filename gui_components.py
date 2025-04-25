# -*- coding: utf-8 -*-
import customtkinter as ctk
from tkinter import messagebox
import datetime
import time
import logging
from tkinter import HORIZONTAL # Import HORIZONTAL

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Note: Replace with actual HoverCalendar import if available
try:
    # from tkcalendar import Calendar # Example using tkcalendar if HoverCalendar unavailable
    # logging.info("tkcalendar imported as fallback.")
    from hover_calendar import HoverCalendar # Keep trying HoverCalendar first
    logging.info("HoverCalendar imported successfully.")
except ImportError:
    HoverCalendar = None # Set to None if neither is found or preferred
    logging.warning("HoverCalendar (or fallback) not found. Calendar functionality will be disabled.")


class GUIComponents:
    """
    Manages the creation and layout of GUI elements for the HOA Cash Flow application,
    with a fixed horizontal layout for the main data sections.
    (Address field removed).
    Includes horizontal scrolling for smaller screens.
    """
    def __init__(self, root, variables, title_var, date_var, display_date, calculator, file_handler, email_sender):
        """
        Initializes the GUIComponents class with fixed horizontal data sections.
        (Args documentation remains the same as before)
        """
        self.root = root
        self.variables = variables
        self.title_var = title_var
        self.date_var = date_var
        self.display_date = display_date
        self.calculator = calculator
        self.file_handler = file_handler
        self.email_sender = email_sender

        # --- Removed 'address_var' from required_vars ---
        self.required_vars = [
             'recipient_emails_var', 'prepared_by_var', 'noted_by_var_1',
            'noted_by_var_2', 'checked_by_var', 'cash_bank_beg', 'cash_hand_beg',
            'monthly_dues', 'certifications', 'membership_fee', 'vehicle_stickers',
            'rentals', 'solicitations', 'interest_income', 'livelihood_fee',
            'inflows_others', 'snacks_meals', 'transportation', 'office_supplies',
            'printing', 'labor', 'billboard', 'cleaning', 'misc_expenses',
            'federation_fee', 'uniforms', 'bod_mtg', 'general_assembly',
            'cash_deposit', 'withholding_tax', 'refund', 'outflows_others',
            'ending_cash_bank', 'ending_cash_hand', 'total_receipts',
            'cash_outflows', 'ending_cash'
        ]
        self._initialize_missing_variables()

        # --- Theme and Appearance (Same as before) ---
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.BG_COLOR = "#F5F5F5"
        self.FRAME_COLOR = "#FFFFFF"
        self.BORDER_COLOR = "#E0E0E0"
        self.TEXT_COLOR = "#333333"
        self.ENTRY_BG_COLOR = "#FAFAFA"
        self.ENTRY_BORDER_COLOR = "#B0BEC5"
        self.DISABLED_BG_COLOR = "#ECEFF1"
        self.BUTTON_FG_COLOR = "#2196F3"
        self.BUTTON_HOVER_COLOR = "#1976D2"
        self.BUTTON_TEXT_COLOR = "#FFFFFF"
        self.TOOLTIP_BG = "#E0E0E0"
        self.TOOLTIP_TEXT = "#333333"
        self.DATE_BTN_FG = "#E3F2FD"
        self.DATE_BTN_HOVER = "#BBDEFB"
        self.DATE_BTN_TEXT = "#0D47A1"

        # --- Screen Dimensions and Adaptive Sizing (Mostly same, but some related to column wrapping removed/unused) ---
        self.root.update_idletasks()
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        logging.info(f"Screen dimensions: {self.screen_width}x{self.screen_height}")
        self.base_font_size = max(9, min(16, int(self.screen_height / 65)))
        self.title_font_size = int(self.base_font_size * 1.2)
        self.button_font_size = self.base_font_size
        self.label_font_size = self.base_font_size
        self.entry_font_size = self.base_font_size
        self.tooltip_font_size = max(8, int(self.base_font_size * 0.9))
        self.base_pad_x = int(self.base_font_size * 0.6)
        self.base_pad_y = int(self.base_font_size * 0.3)
        self.section_pad_x = self.base_pad_x * 2
        self.section_pad_y = self.base_pad_y * 2
        # These column width/threshold values are less critical now for section layout
        # self.min_column_width = 280 # Kept for potential future use or other elements
        # self.max_column_width = 500
        # self.column_layout_threshold_multiplier = 1.7 # No longer used for section wrapping
        self.min_input_width = 120
        self.max_input_width = 300
        # self.last_num_columns = None # No longer needed for section layout
        self.layout_debounce_delay_ms = 100
        self.debounce_id = None

        # --- Initialization ---
        self.create_widgets()
        self.setup_keyboard_shortcuts()
        self.date_var.trace_add('write', self._update_display_date)
        self._update_display_date()
        # Initial layout update might still be needed for canvas size adjustment
        self.root.after(150, self.update_layout) # Keep this to set initial canvas size

    # --- Methods _initialize_missing_variables, _update_display_date, setup_keyboard_shortcuts, _safe_call, create_tooltip, show_calendar remain unchanged ---

    def _initialize_missing_variables(self):
        """Ensures all required StringVars exist in the variables dictionary."""
        initialized_count = 0
        missing_vars = []
        for var_key in self.required_vars:
            if var_key not in self.variables:
                self.variables[var_key] = ctk.StringVar()
                initialized_count += 1
                missing_vars.append(var_key)
        if 'address_var' in self.variables:
            logging.warning("'address_var' found in input variables dictionary but is no longer used.")
        if initialized_count > 0:
            logging.warning(f"Initialized {initialized_count} missing StringVars: {missing_vars}")
        elif not self.variables:
             logging.error("Variables dictionary is empty!")

    def _update_display_date(self, *args):
        """Updates the display date string when date_var changes."""
        raw_date = self.date_var.get()
        try:
            date_obj = datetime.datetime.strptime(raw_date, "%m/%d/%Y")
            self.display_date.set(date_obj.strftime("%b %d, %Y"))
        except ValueError:
            if raw_date:
                 logging.warning(f"Invalid date format entered: {raw_date}. Expected MM/DD/YYYY.")
            self.display_date.set("Select Date")

    def setup_keyboard_shortcuts(self):
        """Binds keyboard shortcuts to specific actions."""
        self.root.bind('<Control-l>', lambda event: self._safe_call(self.file_handler.load_from_documentpdf, "Load"))
        self.root.bind('<Control-e>', lambda event: self._safe_call(self.file_handler.export_to_pdf, "Export to PDF"))
        self.root.bind('<Control-w>', lambda event: self._safe_call(self.file_handler.save_to_docx, "Save to Word"))
        self.root.bind('<Control-g>', lambda event: self._safe_call(self.email_sender.send_email, "Send Email"))
        self.root.bind('<Control-s>', lambda e: messagebox.showinfo("Not Implemented", "Save functionality (Ctrl+S) is not yet implemented."))
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        logging.info("Keyboard shortcuts set up.")

    def _safe_call(self, func, action_name):
        """Safely calls a function and shows an error message if it fails."""
        try:
            func()
            logging.info(f"Action '{action_name}' executed successfully.")
        except AttributeError:
             logging.error(f"Action '{action_name}' failed: Method not found.")
             messagebox.showerror("Error", f"Could not perform '{action_name}'. Feature might be misconfigured.")
        except Exception as e:
            logging.exception(f"Error during '{action_name}' action.")
            messagebox.showerror("Error", f"An unexpected error occurred during {action_name}:\n{e}")

    def create_tooltip(self, widget, text):
        """Creates a simple adaptive tooltip for a given widget."""
        tooltip = None
        def _create_tooltip_window():
            nonlocal tooltip
            if tooltip is not None and tooltip.winfo_exists():
                tooltip.destroy()
            tooltip = ctk.CTkToplevel(self.root)
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry("+9999+9999")
            label = ctk.CTkLabel(
                tooltip, text=text, corner_radius=6, fg_color=self.TOOLTIP_BG,
                text_color=self.TOOLTIP_TEXT, font=("Roboto", self.tooltip_font_size),
                wraplength=max(200, self.screen_width // 4)
            )
            label.pack(padx=self.base_pad_x // 2, pady=self.base_pad_y // 2)
            tooltip.withdraw()
        def show(event):
            if tooltip is None or not tooltip.winfo_exists():
                 _create_tooltip_window()
            if tooltip and tooltip.state() == 'withdrawn':
                 widget.update_idletasks()
                 x = widget.winfo_rootx() + self.base_pad_x
                 y = widget.winfo_rooty() + widget.winfo_height() + self.base_pad_y // 2
                 tooltip.wm_geometry(f"+{x}+{y}")
                 tooltip.deiconify()
        def hide(event):
             if tooltip and tooltip.winfo_exists():
                 tooltip.withdraw()
        widget.bind("<Enter>", show, add="+")
        widget.bind("<Leave>", hide, add="+")

    def show_calendar(self):
        """Displays an adaptively sized calendar popup to select a date."""
        if not HoverCalendar:
            logging.warning("Attempted to open calendar, but HoverCalendar is not available.")
            messagebox.showerror("Error", "Calendar functionality is not available (HoverCalendar library missing).")
            return
        popup = ctk.CTkToplevel(self.root)
        popup.title("Select Date")
        popup_width = min(max(350, int(self.screen_width * 0.25)), 550)
        popup_height = min(max(380, int(self.screen_height * 0.35)), 600)
        self.root.update_idletasks()
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()
        x = main_x + (main_width - popup_width) // 2
        y = main_y + (main_height - popup_height) // 2
        popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")
        popup.transient(self.root)
        popup.grab_set()
        popup.attributes("-topmost", True)
        try:
            cal_font_size = max(10, int(self.base_font_size * 1.1))
            cal = HoverCalendar(popup, font=("Roboto", cal_font_size))
            cal.pack(padx=self.section_pad_x, pady=self.section_pad_y, fill="both", expand=True)
            try:
                current_date_str = self.date_var.get()
                if current_date_str:
                    current_date = datetime.datetime.strptime(current_date_str, "%m/%d/%Y")
                    cal.selection_set(current_date)
            except ValueError:
                 logging.warning(f"Could not pre-select date '{current_date_str}' in calendar.")
            def on_date_select():
                selected_date = cal.selection_get()
                if selected_date:
                    self.date_var.set(selected_date.strftime("%m/%d/%Y"))
                    logging.info(f"Date selected from calendar: {self.date_var.get()}")
                else:
                     logging.warning("Calendar closed without selection.")
                popup.destroy()
            confirm_button_width = min(max(100, int(popup_width * 0.3)), 160)
            confirm_button = ctk.CTkButton(
                popup, text="Confirm", command=on_date_select, font=("Roboto", self.button_font_size),
                width=confirm_button_width, corner_radius=8, fg_color=self.BUTTON_FG_COLOR,
                hover_color=self.BUTTON_HOVER_COLOR, text_color=self.BUTTON_TEXT_COLOR,
            )
            confirm_button.pack(pady=(0, self.section_pad_y))
            popup.protocol("WM_DELETE_WINDOW", popup.destroy)
        except Exception as e:
             logging.exception("Failed to create or display the calendar popup.")
             messagebox.showerror("Calendar Error", f"Could not open the calendar: {e}")
             if popup and popup.winfo_exists():
                 popup.destroy()

    # --- MODIFIED create_widgets ---
    def create_widgets(self):
        """Creates and arranges all the main widgets with fixed horizontal data columns."""
        # --- Main Scrolling Area ---
        self.main_frame = ctk.CTkFrame(self.root, corner_radius=0, fg_color=self.BG_COLOR)
        self.main_frame.pack(fill="both", expand=True)
        # Configure main_frame grid: Row 0 for canvas (weight 1), Row 1 for horizontal scrollbar (weight 0)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=0) # Row for horizontal scrollbar
        self.main_frame.grid_columnconfigure(0, weight=1) # Column for canvas and h_scrollbar
        self.main_frame.grid_columnconfigure(1, weight=0) # Column for vertical scrollbar

        # Create Canvas and Scrollbars
        self.canvas = ctk.CTkCanvas(self.main_frame, highlightthickness=0, bg=self.BG_COLOR)
        self.v_scrollbar = ctk.CTkScrollbar(self.main_frame, orientation="vertical", command=self.canvas.yview)
        # ADDED: Horizontal scrollbar
        self.h_scrollbar = ctk.CTkScrollbar(self.main_frame, orientation=HORIZONTAL, command=self.canvas.xview)

        # Create Scrollable Frame (content holder)
        self.scrollable_frame = ctk.CTkFrame(self.canvas, corner_radius=0, fg_color=self.BG_COLOR)
        self.scrollable_frame.bind(
            "<Configure>", lambda e: self.update_layout() # Call update_layout on content resize
        )

        # Place Scrollable Frame into Canvas
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Link Canvas and Scrollbars
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set) # Added xscrollcommand

        # Grid the Canvas and Scrollbars
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.h_scrollbar.grid(row=1, column=0, sticky="ew") # Place horizontal scrollbar below canvas

        # Allow scrollable_frame content to define its width, configure its grid
        self.scrollable_frame.grid_columnconfigure(0, weight=1) # Let content define width

        # --- Mouse Wheel Scrolling (Vertical and Horizontal) ---
        def _scroll_canvas(event):
            # Vertical Scroll (Mouse Wheel)
            if event.state & 0x1: # Check if Shift key is pressed (for horizontal)
                if event.num == 5 or event.delta < 0: self.canvas.xview_scroll(1, "units")
                elif event.num == 4 or event.delta > 0: self.canvas.xview_scroll(-1, "units")
            else: # No Shift key (for vertical)
                if event.num == 5 or event.delta < 0: self.canvas.yview_scroll(1, "units")
                elif event.num == 4 or event.delta > 0: self.canvas.yview_scroll(-1, "units")

        # Bind mouse wheel for both vertical and horizontal (Shift+Wheel)
        self.canvas.bind("<MouseWheel>", _scroll_canvas) # Handles Windows/macOS trackpad vertical scroll
        self.canvas.bind("<Button-4>", _scroll_canvas) # Handles Linux vertical scroll up
        self.canvas.bind("<Button-5>", _scroll_canvas) # Handles Linux vertical scroll down
        # Ensure the scrollable frame also gets wheel events if mouse is over it directly
        self.scrollable_frame.bind("<MouseWheel>", _scroll_canvas)
        self.scrollable_frame.bind("<Button-4>", _scroll_canvas)
        self.scrollable_frame.bind("<Button-5>", _scroll_canvas)
        # Bind Shift+MouseWheel specifically for horizontal scrolling if needed (might be redundant with above state check)
        # self.canvas.bind("<Shift-MouseWheel>", _scroll_canvas) # Handled in _scroll_canvas logic now
        # self.canvas.bind("<Shift-Button-4>", _scroll_canvas) # Linux horizontal scroll left
        # self.canvas.bind("<Shift-Button-5>", _scroll_canvas) # Linux horizontal scroll right
        # self.scrollable_frame.bind("<Shift-MouseWheel>", _scroll_canvas)
        # self.scrollable_frame.bind("<Shift-Button-4>", _scroll_canvas)
        # self.scrollable_frame.bind("<Shift-Button-5>", _scroll_canvas)


        # --- Content within Scrollable Frame ---

        # --- Date Frame (Same as before, placed in row 0) ---
        date_container_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="transparent")
        date_container_frame.grid(row=0, column=0, sticky="ew", padx=self.section_pad_x, pady=(self.section_pad_y, self.base_pad_y))
        date_container_frame.grid_columnconfigure(0, weight=1) # Empty expanding column
        date_container_frame.grid_columnconfigure(1, weight=0) # Date content column
        date_frame = ctk.CTkFrame(date_container_frame, fg_color="transparent")
        date_frame.grid(row=0, column=1, sticky="e", pady=self.base_pad_y)
        ctk.CTkLabel(date_frame, text="Date:", font=("Roboto", self.label_font_size), text_color=self.TEXT_COLOR).pack(side="left", padx=(0, self.base_pad_x // 2))
        date_button_width = min(max(120, int(self.screen_width * 0.08)), 180)
        date_button = ctk.CTkButton(
            date_frame, textvariable=self.display_date, font=("Roboto", self.button_font_size),
            command=self.show_calendar, corner_radius=6, fg_color=self.DATE_BTN_FG,
            hover_color=self.DATE_BTN_HOVER, text_color=self.DATE_BTN_TEXT, width=date_button_width
        )
        date_button.pack(side="left")
        self.create_tooltip(date_button, "Click to select the report date")

        # --- Button Bar Frame (Same as before, placed in row 1) ---
        button_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="transparent")
        button_frame.grid(row=1, column=0, sticky="ew", padx=self.section_pad_x, pady=(self.base_pad_y, self.section_pad_y))
        buttons_data = [
            ("Load (Ctrl+L)", self.file_handler.load_from_documentpdf, "Load data from DOCX/PDF"),
            ("Clear Fields", self.clear_fields, "Clear all input fields"),
            ("Export PDF (Ctrl+E)", self.file_handler.export_to_pdf, "Export as PDF"),
            ("Save Word (Ctrl+W)", self.file_handler.save_to_docx, "Save as DOCX"),
            ("Email (Ctrl+G)", self.email_sender.send_email, "Send via email"),
        ]
        num_buttons = len(buttons_data)
        button_frame.grid_columnconfigure(tuple(range(num_buttons)), weight=1, uniform="button_group")
        for i, (text, command, tooltip_text) in enumerate(buttons_data):
            btn = ctk.CTkButton(
                button_frame, text=text,
                command=lambda cmd=command, txt=text: self._safe_call(cmd, txt),
                font=("Roboto", self.button_font_size), corner_radius=8,
                fg_color=self.BUTTON_FG_COLOR, hover_color=self.BUTTON_HOVER_COLOR,
                text_color=self.BUTTON_TEXT_COLOR, height=int(self.base_font_size * 2.5)
            )
            btn.grid(row=0, column=i, sticky="ew", padx=self.base_pad_x // 2, pady=self.base_pad_y)
            self.create_tooltip(btn, tooltip_text)

        # --- Main Data Columns Container (Placed in row 2) ---
        self.columns_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="transparent")
        self.columns_frame.grid(row=2, column=0, sticky="nsew", padx=self.section_pad_x / 2, pady=0)

        # --- Configure columns_frame for fixed 5 horizontal columns ---
        num_data_cols = 5
        min_col_width = int(self.min_input_width * 1.7) # Minimum width for each data column
        for i in range(num_data_cols):
            self.columns_frame.grid_columnconfigure(i, weight=1, uniform="data_cols", minsize=min_col_width) # Added minsize
        self.columns_frame.grid_rowconfigure(0, weight=1) # Only one row needed for sections

        # --- Create section frames (as children of columns_frame) ---
        self.beg_frame = self._create_section_frame("Beginning Cash Balances")
        self.inflow_frame = self._create_section_frame("Cash Inflows")
        self.outflow_frame = self._create_section_frame("Cash Outflows")
        self.end_frame = self._create_section_frame("Ending Cash Balances (Calculated)")
        self.totals_frame = self._create_section_frame("Totals (Calculated)")

        # --- Place section frames directly into the grid ---
        self.beg_frame.grid(row=0, column=0, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.inflow_frame.grid(row=0, column=1, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.outflow_frame.grid(row=0, column=2, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.end_frame.grid(row=0, column=3, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.totals_frame.grid(row=0, column=4, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)

        # --- Populate the section frames (content remains the same) ---
        self.populate_columns()

        # --- Names/Signatories Frame (Same as before, placed in row 3) ---
        names_frame = ctk.CTkFrame(self.scrollable_frame, corner_radius=8, fg_color=self.FRAME_COLOR, border_width=1, border_color=self.BORDER_COLOR)
        names_frame.grid(row=3, column=0, sticky="ew", padx=self.section_pad_x, pady=(self.section_pad_y, self.section_pad_y))
        name_fields_data = [
            ("Recipients (comma-separated):", 'recipient_emails_var', "Enter recipient emails, comma-separated"),
            ("Prepared by (Treasurer):", 'prepared_by_var', "Name of HOA Treasurer"),
            ("Noted by (President):", 'noted_by_var_1', "Name of HOA President"),
            ("Noted by (CHUDD HCD-CORDS):", 'noted_by_var_2', "Name of CHUDD HCD-CORDS rep"),
            ("Checked by (Auditor):", 'checked_by_var', "Name of HOA Auditor")
        ]
        num_name_fields = len(name_fields_data)
        min_name_col_width = int(self.min_input_width * 1.5)
        for i in range(num_name_fields): # Configure columns for name fields
            names_frame.grid_columnconfigure(i, weight=1, uniform="name_group", minsize=min_name_col_width)

        #names_frame.grid_columnconfigure(tuple(range(num_name_fields)), weight=1, uniform="name_group") # Original line
        for i, (label_text, var_key, tooltip_text) in enumerate(name_fields_data):
            frame = ctk.CTkFrame(names_frame, fg_color="transparent")
            frame.grid(row=0, column=i, sticky="nsew", padx=self.base_pad_x, pady=self.base_pad_y)
            frame.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                frame, text=label_text, font=("Roboto", self.label_font_size),
                text_color=self.TEXT_COLOR, anchor="w"
            ).grid(row=0, column=0, sticky="ew", pady=(0, self.base_pad_y // 2))
            entry = ctk.CTkEntry(
                frame, textvariable=self.variables[var_key], font=("Roboto", self.entry_font_size),
                corner_radius=6, fg_color=self.ENTRY_BG_COLOR, text_color=self.TEXT_COLOR,
                border_color=self.ENTRY_BORDER_COLOR
            )
            entry.grid(row=1, column=0, sticky="ew")
            self.create_tooltip(entry, tooltip_text)

        # --- Resize Binding (Remains the same, calls debounce_layout) ---
        self.main_frame.bind("<Configure>", self.debounce_layout, add="+")

    # --- _create_section_frame, _create_entry_pair, populate_columns remain unchanged ---

    def _create_section_frame(self, title):
        """Helper to create a consistent section frame."""
        # NOTE: This frame is now placed using .grid() in create_widgets, not pack()
        frame = ctk.CTkFrame(self.columns_frame, corner_radius=8, fg_color=self.FRAME_COLOR, border_width=1, border_color=self.BORDER_COLOR)
        # Content uses pack within this frame
        ctk.CTkLabel(
            frame, text=title, font=("Roboto", self.title_font_size, "bold"),
            text_color=self.TEXT_COLOR, anchor="w"
        ).pack(fill="x", padx=self.base_pad_x * 1.5, pady=(self.base_pad_y * 1.5, self.base_pad_y))
        content_frame = ctk.CTkFrame(frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=self.base_pad_x, pady=(0, self.base_pad_y * 1.5))
        return frame # Return the main frame for gridding

    def _create_entry_pair(self, parent_frame, label_text, var, is_disabled=False, tooltip_text=None):
        """Helper to create an adaptive label-entry pair."""
        item_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
        item_frame.pack(fill="x", padx=self.base_pad_x // 2, pady=self.base_pad_y // 2)
        ctk.CTkLabel(
            item_frame, text=label_text, font=("Roboto", self.label_font_size),
            text_color=self.TEXT_COLOR, anchor="w"
        ).pack(side="left", fill="x", expand=True, padx=(0, self.base_pad_x))
        # Adaptive width calculation for the entry itself is still useful
        input_width = min(max(self.min_input_width, int(self.screen_width * 0.07)), self.max_input_width)
        entry = ctk.CTkEntry(
            item_frame, textvariable=var, width=input_width,
            font=("Roboto", self.entry_font_size), corner_radius=6,
            fg_color=self.DISABLED_BG_COLOR if is_disabled else self.ENTRY_BG_COLOR,
            text_color=self.TEXT_COLOR, border_color=self.ENTRY_BORDER_COLOR,
            state="disabled" if is_disabled else "normal"
        )
        entry.pack(side="right")
        if not is_disabled:
             try:
                 # Ensure calculator and format_entry exist before calling
                 if hasattr(self.calculator, 'format_entry'):
                     self.calculator.format_entry(var, entry)
                 else:
                     logging.warning("Calculator object missing 'format_entry' method.")
                 tooltip = tooltip_text if tooltip_text else "Enter amount (numeric)"
                 self.create_tooltip(entry, tooltip)
             except Exception as e: logging.exception(f"Error applying format_entry or tooltip to {label_text}: {e}")
        elif tooltip_text: self.create_tooltip(entry, tooltip_text)

    def populate_columns(self):
        """Populates the section frames with their respective entries."""
        # Get the 'content_frame' which is the second child created by _create_section_frame
        beg_content = self.beg_frame.winfo_children()[1]
        beg_items = [("Cash in Bank:", 'cash_bank_beg', "Starting bank balance"), ("Cash on Hand:", 'cash_hand_beg', "Starting physical cash")]
        for label, var_key, tooltip in beg_items: self._create_entry_pair(beg_content, label, self.variables[var_key], tooltip_text=tooltip)

        inflow_content = self.inflow_frame.winfo_children()[1]
        inflow_items = [
            ("Monthly dues collected:", 'monthly_dues'), ("Certifications issued:", 'certifications'),
            ("Membership fee:", 'membership_fee'), ("Vehicle stickers:", 'vehicle_stickers'),
            ("Rentals:", 'rentals'), ("Solicitations/Donations:", 'solicitations'),
            ("Interest Income:", 'interest_income'), ("Livelihood Fee:", 'livelihood_fee'),
            ("Others:", 'inflows_others', "Other income sources")]
        for label, var_key, *tooltip in inflow_items: self._create_entry_pair(inflow_content, label, self.variables[var_key], tooltip_text=tooltip[0] if tooltip else None)

        outflow_content = self.outflow_frame.winfo_children()[1]
        outflow_items = [
            ("Snacks/Meals:", 'snacks_meals'), ("Transportation:", 'transportation'),
            ("Office supplies:", 'office_supplies'), ("Printing/Photocopy:", 'printing'),
            ("Labor:", 'labor'), ("Billboard expense:", 'billboard'),
            ("Cleaning charges:", 'cleaning'), ("Misc expenses:", 'misc_expenses'),
            ("Federation fee:", 'federation_fee'), ("Uniforms:", 'uniforms'),
            ("BOD Mtg:", 'bod_mtg', "Board meeting expenses"),
            ("General Assembly:", 'general_assembly', "Assembly expenses"),
            ("Cash Deposit:", 'cash_deposit', "Cash moved hand to bank"),
            ("Withholding tax:", 'withholding_tax'), ("Refund:", 'refund'),
            ("Others:", 'outflows_others', "Other expenses")]
        for label, var_key, *tooltip in outflow_items: self._create_entry_pair(outflow_content, label, self.variables[var_key], tooltip_text=tooltip[0] if tooltip else None)

        end_content = self.end_frame.winfo_children()[1]
        end_items = [("Cash in Bank:", 'ending_cash_bank', "Calculated ending bank balance"), ("Cash on Hand:", 'ending_cash_hand', "Calculated ending cash on hand")]
        for label, var_key, tooltip in end_items: self._create_entry_pair(end_content, label, self.variables[var_key], is_disabled=True, tooltip_text=tooltip)

        total_content = self.totals_frame.winfo_children()[1]
        total_items = [("Total Receipts:", 'total_receipts', "Calculated total inflows"), ("Total Outflows:", 'cash_outflows', "Calculated total outflows"), ("Ending Balance:", 'ending_cash', "Calculated total ending cash")]
        for label, var_key, tooltip in total_items: self._create_entry_pair(total_content, label, self.variables[var_key], is_disabled=True, tooltip_text=tooltip)
        logging.info("GUI columns populated into fixed horizontal layout.")


    # --- debounce_layout remains the same ---
    def debounce_layout(self, event=None):
        """Debounces the layout update function call on window resize."""
        # Only trigger if the main_frame itself resizes (or called without event)
        if event and event.widget != self.main_frame:
             # logging.debug(f"Ignoring configure event from {event.widget}") # Debugging line
             return
        # logging.debug(f"Debouncing layout update due to configure event on {event.widget if event else 'timer'}") # Debugging line
        if self.debounce_id: self.root.after_cancel(self.debounce_id)
        self.debounce_id = self.root.after(self.layout_debounce_delay_ms, self.update_layout)

    # --- MODIFIED update_layout ---
    def update_layout(self):
        """Adjusts the canvas size and scrollregion to fit content."""
        # This function ensures the canvas scrollregion is correct and the window
        # drawn on the canvas matches the content's required size.

        self.main_frame.update_idletasks() # Ensure dimensions are current
        self.scrollable_frame.update_idletasks() # Ensure content dimensions are current

        # Set the canvas window item size to the required size of the scrollable frame
        req_width = self.scrollable_frame.winfo_reqwidth()
        req_height = self.scrollable_frame.winfo_reqheight()
        self.canvas.itemconfig(self.canvas_window, width=req_width, height=req_height)

        # Configure the scroll region to encompass the entire scrollable frame
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        # No need to calculate container_width/height here for setting itemconfig width/height
        # The scrollbars and canvas handle the viewable area automatically based on scrollregion.
        logging.debug(f"Updating canvas window item size to {req_width}x{req_height}px and scrollregion.")

        self.debounce_id = None # Reset debounce ID


    # --- clear_fields remains unchanged ---
    def clear_fields(self):
        """Clears all data entry fields and resets calculated fields."""
        if not self.variables:
            logging.error("Cannot clear fields: variables dictionary missing.")
            messagebox.showerror("Error", "Internal error: Cannot access data fields.")
            return
        if not messagebox.askyesno("Confirm Clear", "Clear all input and calculated fields?"):
            return
        cleared_count = 0
        for key, var in self.variables.items():
             # Check if it's a StringVar before setting
            if isinstance(var, ctk.StringVar):
                var.set("")
                cleared_count += 1
            # else: # Optional: log unexpected types
            #     logging.warning(f"Item '{key}' in variables is not a StringVar ({type(var)}).")

        # self.date_var.set("") # Keep date or clear it? User preference. Let's keep it for now.
        # self._update_display_date() # Update display if date was cleared

        logging.info(f"Cleared {cleared_count} StringVar fields.")
        messagebox.showinfo("Success", "All data fields have been cleared.")

        # Trigger recalculation after clearing
        try:
            if hasattr(self.calculator, 'calculate_totals'):
                self.calculator.calculate_totals()
            else:
                 logging.warning("Calculator has no 'calculate_totals' method to call after clearing.")
        except Exception as e:
            logging.exception("Error recalculating totals after clearing fields.")

# Example Usage (requires placeholder classes/functions if run standalone)
if __name__ == '__main__':

    # --- Placeholder classes/functions for testing ---
    class MockCalculator:
        def format_entry(self, var, entry): pass # Does nothing
        def calculate_totals(self): print("Mock calculate_totals called")

    class MockFileHandler:
        def load_from_documentpdf(self): messagebox.showinfo("Mock", "Load called")
        def export_to_pdf(self): messagebox.showinfo("Mock", "Export PDF called")
        def save_to_docx(self): messagebox.showinfo("Mock", "Save Word called")

    class MockEmailSender:
        def send_email(self): messagebox.showinfo("Mock", "Send Email called")
    # --- End Placeholders ---

    root = ctk.CTk()
    root.title("HOA Cash Flow (Fixed Horizontal + Scroll)")
    # Start with a size that might require horizontal scrolling on smaller screens
    # Make it wider initially to show the fixed layout better if space allows
    root.geometry("1400x750") # Increased width for testing

    # Initialize necessary variables
    variables = {} # Let _initialize_missing_variables handle creation
    title_var = ctk.StringVar(value="HOA Cash Flow Statement")
    date_var = ctk.StringVar(value=datetime.date.today().strftime("%m/%d/%Y")) # Default to today
    display_date = ctk.StringVar()

    # Create mock instances
    calculator = MockCalculator()
    file_handler = MockFileHandler()
    email_sender = MockEmailSender()

    # Create the GUI
    gui = GUIComponents(root, variables, title_var, date_var, display_date, calculator, file_handler, email_sender)

    # Force an initial layout update after a short delay to ensure scrollbars appear if needed
    root.after(200, gui.update_layout)

    root.mainloop()