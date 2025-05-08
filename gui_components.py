# -*- coding: utf-8 -*-
import customtkinter as ctk
from tkinter import messagebox, filedialog
import datetime
import time
import logging
import os
from setting import SettingsWindow
from PIL import Image  # Requires Pillow

try:
    from hover_calendar import HoverCalendar
    logging.info("HoverCalendar imported successfully.")
except ImportError:
    HoverCalendar = None
    logging.warning("HoverCalendar not found. Calendar functionality will be disabled.")

class GUIComponents:
    def __init__(self, root, variables, title_var, date_var, display_date, calculator, file_handler, email_sender, settings_manager):
        self.root = root
        self.variables = variables
        self.title_var = title_var
        self.date_var = date_var
        self.display_date = display_date
        self.calculator = calculator
        self.file_handler = file_handler
        self.email_sender = email_sender
        self.settings_manager = settings_manager

        self.required_vars = [
            'logo_path_var', 'address_var',
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
        self.FOOTER_BG = "#192337"  # Solid blue for footer

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
        self.min_input_width = 120
        self.max_input_width = 300
        self.min_column_width = 200  # Reduced for better responsiveness
        self.layout_debounce_delay_ms = 100
        self.debounce_id = None

        self.create_widgets()
        self.setup_keyboard_shortcuts()
        self.date_var.trace_add('write', self._update_display_date)
        self._update_display_date()
        self.root.after(150, self.update_layout)

    def _initialize_missing_variables(self):
        initialized_count = 0
        missing_vars = []
        for var_key in self.required_vars:
            if var_key not in self.variables:
                default_value = ""
                if var_key == 'address_var':
                    default_value = "Default Address - Configure Me"
                self.variables[var_key] = ctk.StringVar(value=default_value)
                initialized_count += 1
                missing_vars.append(var_key)

        if initialized_count > 0:
            logging.warning(f"Initialized {initialized_count} missing StringVars: {missing_vars}")
        elif not self.variables:
            logging.error("Variables dictionary is empty!")

    def _update_display_date(self, *args):
        raw_date = self.date_var.get()
        try:
            date_obj = datetime.datetime.strptime(raw_date, "%m/%d/%Y")
            self.display_date.set(date_obj.strftime("%b %d, %Y"))
        except ValueError:
            if raw_date:
                logging.warning(f"Invalid date format entered: {raw_date}. Expected MM/DD/YYYY.")
            self.display_date.set("Select Date")

    def setup_keyboard_shortcuts(self):
        self.root.bind('<Control-l>', lambda event: self._safe_call(self.file_handler.load_from_documentpdf, "Load"))
        self.root.bind('<Control-e>', lambda event: self._safe_call(self.file_handler.export_to_pdf, "Export to PDF"))
        self.root.bind('<Control-w>', lambda event: self._safe_call(self.file_handler.save_to_docx, "Save to Word"))
        self.root.bind('<Control-g>', lambda event: self._safe_call(self.email_sender.send_email, "Send Email"))
        self.root.bind('<Control-s>', lambda e: messagebox.showinfo("Not Implemented", "Save functionality (Ctrl+S) is not yet implemented."))
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        logging.info("Keyboard shortcuts set up.")

    def _safe_call(self, func, action_name):
        try:
            func()
            logging.info(f"Action '{action_name}' executed successfully.")
        except AttributeError as e:
            logging.error(f"Action '{action_name}' failed: Method not found or attribute missing: {e}")
            messagebox.showerror("Error", f"Could not perform '{action_name}'. Feature might be misconfigured or an object is missing.")
        except FileNotFoundError as e:
            logging.error(f"Action '{action_name}' failed: File not found: {e}")
            messagebox.showerror("File Error", f"File not found during {action_name}:\n{e}")
        except Exception as e:
            logging.exception(f"Error during '{action_name}' action.")
            messagebox.showerror("Error", f"An unexpected error occurred during {action_name}:\n{e}")

    def create_tooltip(self, widget, text):
        tooltip = None
        tooltip_window = None
        def _create_tooltip_window():
            nonlocal tooltip_window
            if tooltip_window is not None and tooltip_window.winfo_exists():
                tooltip_window.destroy()
            tooltip_window = ctk.CTkToplevel(self.root)
            tooltip_window.wm_overrideredirect(True)
            tooltip_window.wm_geometry("+9999+9999")
            label = ctk.CTkLabel(
                tooltip_window, text=text, corner_radius=6, fg_color=self.TOOLTIP_BG,
                text_color=self.TOOLTIP_TEXT, font=("Roboto", self.tooltip_font_size),
                wraplength=max(200, self.screen_width // 4)
            )
            label.pack(padx=self.base_pad_x // 2, pady=self.base_pad_y // 2)
            tooltip_window.withdraw()
        def show(event):
            if tooltip_window is None or not tooltip_window.winfo_exists():
                _create_tooltip_window()
            if tooltip_window and tooltip_window.state() == 'withdrawn':
                widget.update_idletasks()
                x = widget.winfo_rootx() + self.base_pad_x
                y = widget.winfo_rooty() + widget.winfo_height() + self.base_pad_y // 2
                tooltip_window.wm_geometry(f"+{x}+{y}")
                tooltip_window.deiconify()
        def hide(event):
            if tooltip_window and tooltip_window.winfo_exists():
                tooltip_window.withdraw()
        widget.bind("<Enter>", show, add="+")
        widget.bind("<Leave>", hide, add="+")

    def show_calendar(self):
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

    def _select_logo(self):
        filetypes = [("Image files", "*.png *.jpg *.jpeg *.gif *.bmp")]
        filepath = filedialog.askopenfilename(
            title="Select Logo Image",
            filetypes=filetypes
        )
        if filepath:
            if os.path.exists(filepath):
                self.variables['logo_path_var'].set(filepath)
                filename = os.path.basename(filepath)
                self.logo_path_display.configure(text=f"Selected: {filename}" if len(filename) < 40 else f"Selected: ...{filename[-37:]}")
                logging.info(f"Logo selected: {filepath}")
            else:
                messagebox.showerror("Error", f"Selected file does not exist:\n{filepath}")
                logging.error(f"Selected logo file path does not exist: {filepath}")
        else:
            logging.info("Logo selection cancelled.")

    def show_settings(self):
        """Open the settings window."""
        SettingsWindow(self.root, self.settings_manager)

    def create_widgets(self):

        # Main frame with footer space
        self.main_frame = ctk.CTkFrame(self.root, corner_radius=0, fg_color=self.BG_COLOR)
        self.main_frame.pack(fill="both", expand=True)
        self.main_frame.grid_rowconfigure(0, weight=15)  # Increased weight for content area
        self.main_frame.grid_rowconfigure(1, weight=0)  
        self.main_frame.grid_rowconfigure(2, weight=1)  # Footer
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Content frame (container for all content including header, form, and buttons)
        self.content_frame = ctk.CTkFrame(self.main_frame, corner_radius=0, fg_color=self.BG_COLOR)
        self.content_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(0, 10))  # Increased top padding for better spacing
        self.content_frame.grid_columnconfigure(0, weight=1)

        # Footer
        footer_height = 80  # Keep footer height fixed
        self.footer_frame = ctk.CTkFrame(self.main_frame, height=footer_height, corner_radius=0, fg_color=self.FOOTER_BG)
        self.footer_frame.grid(row=2, column=0, sticky="ew")
        self.footer_frame.grid_propagate(False)  # Prevent resizing

        # Footer text container (aligned right)
        footer_text_frame = ctk.CTkFrame(self.footer_frame, fg_color="transparent")
        footer_text_frame.pack(side="right", padx=self.base_pad_x) 

        # First line: compliance text
        compliance_label = ctk.CTkLabel(
            footer_text_frame,
            text="",
            font=("Roboto", 12, "bold"),
            text_color="#FFFFFF",
            justify="left"
        )
        compliance_label.pack(anchor="e", padx=(0, 20))

        # Footer images (aligned right)
        image1_size = (70, 70)
        image2_size = (273, 70)  # Size for footer images
        try:
            image1_path = self.variables.get('footer_image1_var', ctk.StringVar(value="chud logo.png")).get()
            image2_path = self.variables.get('footer_image2_var', ctk.StringVar(value="xu logo.png")).get()

            if os.path.exists(image1_path):
                img1 = Image.open(image1_path)
                img1 = img1.resize(image1_size, Image.Resampling.LANCZOS)
                ctk_img1 = ctk.CTkImage(light_image=img1, dark_image=img1, size=image1_size)
                img1_label = ctk.CTkLabel(self.footer_frame, image=ctk_img1, text="")
                img1_label.pack(side="left", padx=(20, 10), pady=self.base_pad_y)
                logging.info(f"Footer image 1 loaded: {image1_path}")
            else:
                logging.warning(f"Footer image 1 not found: {image1_path}")
                img1_label = ctk.CTkLabel(self.footer_frame, text="Image 1 Not Found", font=("Roboto", self.label_font_size), text_color="#FFFFFF")
                img1_label.pack(side="left", padx=(20, 10), pady=self.base_pad_y)

            if os.path.exists(image2_path):
                img2 = Image.open(image2_path)
                img2 = img2.resize(image2_size, Image.Resampling.LANCZOS)
                ctk_img2 = ctk.CTkImage(light_image=img2, dark_image=img2, size=image2_size)
                img2_label = ctk.CTkLabel(self.footer_frame, image=ctk_img2, text="")
                img2_label.pack(side="left", padx=(10, 20), pady=self.base_pad_y)
                logging.info(f"Footer image 2 loaded: {image2_path}")
            else:
                logging.warning(f"Footer image 2 not found: {image2_path}")
                img2_label = ctk.CTkLabel(self.footer_frame, text="Image 2 Not Found", font=("Roboto", self.label_font_size), text_color="#FFFFFF")
                img2_label.pack(side="left", padx=(10, 20), pady=self.base_pad_y)

        except Exception as e:
            logging.exception("Failed to load footer images.")
            error_label = ctk.CTkLabel(self.footer_frame, text="Error Loading Images", font=("Roboto", self.label_font_size), text_color="#FFFFFF")
            error_label.pack(side="right", padx=self.base_pad_x, pady=self.base_pad_y)

        # Subdivide content frame
        self.content_frame.grid_rowconfigure(0, weight=1)  # Header
        self.content_frame.grid_rowconfigure(1, weight=6)  # Table
        self.content_frame.grid_rowconfigure(2, weight=1)  # Action Buttons

        # Header configuration (Address, Logo, Settings, Date)
        header_config_frame = ctk.CTkFrame(
            self.content_frame,
            fg_color="transparent"
        )
        header_config_frame.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=self.section_pad_x,
            pady=(0, 0) 
        )
        header_config_frame.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="header_group")

        # Address
        address_frame = ctk.CTkFrame(header_config_frame, fg_color="transparent")
        address_frame.grid(row=0, column=0, sticky="ew", padx=(0, self.base_pad_x))
        ctk.CTkLabel(address_frame, text="Header Address:", font=("Roboto", self.label_font_size), text_color=self.TEXT_COLOR, anchor="w").pack(side="top", fill="x", pady=(0, self.base_pad_y // 2))
        address_entry = ctk.CTkEntry(address_frame, textvariable=self.variables['address_var'], font=("Roboto", self.entry_font_size), corner_radius=6, fg_color=self.ENTRY_BG_COLOR, text_color=self.TEXT_COLOR, border_color=self.ENTRY_BORDER_COLOR)
        address_entry.pack(side="top", fill="x")
        self.create_tooltip(address_entry, "Enter the address to display in the document header")

        # Logo
        logo_frame = ctk.CTkFrame(header_config_frame, fg_color="transparent")
        logo_frame.grid(row=0, column=1, sticky="ew", padx=(self.base_pad_x, self.base_pad_x))
        ctk.CTkLabel(logo_frame, text="Header Logo:", font=("Roboto", self.label_font_size), text_color=self.TEXT_COLOR, anchor="w").pack(side="top", fill="x", pady=(0, self.base_pad_y // 2))
        logo_button = ctk.CTkButton(logo_frame, text="Select Logo Image", font=("Roboto", self.button_font_size), command=self._select_logo, corner_radius=6, fg_color=self.BUTTON_FG_COLOR, hover_color=self.BUTTON_HOVER_COLOR, text_color=self.BUTTON_TEXT_COLOR)
        logo_button.pack(side="left", padx=(0, self.base_pad_x))
        self.create_tooltip(logo_button, "Select a logo (PNG, JPG, etc.) for the header")
        initial_logo_path = self.variables['logo_path_var'].get()
        initial_logo_text = ""
        if initial_logo_path:
            filename = os.path.basename(initial_logo_path)
            initial_logo_text = f"Selected: {filename}" if len(filename) < 40 else f"Selected: ...{filename[-37:]}"
        self.logo_path_display = ctk.CTkLabel(logo_frame, text=initial_logo_text, font=("Roboto", self.label_font_size-1), text_color=self.TEXT_COLOR, anchor="w", wraplength=150)
        self.logo_path_display.pack(side="left", fill="x", expand=True)

        # Settings
        settings_frame = ctk.CTkFrame(header_config_frame, fg_color="transparent")
        settings_frame.grid(row=0, column=2, sticky="ew", padx=(self.base_pad_x, self.base_pad_x))
        ctk.CTkLabel(settings_frame, text="Settings:", font=("Roboto", self.label_font_size), text_color=self.TEXT_COLOR, anchor="w").pack(side="top", fill="x", pady=(0, self.base_pad_y // 2))
        settings_button_width = min(max(120, int(self.screen_width * 0.08)), 180)
        settings_button = ctk.CTkButton(settings_frame, text="Manage Settings", font=("Roboto", self.button_font_size), command=self.show_settings, corner_radius=6, fg_color=self.BUTTON_FG_COLOR, hover_color=self.BUTTON_HOVER_COLOR, text_color=self.BUTTON_TEXT_COLOR, width=settings_button_width)
        settings_button.pack(side="top", anchor="w")
        self.create_tooltip(settings_button, "Configure application settings (email, login credentials)")

        # Date
        date_frame = ctk.CTkFrame(header_config_frame, fg_color="transparent")
        date_frame.grid(row=0, column=3, sticky="ew", padx=(self.base_pad_x, 0))
        ctk.CTkLabel(date_frame, text="Report Date:", font=("Roboto", self.label_font_size), text_color=self.TEXT_COLOR, anchor="w").pack(side="top", fill="x", pady=(0, self.base_pad_y // 2))
        date_button_width = min(max(120, int(self.screen_width * 0.08)), 180)
        date_button = ctk.CTkButton(date_frame, textvariable=self.display_date, font=("Roboto", self.button_font_size), command=self.show_calendar, corner_radius=6, fg_color=self.DATE_BTN_FG, hover_color=self.DATE_BTN_HOVER, text_color=self.DATE_BTN_TEXT, width=date_button_width)
        date_button.pack(side="top", anchor="w")
        self.create_tooltip(date_button, "Click to select the report date")

        # Form sections (middle tables)
        self.columns_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.columns_frame.grid(row=2, column=0, sticky="nsew", padx=self.section_pad_x / 2, pady=(0, self.section_pad_y))
        num_data_cols = 5
        min_col_width = int(self.min_input_width * 1.7)
        for i in range(num_data_cols):
            self.columns_frame.grid_columnconfigure(i, weight=1, uniform="data_cols", minsize=min_col_width)
        self.columns_frame.grid_rowconfigure(0, weight=1)
        self.beg_frame = self._create_section_frame("Beginning Cash Balances")
        self.inflow_frame = self._create_section_frame("Cash Inflows")
        self.outflow_frame = self._create_section_frame("Cash Outflows")
        self.end_frame = self._create_section_frame("Ending Cash Balances (Calculated)")
        self.totals_frame = self._create_section_frame("Totals (Calculated)")
        self.beg_frame.grid(row=0, column=0, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.inflow_frame.grid(row=0, column=1, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.outflow_frame.grid(row=0, column=2, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.end_frame.grid(row=0, column=3, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.totals_frame.grid(row=0, column=4, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.populate_columns()

        # Action buttons (bottom of content)
        button_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        button_frame.grid(row=1, column=0, sticky="ew", padx=self.section_pad_x, pady=(self.base_pad_y // 2, self.section_pad_y // 2))
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
            btn.grid(row=0, column=i, sticky="ew", padx=self.base_pad_x // 4, pady=self.base_pad_y // 4)
            self.create_tooltip(btn, tooltip_text)

        # Names section
        names_frame = ctk.CTkFrame(
            self.content_frame,
            corner_radius=8,
            fg_color=self.FRAME_COLOR,
            border_width=1,
            border_color=self.BORDER_COLOR,
        )
        names_frame.grid(
            row=3, 
            column=0,
            sticky="ew",
            padx=self.section_pad_x,
            pady=(0, 0)
        )
        name_fields_data = [
            ("Recipients (comma-separated):", 'recipient_emails_var', "Enter recipient emails, comma-separated"),
            ("Prepared by (Treasurer):", 'prepared_by_var', "Name of HOA Treasurer"),
            ("Noted by (President):", 'noted_by_var_1', "Name of HOA President"),
            ("Noted by (CHUDD HCD-CORDS):", 'noted_by_var_2', "Name of CHUDD HCD-CORDS rep"),
            ("Checked by (Auditor):", 'checked_by_var', "Name of HOA Auditor")
        ]
        num_name_fields = len(name_fields_data)
        min_name_col_width = int(self.min_input_width * 1.5)
        for i in range(num_name_fields):
            names_frame.grid_columnconfigure(i, weight=1, uniform="name_group", minsize=min_name_col_width)
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

        self.main_frame.bind("<Configure>", self.debounce_layout, add="+")


    def _create_section_frame(self, title):
        frame = ctk.CTkFrame(self.columns_frame, corner_radius=8, fg_color=self.FRAME_COLOR, border_width=1, border_color=self.BORDER_COLOR)
        ctk.CTkLabel(
            frame, text=title, font=("Roboto", self.title_font_size, "bold"),
            text_color=self.TEXT_COLOR, anchor="w"
        ).pack(fill="x", padx=self.base_pad_x * 1.5, pady=(self.base_pad_y * 1.5, self.base_pad_y))
        content_frame = ctk.CTkFrame(frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=self.base_pad_x, pady=(0, self.base_pad_y * 1.5))
        return frame

    def _create_entry_pair(self, parent_frame, label_text, var, is_disabled=False, tooltip_text=None, input_width=None):
        item_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
        item_frame.pack(fill="x", padx=self.base_pad_x // 2, pady=self.base_pad_y // 2)
        ctk.CTkLabel(
            item_frame, text=label_text, font=("Roboto", self.label_font_size),
            text_color=self.TEXT_COLOR, anchor="w"
        ).pack(side="left", fill="x", expand=True, padx=(0, self.base_pad_x))
        input_width = input_width or min(max(self.min_input_width, int(self.screen_width * 0.07)), self.max_input_width)
        entry = ctk.CTkEntry(
            item_frame, textvariable=var, width=input_width,
            font=("Roboto", self.entry_font_size), corner_radius=6,
            fg_color=self.DISABLED_BG_COLOR if is_disabled else self.ENTRY_BG_COLOR,
            text_color=self.TEXT_COLOR, border_color=self.ENTRY_BORDER_COLOR,
            state="disabled" if is_disabled else "normal",
            justify="right"
        )
        entry.pack(side="right")
        if not is_disabled:
            try:
                if hasattr(self.calculator, 'format_entry'):
                    if var.get():
                        self.calculator.format_entry(var, entry)
                else:
                    logging.warning("Calculator object missing 'format_entry' method.")
                tooltip = tooltip_text if tooltip_text else "Enter amount (numeric)"
                self.create_tooltip(entry, tooltip)
            except Exception as e:
                logging.exception(f"Error applying format_entry or tooltip to {label_text}: {e}")
        elif tooltip_text:
            self.create_tooltip(entry, tooltip_text)

    def populate_columns(self):
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
        for label, var_key, *tooltip in inflow_items: self._create_entry_pair(
            inflow_content, 
            label, 
            self.variables[var_key], 
            tooltip_text=tooltip[0] if tooltip else None,
            input_width=self.min_input_width // 1.5
            )

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

    def debounce_layout(self, event=None):
        if event and event.widget != self.main_frame:
            return
        if self.debounce_id: self.root.after_cancel(self.debounce_id)
        self.debounce_id = self.root.after(self.layout_debounce_delay_ms, self.update_layout)

    def update_layout(self):
        self.main_frame.update_idletasks()
        self.content_frame.update_idletasks()
        logging.debug("Updated static layout.")
        self.debounce_id = None

    def clear_fields(self):
        if not self.variables:
            logging.error("Cannot clear fields: variables dictionary missing.")
            messagebox.showerror("Error", "Internal error: Cannot access data fields.")
            return
        if not messagebox.askyesno("Confirm Clear", "Clear all input and calculated fields?\n(Logo, Address, Date, and Signatories will remain)"):
            return

        cleared_count = 0
        fields_to_keep = {
            'logo_path_var', 'address_var', 'date_var', 'display_date',
            'recipient_emails_var', 'prepared_by_var', 'noted_by_var_1',
            'noted_by_var_2', 'checked_by_var', 'title_var'
        }

        for key, var in self.variables.items():
            if key not in fields_to_keep and isinstance(var, ctk.StringVar):
                var.set("")
                cleared_count += 1

        logging.info(f"Cleared {cleared_count} StringVar fields.")
        messagebox.showinfo("Success", "Cash flow data fields have been cleared.")

        try:
            if hasattr(self.calculator, 'calculate_totals'):
                self.calculator.calculate_totals()
            else:
                logging.warning("Calculator has no 'calculate_totals' method to call after clearing.")
        except Exception as e:
            logging.exception("Error recalculating totals after clearing fields.")

if __name__ == '__main__':
    class MockCalculator:
        def format_entry(self, var, entry): pass
        def calculate_totals(self): print("Mock calculate_totals called")

    class MockFileHandler:
        def load_from_documentpdf(self): messagebox.showinfo("Mock", "Load called")
        def export_to_pdf(self): messagebox.showinfo("Mock", "Export PDF called")
        def save_to_docx(self): messagebox.showinfo("Mock", "Save Word called")

    class MockEmailSender:
        def send_email(self): messagebox.showinfo("Mock", "Send Email called")

    from setting import SettingsManager
    root = ctk.CTk()
    root.title("HOA Cash Flow (Fixed Horizontal)")
    root.geometry("1000x600")  # Smaller initial size
    root.resizable(True, True)  # Allow resizing and maximize/restore
    root.minsize(800, 500)  # Set minimum window size
    root.state("normal")  # Ensure normal state
    root.attributes("-fullscreen", False)  # Disable fullscreen
    root.attributes("-toolwindow", False)  # Disable tool window mode
    print(f"Window state: {root.state()}")  # Debug log

    variables = {}
    title_var = ctk.StringVar(value="HOA Cash Flow Statement")
    date_var = ctk.StringVar(value=datetime.date.today().strftime("%m/%d/%Y"))
    display_date = ctk.StringVar()

    calculator = MockCalculator()
    file_handler = MockFileHandler()
    email_sender = MockEmailSender()
    settings_manager = SettingsManager()

    gui = GUIComponents(root, variables, title_var, date_var, display_date, calculator, file_handler, email_sender, settings_manager)

    root.after(200, gui.update_layout)
    root.mainloop()