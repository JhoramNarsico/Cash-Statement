import customtkinter as ctk
from tkinter import messagebox
import datetime
import time

# Note: Replace with actual HoverCalendar import if available
try:
    from hover_calendar import HoverCalendar
except ImportError:
    HoverCalendar = None
    print("Warning: HoverCalendar not found. Calendar functionality will be disabled.")

class GUIComponents:
    def __init__(self, root, variables, title_var, date_var, display_date, calculator, file_handler, email_sender):
        self.root = root
        self.variables = variables
        self.title_var = title_var
        self.date_var = date_var
        self.display_date = display_date
        self.calculator = calculator
        self.file_handler = file_handler
        self.email_sender = email_sender
        
        # Initialize missing variables
        required_vars = [
            'address_var', 'recipient_emails_var', 'prepared_by_var', 'noted_by_var_1', 
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
        for var in required_vars:
            if var not in self.variables:
                self.variables[var] = ctk.StringVar()
        
        # Configure light theme
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # Get screen dimensions
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        
        # Layout optimization
        self.last_layout_update = 0
        self.last_num_columns = None
        self.layout_debounce_delay = 0.05
        self.debounce_id = None
        
        # Size constraints
        self.min_column_width = 250
        self.max_column_width = 400
        self.min_input_width = 150
        self.max_input_width = 300
        
        self.create_widgets()
        self.setup_keyboard_shortcuts()
        self.date_var.trace('w', self.update_display_date)

    def update_display_date(self, *args):
        raw_date = self.date_var.get()
        try:
            date_obj = datetime.datetime.strptime(raw_date, "%m/%d/%Y")
            self.display_date.set(date_obj.strftime("%b %d, %Y"))
        except ValueError:
            pass

    def setup_keyboard_shortcuts(self):
        self.root.bind('<Control-s>', lambda e: messagebox.showinfo("Not Implemented", "Save to CSV not implemented yet"))
        self.root.bind('<Control-e>', lambda e: self.file_handler.export_to_pdf())
        self.root.bind('<Control-l>', lambda e: self.file_handler.load_from_documentpdf())
        self.root.bind('<Control-g>', lambda e: self.email_sender.send_email())
        self.root.bind('<Control-w>', lambda e: self.file_handler.save_to_docx())

    def create_tooltip(self, widget, text):
        tooltip = ctk.CTkToplevel(widget)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry("+1000+1000")
        font_size = max(9, int(self.screen_height / 80))
        label = ctk.CTkLabel(
            tooltip,
            text=text,
            corner_radius=6,
            fg_color="#E0E0E0",
            text_color="#333333",
            font=("Roboto", font_size)
        )
        label.pack(padx=5, pady=5)

        def show(event):
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + 20
            tooltip.wm_geometry(f"+{x}+{y}")
            tooltip.deiconify()

        def hide(event):
            tooltip.withdraw()

        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)
        tooltip.withdraw()

    def show_calendar(self):
        if not HoverCalendar:
            messagebox.showerror("Error", "Calendar functionality is not available.")
            return
        
        popup = ctk.CTkToplevel(self.root)
        popup.title("Select Date")
        
        popup_width = min(max(300, int(self.screen_width * 0.25)), 600)
        popup_height = min(max(300, int(self.screen_height * 0.25)), 600)
        popup.geometry(f"{popup_width}x{popup_height}")
        popup.transient(self.root)
        popup.grab_set()
        popup.withdraw()

        popup.update_idletasks()
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()

        x = main_x + (main_width - popup_width) // 2
        y = main_y + (main_height - popup_height) // 2
        popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")

        font_size = max(10, int(self.screen_height / 60))
        cal = HoverCalendar(
            popup,
            font=("Roboto", font_size)
        )
        cal.pack(padx=20, pady=20, fill="both", expand=True)

        try:
            current_date = datetime.datetime.strptime(self.date_var.get(), "%m/%d/%Y")
            cal.selection_set(current_date)
        except ValueError:
            pass

        def on_date_select():
            selected_date = cal.selection_get()
            if selected_date:
                self.date_var.set(selected_date.strftime("%m/%d/%Y"))
            popup.destroy()

        confirm_button = ctk.CTkButton(
            popup,
            text="Confirm",
            command=on_date_select,
            font=("Roboto", font_size),
            width=min(max(100, int(self.screen_width * 0.1)), 150),
            corner_radius=8,
            fg_color="#2196F3",
            hover_color="#1976D2"
        )
        confirm_button.pack(pady=20)

        popup.deiconify()
        popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    def create_widgets(self):
        self.main_frame = ctk.CTkFrame(self.root, corner_radius=0, fg_color="#F5F5F5")
        self.main_frame.pack(fill="both", expand=True)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.canvas = ctk.CTkCanvas(self.main_frame, highlightthickness=0, bg="#F5F5F5")
        self.scrollbar = ctk.CTkScrollbar(self.main_frame, orientation="vertical", command=self.canvas.yview)
        self.scrollable_frame = ctk.CTkFrame(self.canvas, corner_radius=0, fg_color="#F5F5F5")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Cross-platform scrolling
        def scroll_canvas(event):
            if event.num == 4 or event.delta > 0:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5 or event.delta < 0:
                self.canvas.yview_scroll(1, "units")
        self.canvas.bind_all("<MouseWheel>", scroll_canvas)
        self.canvas.bind_all("<Button-4>", scroll_canvas)
        self.canvas.bind_all("<Button-5>", scroll_canvas)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        self.scrollable_frame.grid_rowconfigure(4, weight=1)

        base_font_size = max(10, int(self.screen_height / 60))

        header_frame = ctk.CTkFrame(self.scrollable_frame, corner_radius=8, fg_color="#FFFFFF", border_width=1, border_color="#E0E0E0")
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        header_frame.grid_columnconfigure(0, weight=1)

        date_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        date_frame.pack(side="right", padx=10, anchor="e")
        ctk.CTkLabel(date_frame, text="Date:", font=("Roboto", base_font_size), text_color="#333333").pack(side="left")
        date_button = ctk.CTkButton(
            date_frame,
            textvariable=self.display_date,
            font=("Roboto", base_font_size),
            command=self.show_calendar,
            corner_radius=6,
            fg_color="#E3F2FD",
            hover_color="#BBDEFB",
            text_color="#0D47A1",
            width=min(max(120, int(self.screen_width * 0.1)), 200)
        )
        date_button.pack(side="left", padx=5)
        self.create_tooltip(date_button, "Click to select a date from the calendar")

        address_frame = ctk.CTkFrame(self.scrollable_frame, corner_radius=8, fg_color="#FFFFFF", border_width=1, border_color="#E0E0E0")
        address_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        
        input_width = min(max(self.min_input_width, int(self.screen_width * 0.2)), self.max_input_width)
        
        ctk.CTkLabel(address_frame, text="Address:", font=("Roboto", base_font_size), text_color="#333333").pack(side="left", padx=10)
        address_entry = ctk.CTkEntry(
            address_frame,
            textvariable=self.variables['address_var'],
            width=input_width * 2,
            font=("Roboto", base_font_size),
            corner_radius=6,
            fg_color="#FAFAFA",
            text_color="#333333",
            border_color="#B0BEC5"
        )
        address_entry.pack(side="left", fill="x", expand=True, padx=10, pady=5)

        button_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="transparent")
        button_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        button_frame.grid_columnconfigure(tuple(range(5)), weight=1)

        buttons = [
            ("Load from Docx/Pdf (Ctrl+L)", self.file_handler.load_from_documentpdf),
            ("Clear All Fields", self.clear_fields),
            ("Export to PDF (Ctrl+E)", self.file_handler.export_to_pdf),
            ("Save to Word (Ctrl+W)", self.file_handler.save_to_docx),
            ("Send via Email (Ctrl+G)", self.email_sender.send_email),
        ]
        for col, (text, command) in enumerate(buttons):
            btn = ctk.CTkButton(
                button_frame,
                text=text,
                command=command,
                font=("Roboto", base_font_size),
                corner_radius=8,
                fg_color="#2196F3",
                hover_color="#1976D2",
                text_color="#FFFFFF",
                height=35
            )
            btn.grid(row=0, column=col, sticky="ew", padx=5, pady=5)

        self.columns_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="transparent")
        self.columns_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=10)
        self.columns_frame.grid_columnconfigure(0, weight=1)
        self.columns_frame.grid_rowconfigure(0, weight=1)

        self.beg_frame = ctk.CTkFrame(self.columns_frame, corner_radius=8, fg_color="#FFFFFF", border_width=1, border_color="#E0E0E0")
        ctk.CTkLabel(self.beg_frame, text="Beginning Cash Balances", font=("Roboto", int(base_font_size + 2), "bold"), text_color="#333333").pack(anchor="w", padx=10, pady=10)

        self.inflow_frame = ctk.CTkFrame(self.columns_frame, corner_radius=8, fg_color="#FFFFFF", border_width=1, border_color="#E0E0E0")
        ctk.CTkLabel(self.inflow_frame, text="Cash Inflows", font=("Roboto", int(base_font_size + 2), "bold"), text_color="#333333").pack(anchor="w", padx=10, pady=10)

        self.outflow_frame = ctk.CTkFrame(self.columns_frame, corner_radius=8, fg_color="#FFFFFF", border_width=1, border_color="#E0E0E0")
        ctk.CTkLabel(self.outflow_frame, text="Cash Outflows", font=("Roboto", int(base_font_size + 2), "bold"), text_color="#333333").pack(anchor="w", padx=10, pady=10)

        self.end_frame = ctk.CTkFrame(self.columns_frame, corner_radius=8, fg_color="#FFFFFF", border_width=1, border_color="#E0E0E0")
        ctk.CTkLabel(self.end_frame, text="Ending Cash Balances", font=("Roboto", int(base_font_size + 2), "bold"), text_color="#333333").pack(anchor="w", padx=10, pady=10)

        self.totals_frame = ctk.CTkFrame(self.columns_frame, corner_radius=8, fg_color="#FFFFFF", border_width=1, border_color="#E0E0E0")
        ctk.CTkLabel(self.totals_frame, text="Totals", font=("Roboto", int(base_font_size + 2), "bold"), text_color="#333333").pack(anchor="w", padx=10, pady=10)

        self.column_frames = [
            self.beg_frame,
            self.inflow_frame,
            self.outflow_frame,
            self.end_frame,
            self.totals_frame
        ]

        names_frame = ctk.CTkFrame(self.scrollable_frame, corner_radius=8, fg_color="#FFFFFF", border_width=1, border_color="#E0E0E0")
        names_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=10)
        names_frame.grid_columnconfigure(tuple(range(5)), weight=1)

        name_fields = [
            ("Recipients (comma-separated):", self.variables['recipient_emails_var']),
            ("Prepared by (HOA Treasurer):", self.variables['prepared_by_var']),
            ("Noted by (HOA President):", self.variables['noted_by_var_1']),
            ("Noted by (CHUDD HCD-CORDS):", self.variables['noted_by_var_2']),
            ("Checked by (HOA Auditor):", self.variables['checked_by_var'])
        ]

        for col, (label, var) in enumerate(name_fields):
            frame = ctk.CTkFrame(names_frame, fg_color="transparent")
            frame.grid(row=0, column=col, sticky="ew", padx=10, pady=10)
            ctk.CTkLabel(
                frame,
                text=label,
                font=("Roboto", base_font_size),
                text_color="#333333"
            ).pack(side="top", anchor="w")
            entry = ctk.CTkEntry(
                frame,
                textvariable=var,
                width=input_width,
                font=("Roboto", base_font_size),
                corner_radius=6,
                fg_color="#FAFAFA",
                text_color="#333333",
                border_color="#B0BEC5"
            )
            entry.pack(side="top", fill="x", expand=True)

        self.populate_columns()
        self.root.bind("<Configure>", self.debounce_layout)

    def populate_columns(self):
        base_font_size = max(10, int(self.screen_height / 60))
        input_width = min(max(self.min_input_width, int(self.screen_width * 0.15)), self.max_input_width)

        beg_inner = ctk.CTkFrame(self.beg_frame, fg_color="transparent")
        beg_inner.pack(fill="both", expand=True, padx=10, pady=10)
        beg_items = [
            ("Cash in Bank (beginning):", self.variables['cash_bank_beg']),
            ("Cash on Hand (beginning):", self.variables['cash_hand_beg'])
        ]
        for label, var in beg_items:
            item_frame = ctk.CTkFrame(beg_inner, fg_color="transparent")
            item_frame.pack(fill="x", padx=5, pady=5)
            ctk.CTkLabel(
                item_frame,
                text=label,
                font=("Roboto", base_font_size),
                text_color="#333333",
                anchor="w"
            ).pack(side="left", fill="x", expand=True)
            entry = ctk.CTkEntry(
                item_frame,
                textvariable=var,
                width=input_width,
                font=("Roboto", base_font_size),
                corner_radius=6,
                fg_color="#FAFAFA",
                text_color="#333333",
                border_color="#B0BEC5"
            )
            entry.pack(side="right", padx=10)
            self.calculator.format_entry(var, entry)

        inflow_inner = ctk.CTkFrame(self.inflow_frame, fg_color="transparent")
        inflow_inner.pack(fill="both", expand=True, padx=10, pady=10)
        inflow_items = [
            ("Monthly dues collected:", self.variables['monthly_dues']),
            ("Certifications issued:", self.variables['certifications']),
            ("Membership fee:", self.variables['membership_fee']),
            ("Vehicle stickers:", self.variables['vehicle_stickers']),
            ("Rentals:", self.variables['rentals']),
            ("Solicitations/Donations:", self.variables['solicitations']),
            ("Interest Income:", self.variables['interest_income']),
            ("Livelihood Management Fee:", self.variables['livelihood_fee']),
            ("Others:", self.variables['inflows_others'])
        ]
        for label, var in inflow_items:
            item_frame = ctk.CTkFrame(inflow_inner, fg_color="transparent")
            item_frame.pack(fill="x", padx=5, pady=5)
            ctk.CTkLabel(
                item_frame,
                text=label,
                font=("Roboto", base_font_size),
                text_color="#333333",
                anchor="w"
            ).pack(side="left", fill="x", expand=True)
            entry = ctk.CTkEntry(
                item_frame,
                textvariable=var,
                width=input_width,
                font=("Roboto", base_font_size),
                corner_radius=6,
                fg_color="#FAFAFA",
                text_color="#333333",
                border_color="#B0BEC5"
            )
            entry.pack(side="right", padx=10)
            self.calculator.format_entry(var, entry)

        outflow_inner = ctk.CTkFrame(self.outflow_frame, fg_color="transparent")
        outflow_inner.pack(fill="both", expand=True, padx=10, pady=10)
        outflow_items = [
            ("Snacks/Meals for visitors:", self.variables['snacks_meals']),
            ("Transportation expenses:", self.variables['transportation']),
            ("Office supplies expense:", self.variables['office_supplies']),
            ("Printing and photocopy:", self.variables['printing']),
            ("Labor:", self.variables['labor']),
            ("Billboard expense:", self.variables['billboard']),
            ("Clearing/cleaning charges:", self.variables['cleaning']),
            ("Miscellaneous expenses:", self.variables['misc_expenses']),
            ("Federation fee:", self.variables['federation_fee']),
            ("HOA-BOD Uniforms:", self.variables['uniforms']),
            ("BOD Mtg:", self.variables['bod_mtg']),
            ("General Assembly:", self.variables['general_assembly']),
            ("Cash Deposit to bank:", self.variables['cash_deposit']),
            ("Withholding tax:", self.variables['withholding_tax']),
            ("Refund:", self.variables['refund']),
            ("Others:", self.variables['outflows_others'])
        ]
        for label, var in outflow_items:
            item_frame = ctk.CTkFrame(outflow_inner, fg_color="transparent")
            item_frame.pack(fill="x", padx=5, pady=5)
            ctk.CTkLabel(
                item_frame,
                text=label,
                font=("Roboto", base_font_size),
                text_color="#333333",
                anchor="w"
            ).pack(side="left", fill="x", expand=True)
            entry = ctk.CTkEntry(
                item_frame,
                textvariable=var,
                width=input_width,
                font=("Roboto", base_font_size),
                corner_radius=6,
                fg_color="#FAFAFA",
                text_color="#333333",
                border_color="#B0BEC5"
            )
            entry.pack(side="right", padx=10)
            self.calculator.format_entry(var, entry)

        end_inner = ctk.CTkFrame(self.end_frame, fg_color="transparent")
        end_inner.pack(fill="both", expand=True, padx=10, pady=10)
        end_items = [
            ("Cash in Bank:", self.variables['ending_cash_bank']),
            ("Cash on Hand:", self.variables['ending_cash_hand'])
        ]
        for label, var in end_items:
            item_frame = ctk.CTkFrame(end_inner, fg_color="transparent")
            item_frame.pack(fill="x", padx=5, pady=5)
            ctk.CTkLabel(
                item_frame,
                text=label,
                font=("Roboto", base_font_size),
                text_color="#333333",
                anchor="w"
            ).pack(side="left", fill="x", expand=True)
            entry = ctk.CTkEntry(
                item_frame,
                textvariable=var,
                width=input_width,
                font=("Roboto", base_font_size),
                state="disabled",
                corner_radius=6,
                fg_color="#ECEFF1",
                text_color="#333333",
                border_color="#B0BEC5"
            )
            entry.pack(side="right", padx=10)

        total_inner = ctk.CTkFrame(self.totals_frame, fg_color="transparent")
        total_inner.pack(fill="both", expand=True, padx=10, pady=10)
        total_items = [
            ("Total Cash Receipts:", self.variables['total_receipts']),
            ("Cash Outflows:", self.variables['cash_outflows']),
            ("Ending Cash Balance:", self.variables['ending_cash'])
        ]
        for label, var in total_items:
            item_frame = ctk.CTkFrame(total_inner, fg_color="transparent")
            item_frame.pack(fill="x", padx=5, pady=5)
            ctk.CTkLabel(
                item_frame,
                text=label,
                font=("Roboto", base_font_size),
                text_color="#333333",
                anchor="w"
            ).pack(side="left", fill="x", expand=True)
            entry = ctk.CTkEntry(
                item_frame,
                textvariable=var,
                width=input_width,
                font=("Roboto", base_font_size),
                state="disabled",
                corner_radius=6,
                fg_color="#ECEFF1",
                text_color="#333333",
                border_color="#B0BEC5"
            )
            entry.pack(side="right", padx=10)

    def debounce_layout(self, event=None):
        if self.debounce_id:
            self.root.after_cancel(self.debounce_id)
        self.debounce_id = self.root.after(int(self.layout_debounce_delay * 1000), self.update_layout)

    def update_layout(self):
        window_width = self.main_frame.winfo_width()
        window_height = self.main_frame.winfo_height()

        num_columns = max(1, window_width // self.min_column_width)
        num_columns = min(num_columns, len(self.column_frames))

        if window_width < self.min_column_width * 1.5:
            num_columns = 1

        if self.last_num_columns == num_columns:
            return

        self.last_num_columns = num_columns

        for frame in self.column_frames:
            frame.grid_forget()

        column_width = min(max(self.min_column_width, window_width // num_columns), self.max_column_width)

        for i, frame in enumerate(self.column_frames):
            row = i // num_columns
            col = i % num_columns
            frame.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)
            frame.configure(width=column_width)

        self.columns_frame.grid_columnconfigure(tuple(range(num_columns)), weight=1, uniform="column")
        self.columns_frame.grid_rowconfigure(tuple(range((len(self.column_frames) + num_columns - 1) // num_columns)), weight=1)

        self.canvas.itemconfig(self.canvas_window, width=window_width)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def clear_fields(self):
        for var in self.variables.values():
            var.set("")
        messagebox.showinfo("Success", "All fields have been cleared")