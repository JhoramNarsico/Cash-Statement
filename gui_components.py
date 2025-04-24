import tkinter as tk
from tkinter import messagebox
import datetime
from hover_calendar import HoverCalendar

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
        
        # Get screen dimensions for relative sizing
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        
        self.create_widgets()
        self.setup_keyboard_shortcuts()
        self.date_var.trace('w', self.update_display_date)

    def update_display_date(self, *args):
        """Convert mm/dd/yyyy from date_var to MMM dd, yyyy for display_date."""
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
        tooltip = tk.Toplevel(widget)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry("+1000+1000")
        font_size = max(9, int(self.screen_height / 80 * 0.85))  # Scaled to 85%
        label = tk.Label(tooltip, text=text, background="#FFFFE0", relief="solid", borderwidth=1, font=("Arial", font_size))
        label.pack()

        def show(event):
            x = widget.winfo_rootx() + 17  # Scaled from 20
            y = widget.winfo_rooty() + 17  # Scaled from 20
            tooltip.wm_geometry(f"+{x}+{y}")
            tooltip.deiconify()

        def hide(event):
            tooltip.withdraw()

        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)
        tooltip.withdraw()

    def show_calendar(self):
        """Show a standalone calendar in a popup window, sized relative to screen."""
        popup = tk.Toplevel(self.root)
        popup.title("Select Date")
        
        popup_width = max(255, min(510, int(self.screen_width * 0.3 * 0.85)))  # Scaled to 85%
        popup_height = max(255, min(510, int(self.screen_height * 0.3 * 0.85)))  # Scaled to 85%
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

        font_size = max(10, int(self.screen_height / 60 * 0.85))  # Scaled to 85%
        cal = HoverCalendar(
            popup,
            font=("Arial", font_size)
        )
        cal.pack(padx=17, pady=17, fill="both", expand=True)  # Scaled padx, pady

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

        confirm_button = tk.Button(
            popup,
            text="Select",
            command=on_date_select,
            font=("Arial", font_size),
            width=int(self.screen_width * 0.008 * 0.85),  # Scaled to 85%
            padx=5,  # Internal padding
            pady=2   # Internal padding
        )
        confirm_button.pack(pady=17)  # Scaled pady

        popup.deiconify()
        popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    def create_widgets(self):
        # Main frame
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=13, pady=13)  # Scaled padx, pady from 15

        # Scrollable frame using Canvas and Scrollbar
        self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = tk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Configure grid weights
        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        self.scrollable_frame.grid_rowconfigure((0, 1, 2, 3), weight=0)

        # Dynamic font size
        base_font_size = max(10, int(self.screen_height / 60 * 0.85))  # Scaled to 85%

        # Date Frame
        header_frame = tk.Frame(self.scrollable_frame, relief="groove", borderwidth=2)
        header_frame.grid(row=0, column=0, sticky="ew", padx=9, pady=7)  # Scaled padx, pady
        header_frame.grid_columnconfigure(0, weight=1)

        date_frame = tk.Frame(header_frame)
        date_frame.pack(side="left", padx=7, anchor="w")  # Scaled padx
        tk.Label(date_frame, text="Date:", font=("Arial", base_font_size)).pack(side="left")
        date_button = tk.Button(
            date_frame,
            textvariable=self.display_date,
            font=("Arial", base_font_size),
            command=self.show_calendar
        )
        date_button.pack(side="left", padx=4)  # Scaled padx
        self.create_tooltip(date_button, "Click to select a date from the calendar")

        # Email and Names Configuration Frame
        email_frame = tk.Frame(self.scrollable_frame, relief="groove", borderwidth=2)
        email_frame.grid(row=1, column=0, sticky="ew", padx=9, pady=7)  # Scaled padx, pady
        email_frame.grid_columnconfigure(0, weight=1)

        # Recipients Field
        tk.Label(email_frame, text="Recipients (comma-separated):", font=("Arial", base_font_size)).pack(side="left", padx=7)  # Scaled padx
        email_entry = tk.Entry(
            email_frame,
            textvariable=self.variables['recipient_emails_var'],
            width=int(self.screen_width * 0.015 * 0.85),  # Scaled to 85%
            font=("Arial", base_font_size)
        )
        email_entry.pack(side="left", padx=7)  # Scaled padx

        # Prepared by Field
        tk.Label(email_frame, text="Prepared by (HOA Treasurer):", font=("Arial", base_font_size)).pack(side="left", padx=7)  # Scaled padx
        prepared_entry = tk.Entry(
            email_frame,
            textvariable=self.variables['prepared_by_var'],
            width=int(self.screen_width * 0.012 * 0.85),  # Scaled to 85%
            font=("Arial", base_font_size)
        )
        prepared_entry.pack(side="left", padx=7)  # Scaled padx

        # Noted by Fields (Two)
        tk.Label(email_frame, text="Noted by (HOA President):", font=("Arial", base_font_size)).pack(side="left", padx=7)  # Scaled padx
        noted_entry_1 = tk.Entry(
            email_frame,
            textvariable=self.variables['noted_by_var_1'],
            width=int(self.screen_width * 0.012 * 0.85),  # Scaled to 85%
            font=("Arial", base_font_size)
        )
        noted_entry_1.pack(side="left", padx=7)  # Scaled padx

        tk.Label(email_frame, text="Noted by (CHUDD HCD-CORDS):", font=("Arial", base_font_size)).pack(side="left", padx=7)  # Scaled padx
        noted_entry_2 = tk.Entry(
            email_frame,
            textvariable=self.variables['noted_by_var_2'],
            width=int(self.screen_width * 0.012 * 0.85),  # Scaled to 85%
            font=("Arial", base_font_size)
        )
        noted_entry_2.pack(side="left", padx=7)  # Scaled padx

        # Checked by Field
        tk.Label(email_frame, text="Checked by (HOA Auditor):", font=("Arial", base_font_size)).pack(side="left", padx=7)  # Scaled padx
        checked_entry = tk.Entry(
            email_frame,
            textvariable=self.variables['checked_by_var'],
            width=int(self.screen_width * 0.012 * 0.85),  # Scaled to 85%
            font=("Arial", base_font_size)
        )
        checked_entry.pack(side="left", padx=7)  # Scaled padx

        # Buttons Frame
        button_frame = tk.Frame(self.scrollable_frame)
        button_frame.grid(row=2, column=0, sticky="ew", padx=9, pady=7)  # Scaled padx, pady
        button_frame.grid_columnconfigure(0, weight=1)

        buttons = [
            ("Load from Docx/Pdf (Ctrl+L)", self.file_handler.load_from_documentpdf),
            ("Clear All Fields", self.clear_fields),
            ("Export to PDF (Ctrl+E)", self.file_handler.export_to_pdf),
            ("Save to Word (Ctrl+W)", self.file_handler.save_to_docx),
            ("Send via Email (Ctrl+G)", self.email_sender.send_email),
        ]
        for text, command in buttons:
            tk.Button(
                button_frame,
                text=text,
                command=command,
                font=("Arial", base_font_size),
                width=int(self.screen_width * 0.015 * 0.85),  # Scaled to 85%
                padx=5,  # Internal padding
                pady=2   # Internal padding
            ).pack(side="left", padx=8, pady=3)  # Scaled padx, pady

        # Columns Frame
        self.columns_frame = tk.Frame(self.scrollable_frame)
        self.columns_frame.grid(row=3, column=0, sticky="nsew", padx=9, pady=7)  # Scaled padx, pady

        # Define column frames
        self.beg_frame = tk.Frame(self.columns_frame, relief="groove", borderwidth=2)
        tk.Label(self.beg_frame, text="Beginning Cash Balances", font=("Arial", int(base_font_size + 2 * 0.85), "bold")).pack(anchor="w", padx=9, pady=7)  # Scaled font, padx, pady

        self.inflow_frame = tk.Frame(self.columns_frame, relief="groove", borderwidth=2)
        tk.Label(self.inflow_frame, text="Cash Inflows", font=("Arial", int(base_font_size + 2 * 0.85), "bold")).pack(anchor="w", padx=9, pady=7)  # Scaled font, padx, pady

        self.outflow_frame = tk.Frame(self.columns_frame, relief="groove", borderwidth=2)
        tk.Label(self.outflow_frame, text="Cash Outflows", font=("Arial", int(base_font_size + 2 * 0.85), "bold")).pack(anchor="w", padx=9, pady=7)  # Scaled font, padx, pady

        self.end_frame = tk.Frame(self.columns_frame, relief="groove", borderwidth=2)
        tk.Label(self.end_frame, text="Ending Cash Balances", font=("Arial", int(base_font_size + 2 * 0.85), "bold")).pack(anchor="w", padx=9, pady=7)  # Scaled font, padx, pady

        self.totals_frame = tk.Frame(self.columns_frame, relief="groove", borderwidth=2)
        tk.Label(self.totals_frame, text="Totals", font=("Arial", int(base_font_size + 2 * 0.85), "bold")).pack(anchor="w", padx=9, pady=7)  # Scaled font, padx, pady

        self.column_frames = [
            self.beg_frame,
            self.inflow_frame,
            self.outflow_frame,
            self.end_frame,
            self.totals_frame
        ]

        self.populate_columns()
        self.root.bind("<Configure>", self.update_layout)

    def populate_columns(self):
        base_font_size = max(10, int(self.screen_height / 60 * 0.85))  # Scaled to 85%
        
        # Beginning Balances
        beg_inner = tk.Frame(self.beg_frame)
        beg_inner.pack(fill="x", padx=9, pady=7)  # Scaled padx, pady
        beg_items = [
            ("Cash in Bank (beginning):", self.variables['cash_bank_beg']),
            ("Cash on Hand (beginning):", self.variables['cash_hand_beg'])
        ]
        for i, (label, var) in enumerate(beg_items):
            tk.Label(beg_inner, text=label, font=("Arial", base_font_size), anchor="w").grid(row=i, column=0, sticky="w", padx=7, pady=3)  # Scaled padx, pady
            entry = tk.Entry(
                beg_inner,
                textvariable=var,
                width=int(self.screen_width * 0.01 * 0.85),  # Scaled to 85%
                font=("Arial", base_font_size)
            )
            entry.grid(row=i, column=1, sticky="e", padx=7, pady=3)  # Scaled padx, pady
            self.calculator.format_entry(var, entry)

        # Cash Inflows
        inflow_inner = tk.Frame(self.inflow_frame)
        inflow_inner.pack(fill="x", padx=9, pady=7)  # Scaled padx, pady
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
        for i, (label, var) in enumerate(inflow_items):
            tk.Label(inflow_inner, text=label, font=("Arial", base_font_size), anchor="w").grid(row=i, column=0, sticky="w", padx=7, pady=3)  # Scaled padx, pady
            entry = tk.Entry(
                inflow_inner,
                textvariable=var,
                width=int(self.screen_width * 0.01 * 0.85),  # Scaled to 85%
                font=("Arial", base_font_size)
            )
            entry.grid(row=i, column=1, sticky="e", padx=7, pady=3)  # Scaled padx, pady
            self.calculator.format_entry(var, entry)

        # Cash Outflows
        outflow_inner = tk.Frame(self.outflow_frame)
        outflow_inner.pack(fill="x", padx=9, pady=7)  # Scaled padx, pady
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
        for i, (label, var) in enumerate(outflow_items):
            tk.Label(outflow_inner, text=label, font=("Arial", base_font_size), anchor="w").grid(row=i, column=0, sticky="w", padx=7, pady=3)  # Scaled padx, pady
            entry = tk.Entry(
                outflow_inner,
                textvariable=var,
                width=int(self.screen_width * 0.01 * 0.85),  # Scaled to 85%
                font=("Arial", base_font_size)
            )
            entry.grid(row=i, column=1, sticky="e", padx=7, pady=3)  # Scaled padx, pady
            self.calculator.format_entry(var, entry)

        # Ending Balances
        end_inner = tk.Frame(self.end_frame)
        end_inner.pack(fill="x", padx=9, pady=7)  # Scaled padx, pady
        end_items = [
            ("Cash in Bank:", self.variables['ending_cash_bank']),
            ("Cash on Hand:", self.variables['ending_cash_hand'])
        ]
        for i, (label, var) in enumerate(end_items):
            tk.Label(end_inner, text=label, font=("Arial", base_font_size), anchor="w").grid(row=i, column=0, sticky="w", padx=7, pady=3)  # Scaled padx, pady
            entry = tk.Entry(
                end_inner,
                textvariable=var,
                width=int(self.screen_width * 0.01 * 0.85),  # Scaled to 85%
                font=("Arial", base_font_size),
                state="disabled"
            )
            entry.grid(row=i, column=1, sticky="e", padx=7, pady=3)  # Scaled padx, pady

        # Totals
        total_inner = tk.Frame(self.totals_frame)
        total_inner.pack(fill="x", padx=9, pady=7)  # Scaled padx, pady
        total_items = [
            ("Total Cash Receipts:", self.variables['total_receipts']),
            ("Cash Outflows:", self.variables['cash_outflows']),
            ("Ending Cash Balance:", self.variables['ending_cash'])
        ]
        for i, (label, var) in enumerate(total_items):
            tk.Label(total_inner, text=label, font=("Arial", base_font_size), anchor="w").grid(row=i, column=0, sticky="w", padx=7, pady=3)  # Scaled padx, pady
            entry = tk.Entry(
                total_inner,
                textvariable=var,
                width=int(self.screen_width * 0.01 * 0.85),  # Scaled to 85%
                font=("Arial", base_font_size),
                state="disabled"
            )
            entry.grid(row=i, column=1, sticky="e", padx=7, pady=3)  # Scaled padx, pady

    def update_layout(self, event=None):
        window_width = self.main_frame.winfo_width()
        min_column_width = int(self.screen_width * 0.2 * 0.85)  # Scaled to 85%
        num_columns = max(1, window_width // min_column_width)
        num_columns = min(num_columns, len(self.column_frames))

        if window_width < min_column_width * 2:
            num_columns = 1

        for frame in self.column_frames:
            frame.grid_forget()

        for i, frame in enumerate(self.column_frames):
            row = i // num_columns
            col = i % num_columns
            frame.grid(row=row, column=col, sticky="nsew", padx=7, pady=7)  # Scaled padx, pady

        self.columns_frame.grid_columnconfigure(tuple(range(num_columns)), weight=1, uniform="column")
        self.columns_frame.grid_rowconfigure(tuple(range((len(self.column_frames) + num_columns - 1) // num_columns)), weight=1)

    def clear_fields(self):
        input_vars = [
            self.variables['cash_bank_beg'], self.variables['cash_hand_beg'], self.variables['monthly_dues'],
            self.variables['certifications'], self.variables['membership_fee'], self.variables['vehicle_stickers'],
            self.variables['rentals'], self.variables['solicitations'], self.variables['interest_income'],
            self.variables['livelihood_fee'], self.variables['inflows_others'], self.variables['snacks_meals'],
            self.variables['transportation'], self.variables['office_supplies'], self.variables['printing'],
            self.variables['labor'], self.variables['billboard'], self.variables['cleaning'], self.variables['misc_expenses'],
            self.variables['federation_fee'], self.variables['uniforms'], self.variables['bod_mtg'],
            self.variables['general_assembly'], self.variables['cash_deposit'], self.variables['withholding_tax'],
            self.variables['refund'], self.variables['outflows_others'], self.variables['ending_cash_bank'],
            self.variables['ending_cash_hand'], self.variables['prepared_by_var'], 
            self.variables['noted_by_var_1'], self.variables['noted_by_var_2'], self.variables['checked_by_var']
        ]
        for var in input_vars:
            var.set("")
        self.variables['total_receipts'].set("")
        self.variables['cash_outflows'].set("")
        self.variables['ending_cash'].set("")
        messagebox.showinfo("Success", "All fields have been cleared")