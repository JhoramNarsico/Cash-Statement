import tkinter as tk
import customtkinter as ctk
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
        # Updated color scheme for light theme
        self.primary_color = "#F5F7FA"
        self.secondary_color = "#E8ECEF"
        self.accent_color = "#007BFF"
        self.text_color = "#212529"
        self.success_color = "#28A745"
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
        label = tk.Label(tooltip, text=text, background="#FFF3CD", relief="solid", borderwidth=1, fg="#212529", font=("Roboto", 10))
        label.pack()

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
        """Show a standalone calendar in a popup window, appearing directly in the center."""
        popup = tk.Toplevel(self.root)
        popup.title("Select Date")
        popup.geometry("400x400")
        popup.transient(self.root)
        popup.grab_set()
        popup.withdraw()

        popup.update_idletasks()
        popup_width = 400
        popup_height = 400
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()

        x = main_x + (main_width - popup_width) // 2
        y = main_y + (main_height - popup_height) // 2
        popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")

        cal = HoverCalendar(
            popup,
            background="#F5F7FA",
            foreground="#212529",
            selectbackground="#007BFF",
            selectforeground="#FFFFFF",
            normalbackground="#FFFFFF",
            normalforeground="#212529",
            weekendbackground="#E8ECEF",
            weekendforeground="#212529",
            headersbackground="#E8ECEF",
            headersforeground="#212529",
            showothermonthdays=False,
            showweeknumbers=False
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
            text="Select",
            command=on_date_select,
            fg_color=self.accent_color,
            text_color="#FFFFFF",
            hover_color="#0056B3",
            font=("Roboto", 14),
            width=120,
            height=40
        )
        confirm_button.pack(pady=20)

        popup.deiconify()
        popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    def create_widgets(self):
        self.main_frame = ctk.CTkFrame(self.root, fg_color=self.primary_color)
        self.main_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # Create a canvas for horizontal scrolling
        self.canvas = tk.Canvas(self.main_frame, bg=self.primary_color, highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)

        # Add a horizontal scrollbar
        self.h_scrollbar = tk.Scrollbar(self.main_frame, orient="horizontal", command=self.canvas.xview)
        self.h_scrollbar.pack(side="bottom", fill="x")

        # Configure canvas to use the scrollbar
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set)

        # Create a frame inside the canvas to hold the scrollable frame
        self.canvas_frame = ctk.CTkFrame(self.canvas, fg_color=self.primary_color)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.canvas_frame, anchor="nw")

        # Create the scrollable frame inside the canvas frame
        self.scrollable_frame = ctk.CTkScrollableFrame(self.canvas_frame, fg_color=self.primary_color)
        self.scrollable_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Bind canvas to update scroll region
        def update_scroll_region(event=None):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        self.canvas_frame.bind("<Configure>", update_scroll_region)

        # Bind mouse wheel for horizontal scrolling
        def on_mouse_wheel(event):
            if event.delta:
                self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                self.canvas.xview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.xview_scroll(1, "units")

        self.canvas.bind_all("<MouseWheel>", on_mouse_wheel)  # Windows
        self.canvas.bind_all("<Button-4>", on_mouse_wheel)   # Linux
        self.canvas.bind_all("<Button-5>", on_mouse_wheel)   # Linux

        # Configure grid weights for responsiveness
        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        self.scrollable_frame.grid_rowconfigure((0, 1, 2, 3), weight=0)

        # Date Frame
        header_frame = ctk.CTkFrame(self.scrollable_frame, fg_color=self.secondary_color, corner_radius=8)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        header_frame.grid_columnconfigure(0, weight=1)

        date_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        date_frame.pack(side="left", padx=8, anchor="w")
        ctk.CTkLabel(date_frame, text="Date:", font=("Roboto", 13), text_color=self.text_color).pack(side="left")
        date_button = ctk.CTkButton(
            date_frame,
            textvariable=self.display_date,
            font=("Roboto", 13),
            fg_color="#FFFFFF",
            text_color=self.text_color,
            border_width=1,
            border_color="#CED4DA",
            command=self.show_calendar,
            width=120,
            height=32,
            corner_radius=6,
            hover_color="#E8ECEF"
        )
        date_button.pack(side="left", padx=5)
        self.create_tooltip(date_button, "Click to select a date from the calendar")

        # Email and Names Configuration Frame
        email_frame = ctk.CTkFrame(self.scrollable_frame, fg_color=self.secondary_color, corner_radius=8)
        email_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=8)
        email_frame.grid_columnconfigure(0, weight=1)

        # Recipients Field (Adjusted width from 320 to 200)
        ctk.CTkLabel(email_frame, text="Recipients (comma-separated):", font=("Roboto", 13), text_color=self.text_color).pack(side="left", padx=8)
        email_entry = ctk.CTkEntry(email_frame, textvariable=self.variables['recipient_emails_var'], width=200, font=("Roboto", 13), fg_color="#FFFFFF", text_color=self.text_color, border_color="#CED4DA")
        email_entry.pack(side="left", padx=8)

        # Prepared by Field
        ctk.CTkLabel(email_frame, text="Prepared by (HOA Treasurer):", font=("Roboto", 13), text_color=self.text_color).pack(side="left", padx=8)
        prepared_entry = ctk.CTkEntry(email_frame, textvariable=self.variables['prepared_by_var'], width=160, font=("Roboto", 13), fg_color="#FFFFFF", text_color=self.text_color, border_color="#CED4DA")
        prepared_entry.pack(side="left", padx=8)

        # Noted by Fields (Two)
        ctk.CTkLabel(email_frame, text="Noted by (HOA President):", font=("Roboto", 13), text_color=self.text_color).pack(side="left", padx=8)
        noted_entry_1 = ctk.CTkEntry(email_frame, textvariable=self.variables['noted_by_var_1'], width=160, font=("Roboto", 13), fg_color="#FFFFFF", text_color=self.text_color, border_color="#CED4DA")
        noted_entry_1.pack(side="left", padx=8)

        ctk.CTkLabel(email_frame, text="Noted by (CHUDD HCD-CORDS):", font=("Roboto", 13), text_color=self.text_color).pack(side="left", padx=8)
        noted_entry_2 = ctk.CTkEntry(email_frame, textvariable=self.variables['noted_by_var_2'], width=160, font=("Roboto", 13), fg_color="#FFFFFF", text_color=self.text_color, border_color="#CED4DA")
        noted_entry_2.pack(side="left", padx=8)

        # Checked by Field
        ctk.CTkLabel(email_frame, text="Checked by (HOA Auditor):", font=("Roboto", 13), text_color=self.text_color).pack(side="left", padx=8)
        checked_entry = ctk.CTkEntry(email_frame, textvariable=self.variables['checked_by_var'], width=160, font=("Roboto", 13), fg_color="#FFFFFF", text_color=self.text_color, border_color="#CED4DA")
        checked_entry.pack(side="left", padx=8)

        # Buttons Frame
        button_frame = ctk.CTkFrame(self.scrollable_frame, fg_color=self.primary_color)
        button_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=8)
        button_frame.grid_columnconfigure(0, weight=1)

        buttons = [
            ("Load from Docx/Pdf (Ctrl+L)", self.file_handler.load_from_documentpdf),
            ("Clear All Fields", self.clear_fields),
            ("Export to PDF (Ctrl+E)", self.file_handler.export_to_pdf),
            ("Save to Word (Ctrl+W)", self.file_handler.save_to_docx),
            ("Send via Email (Ctrl+G)", self.email_sender.send_email),
        ]
        for text, command in buttons:
            ctk.CTkButton(
                button_frame,
                text=text,
                command=command,
                font=("Roboto", 13),
                fg_color=self.accent_color,
                hover_color="#0056B3",
                text_color="#FFFFFF",
                width=160,
                height=36,
                corner_radius=6
            ).pack(side="left", padx=8, pady=5)

        # Columns Frame for responsive layout
        self.columns_frame = ctk.CTkFrame(self.scrollable_frame, fg_color=self.primary_color)
        self.columns_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=8)

        # Define column frames
        self.beg_frame = ctk.CTkFrame(self.columns_frame, fg_color=self.secondary_color, corner_radius=8)
        ctk.CTkLabel(self.beg_frame, text="Beginning Cash Balances", font=("Roboto", 15, "bold"), text_color=self.text_color).pack(anchor="w", padx=10, pady=8)

        self.inflow_frame = ctk.CTkFrame(self.columns_frame, fg_color=self.secondary_color, corner_radius=8)
        ctk.CTkLabel(self.inflow_frame, text="Cash Inflows", font=("Roboto", 15, "bold"), text_color=self.text_color).pack(anchor="w", padx=10, pady=8)

        self.outflow_frame = ctk.CTkFrame(self.columns_frame, fg_color=self.secondary_color, corner_radius=8)
        ctk.CTkLabel(self.outflow_frame, text="Cash Outflows", font=("Roboto", 15, "bold"), text_color=self.text_color).pack(anchor="w", padx=10, pady=8)

        self.end_frame = ctk.CTkFrame(self.columns_frame, fg_color=self.secondary_color, corner_radius=8)
        ctk.CTkLabel(self.end_frame, text="Ending Cash Balances", font=("Roboto", 15, "bold"), text_color=self.text_color).pack(anchor="w", padx=10, pady=8)

        self.totals_frame = ctk.CTkFrame(self.columns_frame, fg_color=self.secondary_color, corner_radius=8)
        ctk.CTkLabel(self.totals_frame, text="Totals", font=("Roboto", 15, "bold"), text_color=self.text_color).pack(anchor="w", padx=10, pady=8)

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
        # Beginning Balances
        beg_inner = ctk.CTkFrame(self.beg_frame, fg_color="transparent")
        beg_inner.pack(fill="x", padx=10, pady=8)
        beg_items = [
            ("Cash in Bank (beginning):", self.variables['cash_bank_beg']),
            ("Cash on Hand (beginning):", self.variables['cash_hand_beg'])
        ]
        for i, (label, var) in enumerate(beg_items):
            ctk.CTkLabel(beg_inner, text=label, font=("Roboto", 12), text_color=self.text_color, anchor="w").grid(row=i, column=0, sticky="w", padx=8, pady=4)
            entry = ctk.CTkEntry(beg_inner, textvariable=var, width=140, font=("Roboto", 12), fg_color="#FFFFFF", text_color=self.text_color, border_color="#CED4DA")
            entry.grid(row=i, column=1, sticky="e", padx=8, pady=4)
            self.calculator.format_entry(var, entry)

        # Cash Inflows
        inflow_inner = ctk.CTkFrame(self.inflow_frame, fg_color="transparent")
        inflow_inner.pack(fill="x", padx=10, pady=8)
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
            ctk.CTkLabel(inflow_inner, text=label, font=("Roboto", 12), text_color=self.text_color, anchor="w").grid(row=i, column=0, sticky="w", padx=8, pady=4)
            entry = ctk.CTkEntry(inflow_inner, textvariable=var, width=140, font=("Roboto", 12), fg_color="#FFFFFF", text_color=self.text_color, border_color="#CED4DA")
            entry.grid(row=i, column=1, sticky="e", padx=8, pady=4)
            self.calculator.format_entry(var, entry)

        # Cash Outflows
        outflow_inner = ctk.CTkFrame(self.outflow_frame, fg_color="transparent")
        outflow_inner.pack(fill="x", padx=10, pady=8)
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
            ctk.CTkLabel(outflow_inner, text=label, font=("Roboto", 12), text_color=self.text_color, anchor="w").grid(row=i, column=0, sticky="w", padx=8, pady=4)
            entry = ctk.CTkEntry(outflow_inner, textvariable=var, width=140, font=("Roboto", 12), fg_color="#FFFFFF", text_color=self.text_color, border_color="#CED4DA")
            entry.grid(row=i, column=1, sticky="e", padx=8, pady=4)
            self.calculator.format_entry(var, entry)

        # Ending Balances
        end_inner = ctk.CTkFrame(self.end_frame, fg_color="transparent")
        end_inner.pack(fill="x", padx=10, pady=8)
        end_items = [
            ("Cash in Bank:", self.variables['ending_cash_bank']),
            ("Cash on Hand:", self.variables['ending_cash_hand'])
        ]
        for i, (label, var) in enumerate(end_items):
            ctk.CTkLabel(end_inner, text=label, font=("Roboto", 12), text_color=self.text_color, anchor="w").grid(row=i, column=0, sticky="w", padx=8, pady=4)
            entry = ctk.CTkEntry(end_inner, textvariable=var, width=140, font=("Roboto", 12), fg_color="#FFFFFF", text_color=self.text_color, border_color="#CED4DA", state="disabled")
            entry.grid(row=i, column=1, sticky="e", padx=8, pady=4)

        # Totals
        total_inner = ctk.CTkFrame(self.totals_frame, fg_color="transparent")
        total_inner.pack(fill="x", padx=10, pady=8)
        total_items = [
            ("Total Cash Receipts:", self.variables['total_receipts']),
            ("Cash Outflows:", self.variables['cash_outflows']),
            ("Ending Cash Balance:", self.variables['ending_cash'])
        ]
        for i, (label, var) in enumerate(total_items):
            ctk.CTkLabel(total_inner, text=label, font=("Roboto", 12), text_color=self.text_color, anchor="w").grid(row=i, column=0, sticky="w", padx=8, pady=4)
            entry = ctk.CTkEntry(total_inner, textvariable=var, width=140, font=("Roboto", 12), fg_color="#FFFFFF", text_color=self.text_color, border_color="#CED4DA", state="disabled")
            entry.grid(row=i, column=1, sticky="e", padx=8, pady=4)

    def update_layout(self, event=None):
        window_width = self.main_frame.winfo_width()
        min_column_width = 300
        num_columns = max(1, window_width // min_column_width)
        num_columns = min(num_columns, len(self.column_frames))

        for frame in self.column_frames:
            frame.grid_forget()

        for i, frame in enumerate(self.column_frames):
            row = i // num_columns
            col = i % num_columns
            frame.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)

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