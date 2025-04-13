import os
import datetime
import sys
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import base64
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
from decimal import Decimal
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from docx import Document
from docx.shared import Inches
import customtkinter as ctk
from tkcalendar import Calendar

# Gmail API setup
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.pickle'
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
SENDER_EMAIL = 'chuddcdo@gmail.com'

class HoverCalendar(Calendar):
    """Custom Calendar class to enable hovering over month and year for navigation."""
    def __init__(self, master=None, **kw):
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

class IntegratedCashFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cash Flow Statement")
        
        # Maximize the window
        self.root.state('zoomed')  # Maximizes the window on Windows
        
        # Set CustomTkinter appearance
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Color scheme
        self.primary_color = "#1C2526"  # Dark background
        self.secondary_color = "#2A3F4D"  # Section backgrounds
        self.accent_color = "#00A7E1"  # Buttons
        self.text_color = "#E0E0E0"  # Light grey text
        self.success_color = "#4CAF50"  # Success indicators
        
        # Email variable
        self.recipient_emails_var = ctk.StringVar()
        
        # Cash flow variables
        self.title_var = ctk.StringVar(value="Statement Of Cash Flows")
        self.today_date = ctk.StringVar(value=datetime.datetime.now().strftime("%m/%d/%Y"))
        self.display_date = ctk.StringVar(value=datetime.datetime.now().strftime("%b %d, %Y"))
        
        self.cash_bank_beg = ctk.StringVar()
        self.cash_hand_beg = ctk.StringVar()
        self.monthly_dues = ctk.StringVar()
        self.certifications = ctk.StringVar()
        self.membership_fee = ctk.StringVar()
        self.vehicle_stickers = ctk.StringVar()
        self.rentals = ctk.StringVar()
        self.solicitations = ctk.StringVar()
        self.interest_income = ctk.StringVar()
        self.livelihood_fee = ctk.StringVar()
        self.inflows_others = ctk.StringVar()
        self.total_receipts = ctk.StringVar()
        self.cash_outflows = ctk.StringVar()
        self.snacks_meals = ctk.StringVar()
        self.transportation = ctk.StringVar()
        self.office_supplies = ctk.StringVar()
        self.printing = ctk.StringVar()
        self.labor = ctk.StringVar()
        self.billboard = ctk.StringVar()
        self.cleaning = ctk.StringVar()
        self.misc_expenses = ctk.StringVar()
        self.federation_fee = ctk.StringVar()
        self.uniforms = ctk.StringVar()
        self.bod_mtg = ctk.StringVar()
        self.general_assembly = ctk.StringVar()
        self.cash_deposit = ctk.StringVar()
        self.withholding_tax = ctk.StringVar()
        self.refund_sericulture = ctk.StringVar()
        self.outflows_others = ctk.StringVar()
        self.outflows_others_2 = ctk.StringVar()
        self.ending_cash = ctk.StringVar()
        self.ending_cash_bank = ctk.StringVar()
        self.ending_cash_hand = ctk.StringVar()
        
        # Sync display_date when today_date changes
        self.today_date.trace('w', self.update_display_date)
        
        self.create_widgets()
        self.setup_keyboard_shortcuts()

    def update_display_date(self, *args):
        """Convert mm/dd/yyyy from today_date to MMM dd, yyyy for display_date."""
        raw_date = self.today_date.get()
        try:
            date_obj = datetime.datetime.strptime(raw_date, "%m/%d/%Y")
            self.display_date.set(date_obj.strftime("%b %d, %Y"))
        except ValueError:
            pass

    def create_widgets(self):
        # Main container (non-scrollable, using grid layout)
        main_container = ctk.CTkFrame(self.root, fg_color=self.primary_color)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)
        main_container.grid_columnconfigure((0, 1, 2), weight=1)  # Three columns with equal weight
        main_container.grid_rowconfigure((0, 1, 2, 3, 4), weight=0)  # Rows for sections

        # Title and Date Frame (Row 0, spans all columns)
        title_frame = ctk.CTkFrame(main_container, fg_color=self.secondary_color, corner_radius=10)
        title_frame.grid(row=0, column=0, columnspan=3, sticky="ew", padx=10, pady=5)
        
        ctk.CTkLabel(title_frame, text="Title:", font=("Roboto", 12), text_color=self.text_color).pack(side="left", padx=5)
        title_entry = ctk.CTkEntry(title_frame, textvariable=self.title_var, width=250, font=("Roboto", 12), fg_color="#3A4F5D", text_color=self.text_color)
        title_entry.pack(side="left", padx=5)
        
        ctk.CTkLabel(title_frame, text="Date:", font=("Roboto", 12), text_color=self.text_color).pack(side="left", padx=(10, 5))
        date_button = ctk.CTkButton(
            title_frame,
            textvariable=self.display_date,
            font=("Arial", 12),
            fg_color="#3A4F5D",
            text_color=self.text_color,
            command=self.show_calendar,
            width=100,
            height=28,
            corner_radius=5,
            hover_color="#4A5F6D"
        )
        date_button.pack(side="left", padx=5)

        # Email Recipients Frame (Row 1, spans all columns)
        email_frame = ctk.CTkFrame(main_container, fg_color=self.secondary_color, corner_radius=10)
        email_frame.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=5)
        
        ctk.CTkLabel(email_frame, text="Recipient Emails (comma-separated):", font=("Roboto", 12), text_color=self.text_color).pack(side="left", padx=5, pady=5)
        email_entry = ctk.CTkEntry(email_frame, textvariable=self.recipient_emails_var, width=350, font=("Roboto", 12), fg_color="#3A4F5D", text_color=self.text_color)
        email_entry.pack(side="left", padx=5, pady=5)

        # Beginning Cash Balances (Column 0, Row 2)
        beg_frame = ctk.CTkFrame(main_container, fg_color=self.secondary_color, corner_radius=10)
        beg_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        
        ctk.CTkLabel(beg_frame, text="Beginning Cash Balances", font=("Roboto", 14, "bold"), text_color=self.text_color).pack(anchor="w", padx=5, pady=3)
        beg_inner = ctk.CTkFrame(beg_frame, fg_color="transparent")
        beg_inner.pack(padx=5, pady=3, fill="x")
        
        beg_items = [
            ("Cash in Bank (beginning):", self.cash_bank_beg),
            ("Cash on Hand (beginning):", self.cash_hand_beg),
        ]
        for i, (label, var) in enumerate(beg_items):
            ctk.CTkLabel(beg_inner, text=label, font=("Roboto", 11), text_color=self.text_color, anchor="w").grid(row=i, column=0, sticky="w", padx=5, pady=2)
            entry = ctk.CTkEntry(beg_inner, textvariable=var, width=150, font=("Roboto", 11), fg_color="#3A4F5D", text_color=self.text_color)
            entry.grid(row=i, column=1, padx=5, pady=2, sticky="e")
            self.format_entry(var, entry)

        # Cash Inflows (Column 0, Row 3)
        inflow_frame = ctk.CTkFrame(main_container, fg_color=self.secondary_color, corner_radius=10)
        inflow_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=5)
        
        ctk.CTkLabel(inflow_frame, text="Cash Inflows", font=("Roboto", 14, "bold"), text_color=self.text_color).pack(anchor="w", padx=5, pady=3)
        inflow_inner = ctk.CTkFrame(inflow_frame, fg_color="transparent")
        inflow_inner.pack(padx=5, pady=3, fill="x")
        
        inflow_items = [
            ("Monthly dues collected:", self.monthly_dues),
            ("Certifications issued:", self.certifications),
            ("Membership fee:", self.membership_fee),
            ("Vehicle stickers:", self.vehicle_stickers),
            ("Rentals (covered courts):", self.rentals),
            ("Solicitations/Donations:", self.solicitations),
            ("Interest Income on bank deposits:", self.interest_income),
            ("Livelihood Management Fee:", self.livelihood_fee),
            ("Others:", self.inflows_others),
        ]
        for i, (label, var) in enumerate(inflow_items):
            ctk.CTkLabel(inflow_inner, text=label, font=("Roboto", 11), text_color=self.text_color, anchor="w").grid(row=i, column=0, sticky="w", padx=5, pady=2)
            entry = ctk.CTkEntry(inflow_inner, textvariable=var, width=150, font=("Roboto", 11), fg_color="#3A4F5D", text_color=self.text_color)
            entry.grid(row=i, column=1, padx=5, pady=2, sticky="e")
            self.format_entry(var, entry)
        
        ctk.CTkLabel(inflow_inner, text="Total Cash receipt:", font=("Roboto", 11), text_color=self.text_color, anchor="w").grid(row=len(inflow_items), column=0, sticky="w", padx=5, pady=2)
        ctk.CTkEntry(inflow_inner, textvariable=self.total_receipts, width=150, font=("Roboto", 11), fg_color="#3A4F5D", text_color=self.text_color, state="disabled").grid(row=len(inflow_items), column=1, padx=5, pady=2, sticky="e")

        # Cash Outflows (Columns 1 and 2, Row 2 and 3)
        # Split into two parts to fit in two columns
        outflow_frame = ctk.CTkFrame(main_container, fg_color=self.secondary_color, corner_radius=10)
        outflow_frame.grid(row=2, column=1, rowspan=2, columnspan=2, sticky="nsew", padx=10, pady=5)
        
        ctk.CTkLabel(outflow_frame, text="Cash Outflows", font=("Roboto", 14, "bold"), text_color=self.text_color).pack(anchor="w", padx=5, pady=3)
        outflow_inner = ctk.CTkFrame(outflow_frame, fg_color="transparent")
        outflow_inner.pack(padx=5, pady=3, fill="both", expand=True)
        outflow_inner.grid_columnconfigure((0, 2), weight=1)

        # First row for the total outflows
        ctk.CTkLabel(outflow_inner, text="Cash Out Flows/Disbursements:", font=("Roboto", 11), text_color=self.text_color, anchor="w").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ctk.CTkEntry(outflow_inner, textvariable=self.cash_outflows, width=150, font=("Roboto", 11), fg_color="#3A4F5D", text_color=self.text_color, state="disabled").grid(row=0, column=1, padx=5, pady=2, sticky="e")

        # Split outflow items into two columns
        outflow_items = [
            ("Snacks/Meals for visitors:", self.snacks_meals),
            ("Transportation expenses:", self.transportation),
            ("Office supplies expense:", self.office_supplies),
            ("Printing and photocopy:", self.printing),
            ("Labor:", self.labor),
            ("Billboard expense:", self.billboard),
            ("Clearing/cleaning charges:", self.cleaning),
            ("Miscellaneous expenses:", self.misc_expenses),
            ("Federation fee:", self.federation_fee),
            ("HOA-BOD Uniforms:", self.uniforms),
            ("BOD Mtg:", self.bod_mtg),
            ("General Assembly:", self.general_assembly),
            ("Cash Deposit to bank:", self.cash_deposit),
            ("Withholding tax on bank deposit:", self.withholding_tax),
            ("Refund for seri-culture:", self.refund_sericulture),
            ("Others:", self.outflows_others),
            ("", self.outflows_others_2),
        ]

        # Split into two columns (9 items in first column, 8 in second)
        mid_point = (len(outflow_items) + 1) // 2  # Roughly half
        for i, (label, var) in enumerate(outflow_items[:mid_point]):
            ctk.CTkLabel(outflow_inner, text=label, font=("Roboto", 11), text_color=self.text_color, anchor="w").grid(row=i+1, column=0, sticky="w", padx=5, pady=2)
            entry = ctk.CTkEntry(outflow_inner, textvariable=var, width=150, font=("Roboto", 11), fg_color="#3A4F5D", text_color=self.text_color)
            entry.grid(row=i+1, column=1, padx=5, pady=2, sticky="e")
            self.format_entry(var, entry)

        for i, (label, var) in enumerate(outflow_items[mid_point:]):
            ctk.CTkLabel(outflow_inner, text=label, font=("Roboto", 11), text_color=self.text_color, anchor="w").grid(row=i+1, column=2, sticky="w", padx=5, pady=2)
            entry = ctk.CTkEntry(outflow_inner, textvariable=var, width=150, font=("Roboto", 11), fg_color="#3A4F5D", text_color=self.text_color)
            entry.grid(row=i+1, column=3, padx=5, pady=2, sticky="e")
            self.format_entry(var, entry)

        # Ending Cash Balances (Row 4, spans all columns)
        end_frame = ctk.CTkFrame(main_container, fg_color=self.secondary_color, corner_radius=10)
        end_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=10, pady=5)
        
        ctk.CTkLabel(end_frame, text="Ending Cash Balances", font=("Roboto", 14, "bold"), text_color=self.text_color).pack(anchor="w", padx=5, pady=3)
        end_inner = ctk.CTkFrame(end_frame, fg_color="transparent")
        end_inner.pack(padx=5, pady=3, fill="x")
        
        ctk.CTkLabel(end_inner, text="Ending cash balance:", font=("Roboto", 11), text_color=self.text_color, anchor="w").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ctk.CTkEntry(end_inner, textvariable=self.ending_cash, width=150, font=("Roboto", 11), fg_color="#3A4F5D", text_color=self.text_color, state="disabled").grid(row=0, column=1, padx=5, pady=2, sticky="e")
        
        ctk.CTkLabel(end_inner, text="Cash in Bank:", font=("Roboto", 11), text_color=self.text_color, anchor="w").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        bank_end_entry = ctk.CTkEntry(end_inner, textvariable=self.ending_cash_bank, width=150, font=("Roboto", 11), fg_color="#3A4F5D", text_color=self.text_color)
        bank_end_entry.grid(row=1, column=1, padx=5, pady=2, sticky="e")
        self.format_entry(self.ending_cash_bank, bank_end_entry)
        
        ctk.CTkLabel(end_inner, text="Cash on Hand:", font=("Roboto", 11), text_color=self.text_color, anchor="w").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        hand_end_entry = ctk.CTkEntry(end_inner, textvariable=self.ending_cash_hand, width=150, font=("Roboto", 11), fg_color="#3A4F5D", text_color=self.text_color)
        hand_end_entry.grid(row=2, column=1, padx=5, pady=2, sticky="e")
        self.format_entry(self.ending_cash_hand, hand_end_entry)

        # Buttons Frame (Row 5, spans all columns)
        button_frame = ctk.CTkFrame(main_container, fg_color=self.primary_color)
        button_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=10)
        
        buttons = [
            ("Save to CSV (Ctrl+S)", self.save_to_csv),
            ("Load from CSV (Ctrl+L)", self.load_from_csv),
            ("Clear All Fields", self.clear_fields),
            ("Export to PDF (Ctrl+E)", self.export_to_pdf),
            ("Save to Word (Ctrl+W)", self.save_to_docx),
            ("Send via Email (Ctrl+G)", self.send_email),
        ]
        for text, command in buttons:
            ctk.CTkButton(
                button_frame,
                text=text,
                command=command,
                font=("Roboto", 12),
                fg_color=self.accent_color,
                hover_color="#008CC1",
                text_color=self.text_color,
                width=180,
                height=35,
                corner_radius=8
            ).pack(side="left", padx=5)

    def show_calendar(self):
        """Show a standalone calendar in a popup window, appearing directly in the center."""
        popup = tk.Toplevel(self.root)
        popup.title("Select Date")
        popup.geometry("300x300")
        popup.transient(self.root)
        popup.grab_set()
        popup.withdraw()

        popup.update_idletasks()
        popup_width = 300
        popup_height = 300
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()

        x = main_x + (main_width - popup_width) // 2
        y = main_y + (main_height - popup_height) // 2
        popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")

        cal = HoverCalendar(
            popup,
            font=("Arial", 12),
            background="#2A3F4D",
            foreground="#E0E0E0",
            selectbackground="#2A3F4D",
            selectforeground="#E0E0E0",
            normalbackground="#FFFFFF",
            normalforeground="#000000",
            weekendbackground="#FFFFFF",
            weekendforeground="#000000",
            headersbackground="#2A3F4D",
            headersforeground="#E0E0E0",
            showothermonthdays=False,
            showweeknumbers=False
        )
        cal.pack(padx=10, pady=10)

        try:
            current_date = datetime.datetime.strptime(self.today_date.get(), "%m/%d/%Y")
            cal.selection_set(current_date)
        except ValueError:
            pass

        def on_date_select():
            selected_date = cal.selection_get()
            if selected_date:
                self.today_date.set(selected_date.strftime("%m/%d/%Y"))
            popup.destroy()

        confirm_button = ctk.CTkButton(
            popup,
            text="Select",
            command=on_date_select,
            fg_color=self.accent_color,
            text_color=self.text_color,
            hover_color="#008CC1"
        )
        confirm_button.pack(pady=10)

        popup.deiconify()
        popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    def save_to_csv(self):
        try:
            filename = f"cash_flow_statement_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            with open(filename, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow([self.title_var.get()])
                writer.writerow([f"For the year month {self.display_date.get()}"])
                writer.writerow([])
                writer.writerow(["Cash in Bank-beg", self.cash_bank_beg.get()])
                writer.writerow(["Cash on Hand-beg", self.cash_hand_beg.get()])
                writer.writerow([])
                writer.writerow(["Cash inflows:"])
                writer.writerow(["Monthly dues collected", self.monthly_dues.get()])
                writer.writerow(["Certifications issued", self.certifications.get()])
                writer.writerow(["Membership fee", self.membership_fee.get()])
                writer.writerow(["Vehicle stickers", self.vehicle_stickers.get()])
                writer.writerow(["Rentals (covered courts)", self.rentals.get()])
                writer.writerow(["Solicitations/Donations", self.solicitations.get()])
                writer.writerow(["Interest Income on bank deposits", self.interest_income.get()])
                writer.writerow(["Livelihood Management Fee", self.livelihood_fee.get()])
                writer.writerow(["Others", self.inflows_others.get()])
                writer.writerow(["Total Cash receipt", self.total_receipts.get()])
                writer.writerow([])
                writer.writerow(["Less:"])
                writer.writerow(["Cash Out Flows/Disbursements", self.cash_outflows.get()])
                writer.writerow(["Snacks/Meals for visitors", self.snacks_meals.get()])
                writer.writerow(["Transportation expenses", self.transportation.get()])
                writer.writerow(["Office supplies expense", self.office_supplies.get()])
                writer.writerow(["Printing and photocopy", self.printing.get()])
                writer.writerow(["Labor", self.labor.get()])
                writer.writerow(["Billboard expense", self.billboard.get()])
                writer.writerow(["Clearing/cleaning charges", self.cleaning.get()])
                writer.writerow(["Miscellaneous expenses", self.misc_expenses.get()])
                writer.writerow(["Federation fee", self.federation_fee.get()])
                writer.writerow(["HOA-BOD Uniforms", self.uniforms.get()])
                writer.writerow(["BOD Mtg", self.bod_mtg.get()])
                writer.writerow(["General Assembly", self.general_assembly.get()])
                writer.writerow(["Cash Deposit to bank", self.cash_deposit.get()])
                writer.writerow(["Withholding tax on bank deposit", self.withholding_tax.get()])
                writer.writerow(["Refund for seri-culture", self.refund_sericulture.get()])
                writer.writerow(["Others", self.outflows_others.get()])
                writer.writerow(["", self.outflows_others_2.get()])
                writer.writerow([])
                writer.writerow(["Ending cash balance", self.ending_cash.get()])
                writer.writerow([])
                writer.writerow(["Breakdown of cash:"])
                writer.writerow(["Cash in Bank", self.ending_cash_bank.get()])
                writer.writerow(["Cash on Hand", self.ending_cash_hand.get()])
            
            messagebox.showinfo("Success", f"Cash flow statement saved to {filename}")
            return filename
        except Exception as e:
            messagebox.showerror("Error", f"Error saving to CSV: {str(e)}")
            return None

    def load_from_csv(self):
        try:
            filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
            if not filename:
                return
                
            with open(filename, 'r') as csvfile:
                reader = csv.reader(csvfile)
                data = list(reader)
                self.title_var.set(data[0][0])
                try:
                    date_str = data[1][0].replace("For the year month ", "")
                    date_obj = datetime.datetime.strptime(date_str, "%b %d, %Y")
                    self.today_date.set(date_obj.strftime("%m/%d/%Y"))
                except ValueError:
                    self.today_date.set(datetime.datetime.now().strftime("%m/%d/%Y"))
                self.cash_bank_beg.set(data[3][1])
                self.cash_hand_beg.set(data[4][1])
                self.monthly_dues.set(data[7][1])
                self.certifications.set(data[8][1])
                self.membership_fee.set(data[9][1])
                self.vehicle_stickers.set(data[10][1])
                self.rentals.set(data[11][1])
                self.solicitations.set(data[12][1])
                self.interest_income.set(data[13][1])
                self.livelihood_fee.set(data[14][1])
                self.inflows_others.set(data[15][1])
                self.total_receipts.set(data[16][1])
                self.cash_outflows.set(data[19][1])
                self.snacks_meals.set(data[20][1])
                self.transportation.set(data[21][1])
                self.office_supplies.set(data[22][1])
                self.printing.set(data[23][1])
                self.labor.set(data[24][1])
                self.billboard.set(data[25][1])
                self.cleaning.set(data[26][1])
                self.misc_expenses.set(data[27][1])
                self.federation_fee.set(data[28][1])
                self.uniforms.set(data[29][1])
                self.bod_mtg.set(data[30][1])
                self.general_assembly.set(data[31][1])
                self.cash_deposit.set(data[32][1])
                self.withholding_tax.set(data[33][1])
                self.refund_sericulture.set(data[34][1])
                self.outflows_others.set(data[35][1])
                self.outflows_others_2.set(data[36][1])
                self.ending_cash.set(data[38][1])
                self.ending_cash_bank.set(data[41][1])
                self.ending_cash_hand.set(data[42][1])
            
            messagebox.showinfo("Success", f"Loaded data from {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Error loading CSV: {str(e)}")

    def clear_fields(self):
        for var in [
            self.cash_bank_beg, self.cash_hand_beg,
            self.monthly_dues, self.certifications, self.membership_fee,
            self.vehicle_stickers, self.rentals, self.solicitations,
            self.interest_income, self.livelihood_fee, self.inflows_others,
            self.snacks_meals, self.transportation, self.office_supplies,
            self.printing, self.labor, self.billboard, self.cleaning,
            self.misc_expenses, self.federation_fee, self.uniforms,
            self.bod_mtg, self.general_assembly, self.cash_deposit,
            self.withholding_tax, self.refund_sericulture, self.outflows_others,
            self.outflows_others_2, self.ending_cash_bank, self.ending_cash_hand
        ]:
            var.set("")
        
        self.total_receipts.set("")
        self.cash_outflows.set("")
        self.ending_cash.set("")
        self.today_date.set(datetime.datetime.now().strftime("%m/%d/%Y"))
        messagebox.showinfo("Success", "All fields have been cleared")

    def export_to_pdf(self, temp=False):
        try:
            self.calculate_totals()
            def format_amount(value):
                if value:
                    try:
                        amount = Decimal(value.replace(',', ''))
                        return f"{amount:,.2f}"
                    except:
                        return value
                return ""
            
            if temp:
                filename = f"temp_cash_flow_statement_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            else:
                filename = self.get_save_filename("PDF", ".pdf")
                if not filename:
                    return None

            doc = SimpleDocTemplate(filename, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(__file__)
            logo_path = os.path.join(base_path, "logo.png")
            if os.path.exists(logo_path):
                elements.append(Image(logo_path, width=100, height=100))
                elements.append(Spacer(1, 12))
            elements.append(Paragraph(self.title_var.get(), styles['Title']))
            elements.append(Paragraph(f"For the year month {self.display_date.get()}", styles['Normal']))
            elements.append(Spacer(1, 12))
            beg_data = [
                ["Cash in Bank-beg", format_amount(self.cash_bank_beg.get())],
                ["Cash on Hand-beg", format_amount(self.cash_hand_beg.get())]
            ]
            beg_table = Table(beg_data, colWidths=[300, 150])
            beg_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ]))
            elements.append(beg_table)
            elements.append(Spacer(1, 12))
            elements.append(Paragraph("<b>Cash inflows:</b>", styles['Normal']))
            elements.append(Spacer(1, 6))
            inflows_data = [
                ["Monthly dues collected", format_amount(self.monthly_dues.get())],
                ["Certifications issued", format_amount(self.certifications.get())],
                ["Membership fee", format_amount(self.membership_fee.get())],
                ["Vehicle stickers", format_amount(self.vehicle_stickers.get())],
                ["Rentals (covered courts)", format_amount(self.rentals.get())],
                ["Solicitations/Donations", format_amount(self.solicitations.get())],
                ["Interest Income on bank deposits", format_amount(self.interest_income.get())],
                ["Livelihood Management Fee", format_amount(self.livelihood_fee.get())],
                ["Others", format_amount(self.inflows_others.get())],
                ["Total Cash receipt", format_amount(self.total_receipts.get())]
            ]
            inflows_table = Table(inflows_data, colWidths=[300, 150])
            inflows_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ]))
            elements.append(inflows_table)
            elements.append(Spacer(1, 12))
            elements.append(Paragraph("<b>Less:</b>", styles['Normal']))
            elements.append(Spacer(1, 6))
            outflows_data = [
                ["Cash Out Flows/Disbursements", format_amount(self.cash_outflows.get())],
                ["Snacks/Meals for visitors", format_amount(self.snacks_meals.get())],
                ["Transportation expenses", format_amount(self.transportation.get())],
                ["Office supplies expense", format_amount(self.office_supplies.get())],
                ["Printing and photocopy", format_amount(self.printing.get())],
                ["Labor", format_amount(self.labor.get())],
                ["Billboard expense", format_amount(self.billboard.get())],
                ["Clearing/cleaning charges", format_amount(self.cleaning.get())],
                ["Miscellaneous expenses", format_amount(self.misc_expenses.get())],
                ["Federation fee", format_amount(self.federation_fee.get())],
                ["HOA-BOD Uniforms", format_amount(self.uniforms.get())],
                ["BOD Mtg", format_amount(self.bod_mtg.get())],
                ["General Assembly", format_amount(self.general_assembly.get())],
                ["Cash Deposit to bank", format_amount(self.cash_deposit.get())],
                ["Withholding tax on bank deposit", format_amount(self.withholding_tax.get())],
                ["Refund for seri-culture", format_amount(self.refund_sericulture.get())],
                ["Others", format_amount(self.outflows_others.get())],
                ["", format_amount(self.outflows_others_2.get())]
            ]
            outflows_table = Table(outflows_data, colWidths=[300, 150])
            outflows_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (0, 0), colors.lightgrey),
                ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ]))
            elements.append(outflows_table)
            elements.append(Spacer(1, 12))
            ending_data = [
                ["Ending cash balance", format_amount(self.ending_cash.get())]
            ]
            ending_table = Table(ending_data, colWidths=[300, 150])
            ending_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (-1, -1), colors.lightgrey),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ]))
            elements.append(ending_table)
            elements.append(Spacer(1, 12))
            elements.append(Paragraph("<b>Breakdown of cash:</b>", styles['Normal']))
            elements.append(Spacer(1, 6))
            breakdown_data = [
                ["Cash in Bank", format_amount(self.ending_cash_bank.get())],
                ["Cash on Hand", format_amount(self.ending_cash_hand.get())]
            ]
            breakdown_table = Table(breakdown_data, colWidths=[300, 150])
            breakdown_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ]))
            elements.append(breakdown_table)
            def add_page_numbers(canvas, doc):
                page_num = canvas.getPageNumber()
                text = f"Page {page_num}"
                canvas.drawRightString(200, 20, text)
            doc.build(elements, onFirstPage=add_page_numbers, onLaterPages=add_page_numbers)
            if not temp:
                messagebox.showinfo("Success", f"PDF successfully exported to {filename}")
            return filename
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting to PDF: {str(e)}")
            return None

    def save_to_docx(self, temp=False):
        try:
            self.calculate_totals()
            def format_amount(value):
                if value:
                    try:
                        amount = Decimal(value.replace(',', ''))
                        return f"{amount:,.2f}"
                    except:
                        return value
                return ""
            
            if temp:
                filename = f"temp_cash_flow_statement_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            else:
                filename = self.get_save_filename("Word", ".docx")
                if not filename:
                    return None

            doc = Document()
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(__file__)
            logo_path = os.path.join(base_path, "logo.png")
            if os.path.exists(logo_path):
                paragraph = doc.add_paragraph()
                run = paragraph.add_run()
                run.add_picture(logo_path, width=Inches(1.5))
                paragraph.alignment = 1
            doc.add_heading(self.title_var.get(), level=1)
            doc.add_paragraph(f"For the year month {self.display_date.get()}")
            doc.add_heading("Beginning Cash Balances", level=2)
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            table.cell(0, 0).text = "Cash in Bank-beg"
            table.cell(0, 1).text = format_amount(self.cash_bank_beg.get())
            table.cell(1, 0).text = "Cash on Hand-beg"
            table.cell(1, 1).text = format_amount(self.cash_hand_beg.get())
            for row in table.rows:
                row.cells[1].paragraphs[0].alignment = 2
            doc.add_heading("Cash Inflows", level=2)
            table = doc.add_table(rows=10, cols=2)
            table.style = 'Table Grid'
            inflow_items = [
                ("Monthly dues collected", self.monthly_dues),
                ("Certifications issued", self.certifications),
                ("Membership fee", self.membership_fee),
                ("Vehicle stickers", self.vehicle_stickers),
                ("Rentals (covered courts)", self.rentals),
                ("Solicitations/Donations", self.solicitations),
                ("Interest Income on bank deposits", self.interest_income),
                ("Livelihood Management Fee", self.livelihood_fee),
                ("Others", self.inflows_others),
                ("Total Cash receipt", self.total_receipts)
            ]
            for i, (label, var) in enumerate(inflow_items):
                table.cell(i, 0).text = label
                table.cell(i, 1).text = format_amount(var.get())
                table.cell(i, 1).paragraphs[0].alignment = 2
            table.cell(9, 0).paragraphs[0].runs[0].bold = True
            doc.add_heading("Less: Cash Outflows", level=2)
            table = doc.add_table(rows=18, cols=2)
            table.style = 'Table Grid'
            outflow_items = [
                ("Cash Out Flows/Disbursements", self.cash_outflows),
                ("Snacks/Meals for visitors", self.snacks_meals),
                ("Transportation expenses", self.transportation),
                ("Office supplies expense", self.office_supplies),
                ("Printing and photocopy", self.printing),
                ("Labor", self.labor),
                ("Billboard expense", self.billboard),
                ("Clearing/cleaning charges", self.cleaning),
                ("Miscellaneous expenses", self.misc_expenses),
                ("Federation fee", self.federation_fee),
                ("HOA-BOD Uniforms", self.uniforms),
                ("BOD Mtg", self.bod_mtg),
                ("General Assembly", self.general_assembly),
                ("Cash Deposit to bank", self.cash_deposit),
                ("Withholding tax on bank deposit", self.withholding_tax),
                ("Refund for seri-culture", self.refund_sericulture),
                ("Others", self.outflows_others),
                ("", self.outflows_others_2)
            ]
            for i, (label, var) in enumerate(outflow_items):
                table.cell(i, 0).text = label
                table.cell(i, 1).text = format_amount(var.get())
                table.cell(i, 1).paragraphs[0].alignment = 2
            table.cell(0, 0).paragraphs[0].runs[0].bold = True
            doc.add_heading("Ending Cash Balance", level=2)
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            table.cell(0, 0).text = "Ending cash balance"
            table.cell(0, 1).text = format_amount(self.ending_cash.get())
            table.cell(0, 1).paragraphs[0].alignment = 2
            table.cell(0, 0).paragraphs[0].runs[0].bold = True
            doc.add_heading("Breakdown of Cash", level=2)
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            table.cell(0, 0).text = "Cash in Bank"
            table.cell(0, 1).text = format_amount(self.ending_cash_bank.get())
            table.cell(1, 0).text = "Cash on Hand"
            table.cell(1, 1).text = format_amount(self.ending_cash_hand.get())
            for row in table.rows:
                row.cells[1].paragraphs[0].alignment = 2
            doc.save(filename)
            if not temp:
                messagebox.showinfo("Success", f"Word document saved to {filename}")
            return filename
        except Exception as e:
            messagebox.showerror("Error", f"Error saving to Word: {str(e)}")
            return None

    def send_email(self):
        try:
            recipient_emails = [email.strip() for email in self.recipient_emails_var.get().split(',') if email.strip()]
            if not recipient_emails:
                messagebox.showerror("Error", "Please enter at least one recipient email.")
                return

            pdf_filename = self.export_to_pdf(temp=True)
            if not pdf_filename:
                return
            docx_filename = self.save_to_docx(temp=True)
            if not docx_filename:
                os.remove(pdf_filename)
                return

            service = self.get_gmail_service()

            msg = MIMEMultipart()
            msg['From'] = SENDER_EMAIL
            msg['To'] = ", ".join(recipient_emails)
            msg['Subject'] = f"Cash Flow Statement - {self.display_date.get()}"
            msg.attach(MIMEText(f"Attached is the cash flow statement for {self.display_date.get()}.\n\nRegards,\nCash Flow App", 'plain'))

            for filename in [pdf_filename, docx_filename]:
                with open(filename, 'rb') as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(filename))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(filename)}"'
                    msg.attach(part)

            raw = base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')
            service.users().messages().send(userId='me', body={'raw': raw}).execute()

            messagebox.showinfo("Success", f"Email sent to {', '.join(recipient_emails)}!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email: {str(e)}")
        finally:
            if 'pdf_filename' in locals() and os.path.exists(pdf_filename):
                os.remove(pdf_filename)
            if 'docx_filename' in locals() and os.path.exists(docx_filename):
                os.remove(docx_filename)

    def setup_keyboard_shortcuts(self):
        self.root.bind('<Control-s>', lambda e: self.save_to_csv())
        self.root.bind('<Control-e>', lambda e: self.export_to_pdf())
        self.root.bind('<Control-l>', lambda e: self.load_from_csv())
        self.root.bind('<Control-g>', lambda e: self.send_email())
        self.root.bind('<Control-w>', lambda e: self.save_to_docx())

    def format_entry(self, var, entry_widget):
        def on_change(*args):
            value = var.get()
            if value:
                try:
                    formatted = f"{Decimal(value.replace(',', '')):,.2f}"
                    var.set(formatted)
                except:
                    pass
        var.trace('w', on_change)
        entry_widget.configure(justify="right")

    def safe_decimal(self, var):
        val = var.get().strip()
        if not val:
            return Decimal("0")
        try:
            val = val.replace(",", "")
            return Decimal(val)
        except:
            raise ValueError(f"Invalid number: {val}")

    def calculate_totals(self):
        try:
            inflow_total = sum([
                self.safe_decimal(self.monthly_dues),
                self.safe_decimal(self.certifications),
                self.safe_decimal(self.membership_fee),
                self.safe_decimal(self.vehicle_stickers),
                self.safe_decimal(self.rentals),
                self.safe_decimal(self.solicitations),
                self.safe_decimal(self.interest_income),
                self.safe_decimal(self.livelihood_fee),
                self.safe_decimal(self.inflows_others)
            ])
            
            outflow_total = sum([
                self.safe_decimal(self.snacks_meals),
                self.safe_decimal(self.transportation),
                self.safe_decimal(self.office_supplies),
                self.safe_decimal(self.printing),
                self.safe_decimal(self.labor),
                self.safe_decimal(self.billboard),
                self.safe_decimal(self.cleaning),
                self.safe_decimal(self.misc_expenses),
                self.safe_decimal(self.federation_fee),
                self.safe_decimal(self.uniforms),
                self.safe_decimal(self.bod_mtg),
                self.safe_decimal(self.general_assembly),
                self.safe_decimal(self.cash_deposit),
                self.safe_decimal(self.withholding_tax),
                self.safe_decimal(self.refund_sericulture),
                self.safe_decimal(self.outflows_others),
                self.safe_decimal(self.outflows_others_2)
            ])
            
            beginning_total = self.safe_decimal(self.cash_bank_beg) + self.safe_decimal(self.cash_hand_beg)
            ending_balance = beginning_total + inflow_total - outflow_total
            
            self.total_receipts.set(f"{inflow_total:,.2f}")
            self.cash_outflows.set(f"{outflow_total:,.2f}")
            self.ending_cash.set(f"{ending_balance:,.2f}")
            
            if not self.ending_cash_bank.get() and not self.ending_cash_hand.get():
                self.ending_cash_bank.set(f"{ending_balance * Decimal('0.8'):,.2f}")
                self.ending_cash_hand.set(f"{ending_balance * Decimal('0.2'):,.2f}")
        except Exception as e:
            messagebox.showerror("Error", f"Calculation error: {str(e)}\nPlease ensure all values are valid numbers.")

    def get_save_filename(self, file_type, default_extension):
        default_filename = f"cash_flow_statement_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}{default_extension}"
        filetypes = [(f"{file_type} files", f"*{default_extension}"), ("All files", "*.*")]
        filename = filedialog.asksaveasfilename(
            defaultextension=default_extension,
            initialfile=default_filename,
            filetypes=filetypes,
            title=f"Save {file_type} As"
        )
        return filename

    def get_gmail_service(self):
        creds = None
        if os.path.exists(TOKEN_FILE):
            with open(TOKEN_FILE, 'rb') as token:
                creds = pickle.load(token)
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
                with open(TOKEN_FILE, 'wb') as token:
                    pickle.dump(creds, token)
        
        return build('gmail', 'v1', credentials=creds)

if __name__ == "__main__":
    root = ctk.CTk()
    app = IntegratedCashFlowApp(root)
    root.mainloop()