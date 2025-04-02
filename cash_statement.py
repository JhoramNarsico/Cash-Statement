import tkinter as tk
from tkinter import ttk, messagebox
import datetime
import csv
import os
import pypandoc
import subprocess
from Google import Create_Service
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from docxtpl import DocxTemplate
from decimal import Decimal
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet

class CashFlowStatementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cash Flow Statement Generator")
        self.root.geometry("800x800")
        
        # Variables to store entry values
        self.title_var = tk.StringVar(value="Statement Of Cash Flows")
        self.today_date = datetime.datetime.now().strftime("%B %d, %Y")
        
        self.cash_bank_beg = tk.DoubleVar()
        self.cash_hand_beg = tk.DoubleVar()
        
        # Cash inflows variables
        self.monthly_dues = tk.DoubleVar()
        self.certifications = tk.DoubleVar()
        self.membership_fee = tk.DoubleVar()
        self.vehicle_stickers = tk.DoubleVar()
        self.rentals = tk.DoubleVar()
        self.solicitations = tk.DoubleVar()
        self.interest_income = tk.DoubleVar()
        self.livelihood_fee = tk.DoubleVar()
        self.inflows_others = tk.DoubleVar()
        self.total_receipts = tk.DoubleVar()
        
        # Cash outflows variables
        self.cash_outflows = tk.DoubleVar()
        self.snacks_meals = tk.DoubleVar()
        self.transportation = tk.DoubleVar()
        self.office_supplies = tk.DoubleVar()
        self.printing = tk.DoubleVar()
        self.labor = tk.DoubleVar()
        self.billboard = tk.DoubleVar()
        self.cleaning = tk.DoubleVar()
        self.misc_expenses = tk.DoubleVar()
        self.federation_fee = tk.DoubleVar()
        self.uniforms = tk.DoubleVar()
        self.bod_mtg = tk.DoubleVar()
        self.general_assembly = tk.DoubleVar()
        self.cash_deposit = tk.DoubleVar()
        self.withholding_tax = tk.DoubleVar()
        self.refund_sericulture = tk.DoubleVar()
        self.outflows_others = tk.DoubleVar()
        self.outflows_others_2 = tk.DoubleVar()
        
        # Ending balances
        self.ending_cash = tk.DoubleVar()
        self.ending_cash_bank = tk.DoubleVar()
        self.ending_cash_hand = tk.DoubleVar()
        
        self.create_widgets()
        self.setup_keyboard_shortcuts()

    def setup_keyboard_shortcuts(self):
        self.root.bind('<Control-s>', lambda e: self.save_to_csv())
        self.root.bind('<Control-e>', lambda e: self.export_to_pdf())
        self.root.bind('<Control-c>', lambda e: self.calculate_totals())
        self.root.bind('<Control-l>', lambda e: self.load_from_csv())

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
        entry_widget.config(justify='right')

    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Title and Date
        title_frame = ttk.Frame(scrollable_frame)
        title_frame.pack(fill="x", pady=5)
        
        ttk.Label(title_frame, text="Title:").pack(side="left", padx=5)
        ttk.Entry(title_frame, textvariable=self.title_var, width=40).pack(side="left", padx=5)
        ttk.Label(title_frame, text=f"Date: {self.today_date}").pack(side="left", padx=20)
        
        # Beginning Cash Balances
        beg_frame = ttk.LabelFrame(scrollable_frame, text="Beginning Cash Balances")
        beg_frame.pack(fill="x", pady=5)
        
        ttk.Label(beg_frame, text="Cash in Bank (beginning):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        bank_beg_entry = ttk.Entry(beg_frame, textvariable=self.cash_bank_beg, width=15)
        bank_beg_entry.grid(row=0, column=1, padx=5, pady=2)
        self.format_entry(self.cash_bank_beg, bank_beg_entry)
        
        ttk.Label(beg_frame, text="Cash on Hand (beginning):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        hand_beg_entry = ttk.Entry(beg_frame, textvariable=self.cash_hand_beg, width=15)
        hand_beg_entry.grid(row=1, column=1, padx=5, pady=2)
        self.format_entry(self.cash_hand_beg, hand_beg_entry)
        
        # Cash Inflows
        inflow_frame = ttk.LabelFrame(scrollable_frame, text="Cash Inflows")
        inflow_frame.pack(fill="x", pady=5)
        
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
            ttk.Label(inflow_frame, text=label).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            entry = ttk.Entry(inflow_frame, textvariable=var, width=15)
            entry.grid(row=i, column=1, padx=5, pady=2)
            self.format_entry(var, entry)
        
        ttk.Label(inflow_frame, text="Total Cash receipt:").grid(row=len(inflow_items), column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(inflow_frame, textvariable=self.total_receipts, width=15, state="readonly").grid(row=len(inflow_items), column=1, padx=5, pady=2)
        
        # Cash Outflows
        outflow_frame = ttk.LabelFrame(scrollable_frame, text="Cash Outflows")
        outflow_frame.pack(fill="x", pady=5)
        
        ttk.Label(outflow_frame, text="Cash Out Flows/Disbursements:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(outflow_frame, textvariable=self.cash_outflows, width=15, state="readonly").grid(row=0, column=1, padx=5, pady=2)
        
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
            ("", self.outflows_others_2)
        ]
        
        for i, (label, var) in enumerate(outflow_items):
            ttk.Label(outflow_frame, text=label).grid(row=i+1, column=0, sticky="w", padx=5, pady=2)
            entry = ttk.Entry(outflow_frame, textvariable=var, width=15)
            entry.grid(row=i+1, column=1, padx=5, pady=2)
            self.format_entry(var, entry)
        
        # Ending Cash Balances
        end_frame = ttk.LabelFrame(scrollable_frame, text="Ending Cash Balances")
        end_frame.pack(fill="x", pady=5)
        
        ttk.Label(end_frame, text="Ending cash balance:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(end_frame, textvariable=self.ending_cash, width=15, state="readonly").grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(end_frame, text="Cash in Bank:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        bank_end_entry = ttk.Entry(end_frame, textvariable=self.ending_cash_bank, width=15)
        bank_end_entry.grid(row=1, column=1, padx=5, pady=2)
        self.format_entry(self.ending_cash_bank, bank_end_entry)
        
        ttk.Label(end_frame, text="Cash on Hand:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        hand_end_entry = ttk.Entry(end_frame, textvariable=self.ending_cash_hand, width=15)
        hand_end_entry.grid(row=2, column=1, padx=5, pady=2)
        self.format_entry(self.ending_cash_hand, hand_end_entry)
        
        # Buttons Frame
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(fill="x", pady=10)
        
        ttk.Button(button_frame, text="Calculate Totals (Ctrl+C)", command=self.calculate_totals).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Save to CSV (Ctrl+S)", command=self.save_to_csv).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Load from CSV (Ctrl+L)", command=self.load_from_csv).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Clear All Fields", command=self.clear_fields).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Export to PDF (Ctrl+E)", command=self.export_to_pdf).pack(side="left", padx=5)

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
            
            messagebox.showinfo("Success", "Calculations complete!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Calculation error: {str(e)}\nPlease ensure all values are valid numbers.")

    def save_to_csv(self):
        try:
            filename = f"cash_flow_statement_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            with open(filename, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow([self.title_var.get()])
                writer.writerow([f"For the year month {self.today_date}"])
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
            
        except Exception as e:
            messagebox.showerror("Error", f"Error saving to CSV: {str(e)}")

    def load_from_csv(self):
        try:
            from tkinter import filedialog
            filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
            if not filename:
                return
                
            with open(filename, 'r') as csvfile:
                reader = csv.reader(csvfile)
                data = list(reader)
                self.title_var.set(data[0][0])
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
        
        messagebox.showinfo("Success", "All fields have been cleared")

<<<<<<< HEAD
=======
    def convertToPdf(self, filename):
        file_path = (Path(__file__).parent / filename).resolve()
        subprocess.run(["C:\Program Files\LibreOffice\program\soffice.exe", "--headless", "--convert-to", "pdf", str(file_path)]) # libre office location should be speicific
        messagebox.showinfo("Success", "PDF successfully exported and generated")
    
    def print_statement(self):
        try:
            document_path = Path(__file__).parent.parent / "sampledocumentslp.docx"
            doc = DocxTemplate(document_path)

            amount1 = self.cash_bank_beg.get()
            amount2 = self.cash_hand_beg.get()
            amount3 = self.monthly_dues.get()
            amount4 = self.certifications.get()
            amount5 = self.membership_fee.get()
            amount6 = self.vehicle_stickers.get()
            amount7 = self.rentals.get()
            amount8 = self.solicitations.get()
            amount9 = self.interest_income.get()
            amount10 = self.livelihood_fee.get()
            amount11 = self.inflows_others.get()
            amount12 = amount3+ amount4 + amount5 + amount6 + amount7 + amount8 + amount9 + amount10 + amount11
            amount13 = self.cash_outflows.get()
            amount14 = self.snacks_meals.get()
            amount15 = self.transportation.get()
            amount16 = self.office_supplies.get()
            amount17 = self.printing.get()
            amount18 = self.labor.get()
            amount19 = self.billboard.get()
            amount20 = self.cleaning.get()
            amount21 = self.misc_expenses.get()
            amount22 = self.federation_fee.get()
            amount23 = self.uniforms.get()
            amount24 = self.bod_mtg.get()
            amount25 = self.general_assembly.get()
            amount26 = self.cash_deposit.get()
            amount27 = self.withholding_tax.get()
            amount28 = self.refund_sericulture.get()
            amount29 = self.outflows_others_2.get()
            amount30 = amount13 + amount14 + amount15 + amount16 + amount17 + amount18 + amount19 + amount20 + amount21 + amount22 + amount23 + amount24 + amount25 + amount26 + amount27 + amount28 + amount29
            amount31 = amount12 - amount30
            amount32 = amount1
            amount33 = amount31 - amount1

            context = {
                "amount1": amount1, "amount2": amount2, "amount3": amount3, "amount4": amount4, "amount5": amount5,
                "amount6": amount6, "amount7": amount7, "amount8": amount8, "amount9": amount9, "amount10": amount10,
                "amount11": amount11, "amount12": amount12, "amount13": amount13, "amount14": amount14, "amount15": amount15,
                "amount16": amount16, "amount17": amount17, "amount18": amount18, "amount19": amount19, "amount20": amount20,
                "amount21": amount21, "amount22": amount22, "amount23": amount23, "amount24": amount24, "amount25": amount25,
                "amount26": amount26, "amount27": amount27, "amount28": amount28, "amount29": amount29, "amount30": amount30,
                "amount31": amount31, "amount32": amount32, "amount33": amount33
}
            doc.render(context)
            doc.save(Path(__file__).parent / "generated_doc.docx")
            self.convertToPdf("generated_doc.docx")
        except Exception as e:
            messagebox.showerror("Error", f"Error creating print file: {str(e)}")
 
>>>>>>> 5995b6b (initial commit)
    def export_to_pdf(self):
        CLIENT_SECRET__FILE = 'client_secret.json'
        API_NAME = 'gmail'
        api_version = 'v1'
        SCOPES = ['https://mail.google.com/']

        service = Create_Service(CLIENT_SECRET__FILE, API_NAME, api_version, SCOPES)

        try:
    # Path to your PDF file
            pdf_filename = "generated_doc.pdf"  # Change this to your actual file

    # Create email message
            emailMsg = "Testing with PDF attachment."
            mimeMessage = MIMEMultipart()
            mimeMessage["to"] = "ram05cembrano@gmail.com"
            mimeMessage["subject"] = "Test with PDF Attachment"
            mimeMessage.attach(MIMEText(emailMsg, "plain"))

    # Open the PDF file in binary mode and attach it
            with open(pdf_filename, "rb") as pdf_file:
                pdf_attachment = MIMEBase("application", "octet-stream")
                pdf_attachment.set_payload(pdf_file.read())

    # Encode file in base64
                encoders.encode_base64(pdf_attachment)

    # Set headers for attachment
                pdf_attachment.add_header("Content-Disposition", f"attachment; filename={os.path.basename(pdf_filename)}")

    # Attach PDF to email
                mimeMessage.attach(pdf_attachment)

    # Encode the entire email message as base64
                raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()

    # Send the email using Gmail API
                message = service.users().messages().send(userId="me", body={"raw": raw_string}).execute()

            print("Email sent successfully:", message)
            messagebox.showinfo("Success", f"PDF successfully exported and emailed!")

<<<<<<< HEAD
            filename = f"cash_flow_statement_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            doc = SimpleDocTemplate(filename, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []
            
            if os.path.exists("logo.png"):
                elements.append(Image("logo.png", width=100, height=100))
                elements.append(Spacer(1, 12))
            
            elements.append(Paragraph(self.title_var.get(), styles['Title']))
            elements.append(Paragraph(f"For the year month {self.today_date}", styles['Normal']))
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
            messagebox.showinfo("Success", f"PDF successfully exported to {filename}")
            
=======
>>>>>>> 5995b6b (initial commit)
        except Exception as e:
            messagebox.showerror("Error",f"Error exporting to PDF: {str(e)}\n\nMake sure you have ReportLab installed by running:\npip install reportlab",)

if __name__ == "__main__":
    root = tk.Tk()
    app = CashFlowStatementApp(root)
    root.mainloop()