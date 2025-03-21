import tkinter as tk
from tkinter import ttk, messagebox
import datetime
import csv
import os
from decimal import Decimal
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

class CashFlowStatementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cash Flow Statement Generator")
        self.root.geometry("800x800")
        
        # Variables to store entry values
        self.title_var = tk.StringVar(value="Statement Of Cash Flows")
        self.today_date = datetime.datetime.now().strftime("%B %d, %Y")
        
        self.cash_bank_beg = tk.StringVar()
        self.cash_hand_beg = tk.StringVar()
        
        # Cash inflows variables
        self.monthly_dues = tk.StringVar()
        self.certifications = tk.StringVar()
        self.membership_fee = tk.StringVar()
        self.vehicle_stickers = tk.StringVar()
        self.rentals = tk.StringVar()
        self.solicitations = tk.StringVar()
        self.interest_income = tk.StringVar()
        self.livelihood_fee = tk.StringVar()
        self.inflows_others = tk.StringVar()
        self.total_receipts = tk.StringVar()
        
        # Cash outflows variables
        self.cash_outflows = tk.StringVar()
        self.snacks_meals = tk.StringVar()
        self.transportation = tk.StringVar()
        self.office_supplies = tk.StringVar()
        self.printing = tk.StringVar()
        self.labor = tk.StringVar()
        self.billboard = tk.StringVar()
        self.cleaning = tk.StringVar()
        self.misc_expenses = tk.StringVar()
        self.federation_fee = tk.StringVar()
        self.uniforms = tk.StringVar()
        self.bod_mtg = tk.StringVar()
        self.general_assembly = tk.StringVar()
        self.cash_deposit = tk.StringVar()
        self.withholding_tax = tk.StringVar()
        self.refund_sericulture = tk.StringVar()
        self.outflows_others = tk.StringVar()
        self.outflows_others_2 = tk.StringVar()
        
        # Ending balances
        self.ending_cash = tk.StringVar()
        self.ending_cash_bank = tk.StringVar()
        self.ending_cash_hand = tk.StringVar()
        
        self.create_widgets()
    
    def create_widgets(self):
        # Create a main frame with scrolling
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create a canvas with scrollbar
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
        
        title_label = ttk.Label(title_frame, text="Title:")
        title_label.pack(side="left", padx=5)
        title_entry = ttk.Entry(title_frame, textvariable=self.title_var, width=40)
        title_entry.pack(side="left", padx=5)
        
        date_label = ttk.Label(title_frame, text=f"Date: {self.today_date}")
        date_label.pack(side="left", padx=20)
        
        # Beginning Cash Balances
        beg_frame = ttk.LabelFrame(scrollable_frame, text="Beginning Cash Balances")
        beg_frame.pack(fill="x", pady=5)
        
        ttk.Label(beg_frame, text="Cash in Bank (beginning):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(beg_frame, textvariable=self.cash_bank_beg, width=15).grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(beg_frame, text="Cash on Hand (beginning):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(beg_frame, textvariable=self.cash_hand_beg, width=15).grid(row=1, column=1, padx=5, pady=2)
        
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
            ttk.Entry(inflow_frame, textvariable=var, width=15).grid(row=i, column=1, padx=5, pady=2)
        
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
            ttk.Entry(outflow_frame, textvariable=var, width=15).grid(row=i+1, column=1, padx=5, pady=2)
        
        # Ending Cash Balances
        end_frame = ttk.LabelFrame(scrollable_frame, text="Ending Cash Balances")
        end_frame.pack(fill="x", pady=5)
        
        ttk.Label(end_frame, text="Ending cash balance:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(end_frame, textvariable=self.ending_cash, width=15, state="readonly").grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(end_frame, text="Cash in Bank:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(end_frame, textvariable=self.ending_cash_bank, width=15).grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Label(end_frame, text="Cash on Hand:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(end_frame, textvariable=self.ending_cash_hand, width=15).grid(row=2, column=1, padx=5, pady=2)
        
        # Buttons Frame
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(fill="x", pady=10)
        
        calc_button = ttk.Button(button_frame, text="Calculate Totals", command=self.calculate_totals)
        calc_button.pack(side="left", padx=5)
        
        save_button = ttk.Button(button_frame, text="Save to CSV", command=self.save_to_csv)
        save_button.pack(side="left", padx=5)
        
        clear_button = ttk.Button(button_frame, text="Clear All Fields", command=self.clear_fields)
        clear_button.pack(side="left", padx=5)
        
        print_button = ttk.Button(button_frame, text="Print Statement", command=self.print_statement)
        print_button.pack(side="left", padx=5)
        
        # New PDF export button
        pdf_button = ttk.Button(button_frame, text="Export to PDF", command=self.export_to_pdf)
        pdf_button.pack(side="left", padx=5)
    
    def calculate_totals(self):
        try:
            # Convert entries to decimal, treating empty as 0
            def safe_decimal(var):
                val = var.get()
                if val == "":
                    return Decimal("0")
                return Decimal(val)
            
            # Calculate total cash receipts
            inflow_total = sum([
                safe_decimal(self.monthly_dues),
                safe_decimal(self.certifications),
                safe_decimal(self.membership_fee),
                safe_decimal(self.vehicle_stickers),
                safe_decimal(self.rentals),
                safe_decimal(self.solicitations),
                safe_decimal(self.interest_income),
                safe_decimal(self.livelihood_fee),
                safe_decimal(self.inflows_others)
            ])
            
            # Calculate total cash outflows
            outflow_total = sum([
                safe_decimal(self.snacks_meals),
                safe_decimal(self.transportation),
                safe_decimal(self.office_supplies),
                safe_decimal(self.printing),
                safe_decimal(self.labor),
                safe_decimal(self.billboard),
                safe_decimal(self.cleaning),
                safe_decimal(self.misc_expenses),
                safe_decimal(self.federation_fee),
                safe_decimal(self.uniforms),
                safe_decimal(self.bod_mtg),
                safe_decimal(self.general_assembly),
                safe_decimal(self.cash_deposit),
                safe_decimal(self.withholding_tax),
                safe_decimal(self.refund_sericulture),
                safe_decimal(self.outflows_others),
                safe_decimal(self.outflows_others_2)
            ])
            
            # Calculate ending cash balance
            beginning_total = safe_decimal(self.cash_bank_beg) + safe_decimal(self.cash_hand_beg)
            ending_balance = beginning_total + inflow_total - outflow_total
            
            # Update the calculated fields
            self.total_receipts.set(str(inflow_total))
            self.cash_outflows.set(str(outflow_total))
            self.ending_cash.set(str(ending_balance))
            
            # Auto-fill the ending cash breakdown if not manually entered
            if not self.ending_cash_bank.get() and not self.ending_cash_hand.get():
                # Just for a default split, put most in bank
                self.ending_cash_bank.set(str(ending_balance * Decimal("0.8")))
                self.ending_cash_hand.set(str(ending_balance * Decimal("0.2")))
            
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
    
    def clear_fields(self):
        # Clear all entry fields
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
        
        # Clear calculated fields
        self.total_receipts.set("")
        self.cash_outflows.set("")
        self.ending_cash.set("")
        
        messagebox.showinfo("Success", "All fields have been cleared")
    
    def print_statement(self):
        try:
            # Create a simple text representation that could be printed
            filename = f"cash_flow_statement_print_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            
            with open(filename, 'w') as f:
                f.write(f"{self.title_var.get()}\n")
                f.write(f"For the year month {self.today_date}\n\n")
                
                f.write(f"Cash in Bank-beg ({self.today_date}): {self.cash_bank_beg.get()}\n")
                f.write(f"Cash on Hand-beg: {self.cash_hand_beg.get()}\n\n")
                
                f.write("Cash inflows:\n")
                f.write(f"Monthly dues collected: {self.monthly_dues.get()}\n")
                f.write(f"Certifications issued: {self.certifications.get()}\n")
                f.write(f"Membership fee: {self.membership_fee.get()}\n")
                f.write(f"Vehicle stickers: {self.vehicle_stickers.get()}\n")
                f.write(f"Rentals (covered courts): {self.rentals.get()}\n")
                f.write(f"Solicitations/Donations: {self.solicitations.get()}\n")
                f.write(f"Interest Income on bank deposits: {self.interest_income.get()}\n")
                f.write(f"Livelihood Management Fee: {self.livelihood_fee.get()}\n")
                f.write(f"Others: {self.inflows_others.get()}\n")
                f.write(f"Total Cash receipt: {self.total_receipts.get()}\n\n")
                
                f.write("Less:\n")
                f.write(f"Cash Out Flows/Disbursements: {self.cash_outflows.get()}\n")
                f.write(f"Snacks/Meals for visitors: {self.snacks_meals.get()}\n")
                f.write(f"Transportation expenses: {self.transportation.get()}\n")
                f.write(f"Office supplies expense: {self.office_supplies.get()}\n")
                f.write(f"Printing and photocopy: {self.printing.get()}\n")
                f.write(f"Labor: {self.labor.get()}\n")
                f.write(f"Billboard expense: {self.billboard.get()}\n")
                f.write(f"Clearing/cleaning charges: {self.cleaning.get()}\n")
                f.write(f"Miscellaneous expenses: {self.misc_expenses.get()}\n")
                f.write(f"Federation fee: {self.federation_fee.get()}\n")
                f.write(f"HOA-BOD Uniforms: {self.uniforms.get()}\n")
                f.write(f"BOD Mtg: {self.bod_mtg.get()}\n")
                f.write(f"General Assembly: {self.general_assembly.get()}\n")
                f.write(f"Cash Deposit to bank: {self.cash_deposit.get()}\n")
                f.write(f"Withholding tax on bank deposit: {self.withholding_tax.get()}\n")
                f.write(f"Refund for seri-culture: {self.refund_sericulture.get()}\n")
                f.write(f"Others: {self.outflows_others.get()} {self.outflows_others_2.get()}\n\n")
                
                f.write(f"Ending cash balance: {self.ending_cash.get()}\n\n")
                
                f.write("Breakdown of cash:\n")
                f.write(f"Cash in Bank: {self.ending_cash_bank.get()}\n")
                f.write(f"Cash on Hand: {self.ending_cash_hand.get()}\n")
            
            # Confirm file was created and give instructions
            if os.path.exists(filename):
                messagebox.showinfo("Success", f"Print file created: {filename}\nYou can open this file and print it from your text editor.")
            else:
                raise Exception("File was not created successfully.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error creating print file: {str(e)}")
    
    def export_to_pdf(self):
        try:
            # Format amounts for display
            def format_amount(value):
                if value:
                    try:
                        # Try to convert to Decimal for proper formatting
                        amount = Decimal(value)
                        return f"{amount:,.2f}"
                    except:
                        return value
                return ""

            # Create filename with timestamp
            filename = f"cash_flow_statement_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            
            # Create the PDF document
            doc = SimpleDocTemplate(filename, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []
            
            # Title
            title_style = styles['Title']
            elements.append(Paragraph(self.title_var.get(), title_style))
            elements.append(Paragraph(f"For the year month {self.today_date}", styles['Normal']))
            elements.append(Spacer(1, 12))
            
            # Beginning balances
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
            
            # Cash inflows
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
            
            # Cash outflows
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
            
            # Ending balances
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
            
            # Breakdown of cash
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
            
            # Build the PDF
            doc.build(elements)
            
            messagebox.showinfo("Success", f"PDF successfully exported to {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting to PDF: {str(e)}\n\nMake sure you have ReportLab installed by running:\npip install reportlab")


# Main application
if __name__ == "__main__":
    root = tk.Tk()
    app = CashFlowStatementApp(root)
    root.mainloop()