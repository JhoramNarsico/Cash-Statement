import os
import re
from decimal import Decimal
from tkinter import filedialog, messagebox
import pdfplumber
from docx import Document
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from docx.shared import Inches
import sys
import datetime

class FileHandler:
    def __init__(self, variables, title_var, date_var):
        self.variables = variables
        self.title_var = title_var
        self.date_var = date_var

    def parse_amount(self, text):
        if not text or text.strip() == "":
            return ""
        try:
            return str(Decimal(re.sub(r'[^\d.]', '', text)))
        except:
            return text

    def format_date_for_display(self, date_str):
        try:
            date_obj = datetime.datetime.strptime(date_str, "%m/%d/%Y")
            return date_obj.strftime("%B %d, %Y")
        except ValueError:
            return date_str

    def load_from_docx(self, filename):
        try:
            print("Loading DOCX")
            with open(filename, 'rb') as file:
                doc = Document(filename)

            label_to_var = {
                "Cash in Bank-beg": self.variables['cash_bank_beg'],
                "Cash on Hand-beg": self.variables['cash_hand_beg'],
                "Monthly dues collected": self.variables['monthly_dues'],
                "Certifications issued": self.variables['certifications'],
                "Membership fee": self.variables['membership_fee'],
                "Vehicle stickers": self.variables['vehicle_stickers'],
                "Rentals": self.variables['rentals'],
                "Solicitations/Donations": self.variables['solicitations'],
                "Interest Income on bank deposits": self.variables['interest_income'],
                "Livelihood Management Fee": self.variables['livelihood_fee'],
                "Others(inflow)": self.variables['inflows_others'],
                "Cash Out Flows/Disbursements": self.variables['cash_outflows'],
                "Snacks/Meals for visitors": self.variables['snacks_meals'],
                "Transportation expenses": self.variables['transportation'],
                "Office supplies expense": self.variables['office_supplies'],
                "Printing and photocopy": self.variables['printing'],
                "Labor": self.variables['labor'],
                "Billboard expense": self.variables['billboard'],
                "Clearing/cleaning charges": self.variables['cleaning'],
                "Miscellaneous expenses": self.variables['misc_expenses'],
                "Federation fee": self.variables['federation_fee'],
                "HOA-BOD Uniforms": self.variables['uniforms'],
                "BOD Mtg": self.variables['bod_mtg'],
                "General Assembly": self.variables['general_assembly'],
                "Cash Deposit to bank": self.variables['cash_deposit'],
                "Withholding tax on bank deposit": self.variables['withholding_tax'],
                "Refund": self.variables['refund'],
                "Others(outflow)": self.variables['outflows_others']
            }

            for table in doc.tables:
                for row in table.rows:
                    if len(row.cells) >= 2:
                        label = row.cells[0].text.strip()
                        value = row.cells[1].text.strip()
                        print(f"Extracted Label: '{label}', Value: '{value}'")
                        if label.lower().startswith("for the year month"):
                            try:
                                date_str = label.split("month")[1].strip()
                                date_obj = datetime.datetime.strptime(date_str, "%B %d, %Y")
                                self.date_var.set(date_obj.strftime("%m/%d/%Y"))
                            except:
                                pass
                        if label in label_to_var:
                            label_to_var[label].set(self.parse_amount(value))

            messagebox.showinfo("Success", "Document loaded successfully")
            return True

        except Exception as e:
            messagebox.showerror("Error", f"Error loading Word document: {str(e)}")
            return False

    def load_from_pdf(self, filename):
        try:
            print(f"Attempting to load PDF: {filename}")

            label_to_var = {
                "Cash in Bank-beg": self.variables['cash_bank_beg'],
                "Cash on Hand-beg": self.variables['cash_hand_beg'],
                "Monthly dues collected": self.variables['monthly_dues'],
                "Certifications issued": self.variables['certifications'],
                "Membership fee": self.variables['membership_fee'],
                "Vehicle stickers": self.variables['vehicle_stickers'],
                "Rentals": self.variables['rentals'],
                "Solicitations/Donations": self.variables['solicitations'],
                "Interest Income on bank deposits": self.variables['interest_income'],
                "Livelihood Management Fee": self.variables['livelihood_fee'],
                "Others(inflow)": self.variables['inflows_others'],
                "Total Cash receipt": self.variables['total_receipts'],
                "Cash Out Flows/Disbursements": self.variables['cash_outflows'],
                "Snacks/Meals for visitors": self.variables['snacks_meals'],
                "Transportation expenses": self.variables['transportation'],
                "Office supplies expense": self.variables['office_supplies'],
                "Printing and photocopy": self.variables['printing'],
                "Labor": self.variables['labor'],
                "Billboard expense": self.variables['billboard'],
                "Clearing/cleaning charges": self.variables['cleaning'],
                "Miscellaneous expenses": self.variables['misc_expenses'],
                "Federation fee": self.variables['federation_fee'],
                "HOA-BOD Uniforms": self.variables['uniforms'],
                "BOD Mtg": self.variables['bod_mtg'],
                "General Assembly": self.variables['general_assembly'],
                "Cash Deposit to bank": self.variables['cash_deposit'],
                "Withholding tax on bank deposit": self.variables['withholding_tax'],
                "Refund": self.variables['refund'],
                "Others(outflow)": self.variables['outflows_others']
            }

            with pdfplumber.open(filename) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table or not table[0] or not table[0][0]:
                            print("Skipping empty or invalid table")
                            continue
                        for row in table:
                            if not row or len(row) < 2 or not row[0]:
                                print("Skipping empty or invalid row")
                                continue
                            label = row[0].strip()
                            value = row[1].strip() if row[1] else ""
                            print(f"Extracted: Label='{label}', Value='{value}'")
                            if label in label_to_var:
                                try:
                                    parsed_value = self.parse_amount(value)
                                    print(f"Setting {label} to {parsed_value}")
                                    label_to_var[label].set(parsed_value)
                                except Exception as e:
                                    print(f"Error setting {label}: {e}")

            with pdfplumber.open(filename) as pdf:
                first_page = pdf.pages[0]
                text = first_page.extract_text()
                for line in text.split('\n'):
                    line = line.strip()
                    if line.startswith("For the year month"):
                        try:
                            date_str = line.replace("For the year month", "").strip()
                            date_obj = datetime.datetime.strptime(date_str, "%B %d, %Y")
                            self.date_var.set(date_obj.strftime("%m/%d/%Y"))
                        except Exception as e:
                            print(f"Error parsing date: {e}")
                    elif line == self.title_var.get():
                        continue
                    elif line and not line.startswith(("Cash", "Less:", "Breakdown")):
                        self.title_var.set(line)

            messagebox.showinfo("Success", "PDF data loaded successfully!")
            return True

        except Exception as e:
            messagebox.showerror("Error", f"Error loading PDF: {str(e)}")
            return False

    def export_to_pdf(self):
        try:
            def format_amount(value):
                if value:
                    try:
                        amount = Decimal(value.replace(',', ''))
                        return f"{amount:,.2f}"
                    except:
                        return value
                return ""
            default_filename = f"cash_flow_statement_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            filename = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile=default_filename,
                title="Save PDF As"
            )
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
            elements.append(Paragraph(f"For the year month {self.format_date_for_display(self.date_var.get())}", styles['Normal']))
            elements.append(Spacer(1, 12))
            beg_data = [
                ["Cash in Bank-beg", format_amount(self.variables['cash_bank_beg'].get())],
                ["Cash on Hand-beg", format_amount(self.variables['cash_hand_beg'].get())]
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
                ["Monthly dues collected", format_amount(self.variables['monthly_dues'].get())],
                ["Certifications issued", format_amount(self.variables['certifications'].get())],
                ["Membership fee", format_amount(self.variables['membership_fee'].get())],
                ["Vehicle stickers", format_amount(self.variables['vehicle_stickers'].get())],
                ["Rentals", format_amount(self.variables['rentals'].get())],
                ["Solicitations/Donations", format_amount(self.variables['solicitations'].get())],
                ["Interest Income on bank deposits", format_amount(self.variables['interest_income'].get())],
                ["Livelihood Management Fee", format_amount(self.variables['livelihood_fee'].get())],
                ["Others(inflow)", format_amount(self.variables['inflows_others'].get())],
                ["Total Cash receipt", format_amount(self.variables['total_receipts'].get())]
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
                ["Cash Out Flows/Disbursements", format_amount(self.variables['cash_outflows'].get())],
                ["Snacks/Meals for visitors", format_amount(self.variables['snacks_meals'].get())],
                ["Transportation expenses", format_amount(self.variables['transportation'].get())],
                ["Office supplies expense", format_amount(self.variables['office_supplies'].get())],
                ["Printing and photocopy", format_amount(self.variables['printing'].get())],
                ["Labor", format_amount(self.variables['labor'].get())],
                ["Billboard expense", format_amount(self.variables['billboard'].get())],
                ["Clearing/cleaning charges", format_amount(self.variables['cleaning'].get())],
                ["Miscellaneous expenses", format_amount(self.variables['misc_expenses'].get())],
                ["Federation fee", format_amount(self.variables['federation_fee'].get())],
                ["HOA-BOD Uniforms", format_amount(self.variables['uniforms'].get())],
                ["BOD Mtg", format_amount(self.variables['bod_mtg'].get())],
                ["General Assembly", format_amount(self.variables['general_assembly'].get())],
                ["Cash Deposit to bank", format_amount(self.variables['cash_deposit'].get())],
                ["Withholding tax on bank deposit", format_amount(self.variables['withholding_tax'].get())],
                ["Refund", format_amount(self.variables['refund'].get())],
                ["Others(outflow)", format_amount(self.variables['outflows_others'].get())]
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
                ["Ending cash balance", format_amount(self.variables['ending_cash'].get())]
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
                ["Cash in Bank", format_amount(self.variables['ending_cash_bank'].get())],
                ["Cash on Hand", format_amount(self.variables['ending_cash_hand'].get())]
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
            return filename
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting to PDF: {str(e)}\n\nMake sure you have ReportLab installed by running:\npip install reportlab")
            return None

    def save_to_docx(self):
        try:
            def format_amount(value):
                if value:
                    try:
                        amount = Decimal(value.replace(',', ''))
                        return f"{amount:,.2f}"
                    except:
                        return value
                return ""
            default_filename = f"cash_flow_statement_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            filename = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx")],
                initialfile=default_filename,
                title="Save Word Document As"
            )
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
            doc.add_paragraph(f"For the year month {self.format_date_for_display(self.date_var.get())}")
            doc.add_heading("Beginning Cash Balances", level=2)
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            table.cell(0, 0).text = "Cash in Bank-beg"
            table.cell(0, 1).text = format_amount(self.variables['cash_bank_beg'].get())
            table.cell(1, 0).text = "Cash on Hand-beg"
            table.cell(1, 1).text = format_amount(self.variables['cash_hand_beg'].get())
            for row in table.rows:
                row.cells[1].paragraphs[0].alignment = 2
            doc.add_heading("Cash Inflows", level=2)
            table = doc.add_table(rows=10, cols=2)
            table.style = 'Table Grid'
            inflow_items = [
                ("Monthly dues collected", self.variables['monthly_dues']),
                ("Certifications issued", self.variables['certifications']),
                ("Membership fee", self.variables['membership_fee']),
                ("Vehicle stickers", self.variables['vehicle_stickers']),
                ("Rentals", self.variables['rentals']),
                ("Solicitations/Donations", self.variables['solicitations']),
                ("Interest Income on bank deposits", self.variables['interest_income']),
                ("Livelihood Management Fee", self.variables['livelihood_fee']),
                ("Others(inflow)", self.variables['inflows_others']),
                ("Total Cash receipt", self.variables['total_receipts'])
            ]
            for i, (label, var) in enumerate(inflow_items):
                table.cell(i, 0).text = label
                table.cell(i, 1).text = format_amount(var.get())
                table.cell(i, 1).paragraphs[0].alignment = 2
            table.cell(9, 0).paragraphs[0].runs[0].bold = True
            doc.add_heading("Less: Cash Outflows", level=2)
            table = doc.add_table(rows=17, cols=2)
            table.style = 'Table Grid'
            outflow_items = [
                ("Cash Out Flows/Disbursements", self.variables['cash_outflows']),
                ("Snacks/Meals for visitors", self.variables['snacks_meals']),
                ("Transportation expenses", self.variables['transportation']),
                ("Office supplies expense", self.variables['office_supplies']),
                ("Printing and photocopy", self.variables['printing']),
                ("Labor", self.variables['labor']),
                ("Billboard expense", self.variables['billboard']),
                ("Clearing/cleaning charges", self.variables['cleaning']),
                ("Miscellaneous expenses", self.variables['misc_expenses']),
                ("Federation fee", self.variables['federation_fee']),
                ("HOA-BOD Uniforms", self.variables['uniforms']),
                ("BOD Mtg", self.variables['bod_mtg']),
                ("General Assembly", self.variables['general_assembly']),
                ("Cash Deposit to bank", self.variables['cash_deposit']),
                ("Withholding tax on bank deposit", self.variables['withholding_tax']),
                ("Refund", self.variables['refund']),
                ("Others(outflow)", self.variables['outflows_others'])
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
            table.cell(0, 1).text = format_amount(self.variables['ending_cash'].get())
            table.cell(0, 1).paragraphs[0].alignment = 2
            table.cell(0, 0).paragraphs[0].runs[0].bold = True
            doc.add_heading("Breakdown of Cash", level=2)
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            table.cell(0, 0).text = "Cash in Bank"
            table.cell(0, 1).text = format_amount(self.variables['ending_cash_bank'].get())
            table.cell(1, 0).text = "Cash on Hand"
            table.cell(1, 1).text = format_amount(self.variables['ending_cash_hand'].get())
            for row in table.rows:
                row.cells[1].paragraphs[0].alignment = 2
            doc.save(filename)
            messagebox.showinfo("Success", f"Word document saved to {filename}")
            return filename
        except Exception as e:
            messagebox.showerror("Error", f"Error saving to Word: {str(e)}\n\nMake sure you have python-docx installed by running:\npip install python-docx")
            return None

    def load_from_documentpdf(self):
        try:
            filename = filedialog.askopenfilename(
                filetypes=[
                    ("Word Documents", "*.docx"),
                    ("PDF Files", "*.pdf"),
                ],
                title="Select a Document (DOCX or PDF)"
            )
            if not filename:
                return

            if filename.endswith('.pdf'):
                try:
                    self.load_from_pdf(filename)
                except Exception as e:
                    messagebox.showerror("Error", f"Error loading PDF: {str(e)}")
            elif filename.endswith('.docx'):
                try:
                    self.load_from_docx(filename)
                except Exception as e:
                    messagebox.showerror("Error", f"Error loading DOCX: {str(e)}")
            else:
                messagebox.showerror("Error", "Unsupported file format. Please select a PDF or DOCX file.")

            messagebox.showinfo("Success", f"Loaded data from {filename}")

        except Exception as e:
            messagebox.showerror("Error", f"Error loading document: {str(e)}")