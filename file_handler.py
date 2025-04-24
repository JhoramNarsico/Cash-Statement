import os
import re
from decimal import Decimal
from tkinter import filedialog, messagebox
import pdfplumber
from docx import Document
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from docx.shared import Inches, Pt
import sys
import datetime

class FileHandler:
    def __init__(self, variables, title_var, date_var, prepared_by_var, noted_by_var_1, noted_by_var_2, checked_by_var):
        self.variables = variables
        self.title_var = title_var
        self.date_var = date_var
        self.prepared_by_var = prepared_by_var
        self.noted_by_var_1 = noted_by_var_1
        self.noted_by_var_2 = noted_by_var_2
        self.checked_by_var = checked_by_var

    def parse_amount(self, text):
        """Parse text to extract numerical amount, removing non-numeric characters except decimal."""
        if not text or text.strip() == "":
            return ""
        try:
            return str(Decimal(re.sub(r'[^\d.]', '', text)))
        except:
            return text

    def format_date_for_display(self, date_str):
        """Convert mm/dd/yyyy to MMMM dd, yyyy for display."""
        try:
            date_obj = datetime.datetime.strptime(date_str, "%m/%d/%Y")
            return date_obj.strftime("%B %d, %Y")
        except ValueError:
            return date_str

    def load_from_docx(self, filename):
        """Load data from a Word document into the application variables."""
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

            # Extract footer names (Prepared by, Noted by, Checked by)
            footer = doc.sections[0].footer
            footer_text = ""
            for para in footer.paragraphs:
                footer_text += para.text + "\n"
            for line in footer_text.split("\n"):
                line = line.strip()
                if line.startswith("Prepared by:"):
                    self.prepared_by_var.set(line.replace("Prepared by:", "").strip())
                elif line.startswith("Noted by:"):
                    noted_name = line.replace("Noted by:", "").strip()
                    if not self.noted_by_var_1.get():
                        self.noted_by_var_1.set(noted_name)
                    elif not self.noted_by_var_2.get():
                        self.noted_by_var_2.set(noted_name)
                elif line.startswith("Checked by:"):
                    self.checked_by_var.set(line.replace("Checked by:", "").strip())

            messagebox.showinfo("Success", "Document loaded successfully")
            return True

        except Exception as e:
            messagebox.showerror("Error", f"Error loading Word document: {str(e)}")
            return False

    def load_from_pdf(self, filename):
        """Load data from a PDF document into the application variables."""
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

            # Extract date and footer names
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
                    elif line.startswith("Prepared by:"):
                        self.prepared_by_var.set(line.replace("Prepared by:", "").strip())
                    elif line.startswith("Noted by:"):
                        noted_name = line.replace("Noted by:", "").strip()
                        if not self.noted_by_var_1.get():
                            self.noted_by_var_1.set(noted_name)
                        elif not self.noted_by_var_2.get():
                            self.noted_by_var_2.set(noted_name)
                    elif line.startswith("Checked by:"):
                        self.checked_by_var.set(line.replace("Checked by:", "").strip())

            messagebox.showinfo("Success", "PDF data loaded successfully!")
            return True

        except Exception as e:
            messagebox.showerror("Error", f"Error loading PDF: {str(e)}")
            return False

    def export_to_pdf(self):
        """Export data to a single-page PDF matching the Word document format."""
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

            # Define custom page size: 8.5 x 13 inches (612 x 936

            folio_size = (8.5 * 72, 13 * 72)
            doc = SimpleDocTemplate(
                filename,
                pagesize=folio_size,
                topMargin=36,  # 0.5in
                bottomMargin=36,  # 0.5in
                leftMargin=36,  # 0.5in
                rightMargin=36  # 0.5in
            )
            styles = getSampleStyleSheet()

            # Create custom styles
            header_style = styles['Normal']
            header_style.alignment = 1  # Center
            header_style.fontSize = 10
            header_style.leading = 12
            header_style.fontName = 'Helvetica'

            title_style = styles['Normal']
            title_style.alignment = 1
            title_style.fontSize = 12
            title_style.leading = 14
            title_style.fontName = 'Helvetica-Bold'

            date_style = styles['Normal']
            date_style.alignment = 1
            date_style.fontSize = 8
            date_style.leading = 10
            date_style.fontName = 'Helvetica'

            table_style = styles['Normal']
            table_style.fontSize = 8
            table_style.leading = 10
            table_style.fontName = 'Helvetica'

            table_bold_style = styles['Normal']
            table_bold_style.fontSize = 8
            table_bold_style.leading = 10
            table_bold_style.fontName = 'Helvetica-Bold'

            footer_style = styles['Normal']
            footer_style.fontSize = 8
            footer_style.leading = 10
            footer_style.fontName = 'Helvetica'

            elements = []

            # Header (matching Word format)
            elements.append(Paragraph("Buena Oro Homeowners Association Inc.", header_style))
            elements.append(Paragraph("Macansandig, Cagayan de Oro City", header_style))
            elements.append(Paragraph("CASH FLOW STATEMENT", title_style))
            elements.append(Paragraph(f"For the Month of {self.format_date_for_display(self.date_var.get())}", date_style))
            elements.append(Spacer(1, 6))  # Reduced spacing to match Word

            # Beginning Cash Balances
            elements.append(Paragraph("Beginning Cash Balances", header_style))
            beg_data = [
                ["Cash in Bank-beg", format_amount(self.variables['cash_bank_beg'].get())],
                ["Cash on Hand-beg", format_amount(self.variables['cash_hand_beg'].get())]
            ]
            beg_data = [
                ["Cash in Bank-beg", format_amount(self.variables['cash_bank_beg'].get())],
                ["Cash on Hand-beg", format_amount(self.variables['cash_hand_beg'].get())]
            ]
            beg_table = Table(beg_data, colWidths=[360, 180], rowHeights=[14 for _ in range(len(beg_data))])

            beg_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (1, 0), (1, -1), 3),
            ]))
            elements.append(beg_table)
            elements.append(Spacer(1, 6))

            # Cash Inflows
            elements.append(Paragraph("Cash Inflows", header_style))
            inflows_data = [
                ["Monthly Dues Collected", format_amount(self.variables['monthly_dues'].get())],
                ["Certifications Issued", format_amount(self.variables['certifications'].get())],
                ["Membership Fee", format_amount(self.variables['membership_fee'].get())],
                ["Vehicle Stickers", format_amount(self.variables['vehicle_stickers'].get())],
                ["Rentals", format_amount(self.variables['rentals'].get())],
                ["Solicitations/Donations", format_amount(self.variables['solicitations'].get())],
                ["Interest Income on Bank Deposits", format_amount(self.variables['interest_income'].get())],
                ["Livelihood Management Fee", format_amount(self.variables['livelihood_fee'].get())],
                ["Others (Inflow)", format_amount(self.variables['inflows_others'].get())],
                ["Total Cash Receipts", format_amount(self.variables['total_receipts'].get())]
            ]
            inflows_table = Table(inflows_data, colWidths=[360, 180], rowHeights=[14.4]*10)
            inflows_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
                ('FONTNAME', (0, 0), (-1, -2), 'Helvetica'),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (1, 0), (1, -1), 3),
            ]))
            elements.append(inflows_table)
            elements.append(Spacer(1, 6))

            # Cash Outflows
            elements.append(Paragraph("Less: Cash Outflows", header_style))
            outflows_data = [
                ["Cash Outflows/Disbursements", format_amount(self.variables['cash_outflows'].get())],
                ["Snacks/Meals for Visitors", format_amount(self.variables['snacks_meals'].get())],
                ["Transportation Expenses", format_amount(self.variables['transportation'].get())],
                ["Office Supplies Expense", format_amount(self.variables['office_supplies'].get())],
                ["Printing and Photocopy", format_amount(self.variables['printing'].get())],
                ["Labor", format_amount(self.variables['labor'].get())],
                ["Billboard Expense", format_amount(self.variables['billboard'].get())],
                ["Clearing/Cleaning Charges", format_amount(self.variables['cleaning'].get())],
                ["Miscellaneous Expenses", format_amount(self.variables['misc_expenses'].get())],
                ["Federation Fee", format_amount(self.variables['federation_fee'].get())],
                ["HOA-BOD Uniforms", format_amount(self.variables['uniforms'].get())],
                ["BOD Meeting", format_amount(self.variables['bod_mtg'].get())],
                ["General Assembly", format_amount(self.variables['general_assembly'].get())],
                ["Cash Deposit to Bank", format_amount(self.variables['cash_deposit'].get())],
                ["Withholding Tax on Bank Deposit", format_amount(self.variables['withholding_tax'].get())],
                ["Refund", format_amount(self.variables['refund'].get())],
                ["Others (Outflow)", format_amount(self.variables['outflows_others'].get())]
            ]
            outflows_table = Table(outflows_data, colWidths=[360, 180], rowHeights=[14.4]*17)
            outflows_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (0, 0), colors.lightgrey),
                ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (1, 0), (1, -1), 3),
            ]))
            elements.append(outflows_table)
            elements.append(Spacer(1, 6))

            # Ending Cash Balance
            elements.append(Paragraph("Ending Cash Balance", header_style))
            ending_data = [
                ["Ending Cash Balance", format_amount(self.variables['ending_cash'].get())]
            ]
            ending_table = Table(ending_data, colWidths=[360, 180], rowHeights=[14.4])
            ending_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (-1, -1), colors.lightgrey),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (1, 0), (1, -1), 3),
            ]))
            elements.append(ending_table)
            elements.append(Spacer(1, 6))

            # Breakdown of Cash
            elements.append(Paragraph("Breakdown of Cash", header_style))
            breakdown_data = [
                ["Cash in Bank", format_amount(self.variables['ending_cash_bank'].get())],
                ["Cash on Hand", format_amount(self.variables['ending_cash_hand'].get())]
            ]
            breakdown_table = Table(breakdown_data, colWidths=[360, 180], rowHeights=[14.4, 14.4])
            breakdown_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (1, 0), (1, -1), 3),
            ]))
            elements.append(breakdown_table)
            elements.append(Spacer(1, 12))

            # Footer (matching Word format with tab stops)
            prepared_name = self.prepared_by_var.get() or "_______________________"
            noted_name_1 = self.noted_by_var_1.get() or "_______________________"
            noted_name_2 = self.noted_by_var_2.get() or "_______________________"
            checked_name = self.checked_by_var.get() or "_______________________"
            
            # Simulate tab stops using spaces and alignment
            footer_para1 = Paragraph(
                f"Prepared by: {prepared_name}{' ' * 50}Checked by: {checked_name}",
                footer_style
            )
            footer_para2 = Paragraph(
                f"Noted by: {noted_name_1}{' ' * 50}Noted by: {noted_name_2}",
                footer_style
            )
            elements.append(footer_para1)
            elements.append(footer_para2)

            doc.build(elements)
            messagebox.showinfo("Success", f"PDF successfully exported to {filename}")
            return filename

        except Exception as e:
            messagebox.showerror("Error", f"Error exporting to PDF: {str(e)}\n\nMake sure you have ReportLab installed by running:\npip install reportlab")
            return None

    def save_to_docx(self):
        """Save data to a Word document with two Noted by fields in the footer."""
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
            from docx.oxml.ns import qn
            from docx.shared import Inches, Pt
            section = doc.sections[0]
            section.page_width = Inches(8.5)
            section.page_height = Inches(13)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(1.2)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            # Footer with names
            footer = section.footer
            prepared_name = self.prepared_by_var.get() or "_______________________"
            noted_name_1 = self.noted_by_var_1.get() or "_______________________"
            noted_name_2 = self.noted_by_var_2.get() or "_______________________"
            checked_name = self.checked_by_var.get() or "_______________________"
            footer_para1 = footer.add_paragraph()
            footer_para1.alignment = 0
            run = footer_para1.add_run(f"Prepared by: {prepared_name}\tChecked by: {checked_name}")
            run.font.size = Pt(8)
            tab_stops = footer_para1.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(7.5), alignment=2)
            footer_para2 = footer.add_paragraph()
            footer_para2.alignment = 0
            run = footer_para2.add_run(f"Noted by: {noted_name_1}\tNoted by: {noted_name_2}")
            run.font.size = Pt(8)
            tab_stops = footer_para2.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(7.5), alignment=2)
            # Header
            p = doc.add_paragraph()
            run = p.add_run("Buena Oro Homeowners Association Inc.")
            run.font.size = Pt(10)
            p.alignment = 1
            p = doc.add_paragraph()
            run = p.add_run("Macansandig, Cagayan de Oro City")
            run.font.size = Pt(10)
            p.alignment = 1
            p = doc.add_paragraph()
            run = p.add_run("CASH FLOW STATEMENT")
            run.bold = True
            run.font.size = Pt(12)
            p.alignment = 1
 # The 'get()' method is used to retrieve the current string value from a Tkinter StringVar object.
# It is not a standard Python method, but specific to Tkinter and similar frameworks.
# For instance, in this context, 'self.variables' is a dictionary containing Tkinter StringVar objects.
# The '.get()' method is used to fetch the current value of these variables.

            p = doc.add_paragraph()
            run = p.add_run(f"For the Month of {self.format_date_for_display(self.date_var.get())}")
            run.font.size = Pt(8)
            p.alignment = 1
            # Beginning Cash Balances
            doc.add_heading("Beginning Cash Balances", level=2).style.font.size = Pt(10)
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            table.autofit = True
            table.columns[0].width = Inches(5.0)
            table.columns[1].width = Inches(2.5)
            for row in table.rows:
                row.height = Inches(0.2)
            table.cell(0, 0).text = "Cash in Bank-beg"
            table.cell(0, 1).text = format_amount(self.variables['cash_bank_beg'].get())
            table.cell(1, 0).text = "Cash on Hand-beg"
            table.cell(1, 1).text = format_amount(self.variables['cash_hand_beg'].get())
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    cell.paragraphs[0].runs[0].font.size = Pt(8)
                    cell.paragraphs[0].alignment = 2 if j == 1 else 0
            # Cash Inflows
            doc.add_heading("Cash Inflows", level=2).style.font.size = Pt(10)
            table = doc.add_table(rows=10, cols=2)
            table.style = 'Table Grid'
            table.autofit = True
            table.columns[0].width = Inches(5.0)
            table.columns[1].width = Inches(2.5)
            for row in table.rows:
                row.height = Inches(0.2)
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
                for j, cell in enumerate(table.rows[i].cells):
                    cell.paragraphs[0].runs[0].font.size = Pt(8)
                    cell.paragraphs[0].alignment = 2 if j == 1 else 0
            table.cell(9, 0).paragraphs[0].runs[0].bold = True
            # Cash Outflows
            doc.add_heading("Less: Cash Outflows", level=2).style.font.size = Pt(10)
            table = doc.add_table(rows=17, cols=2)
            table.style = 'Table Grid'
            table.autofit = True
            table.columns[0].width = Inches(5.0)
            table.columns[1].width = Inches(2.5)
            for row in table.rows:
                row.height = Inches(0.2)
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
                for j, cell in enumerate(table.rows[i].cells):
                    cell.paragraphs[0].runs[0].font.size = Pt(8)
                    cell.paragraphs[0].alignment = 2 if j == 1 else 0
            table.cell(0, 0).paragraphs[0].runs[0].bold = True
            # Ending Cash Balance
            doc.add_heading("Ending Cash Balance", level=2).style.font.size = Pt(10)
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            table.autofit = True
            table.columns[0].width = Inches(5.0)
            table.columns[1].width = Inches(2.5)
            table.rows[0].height = Inches(0.2)
            table.cell(0, 0).text = "Ending cash balance"
            table.cell(0, 1).text = format_amount(self.variables['ending_cash'].get())
            for j, cell in enumerate(table.rows[0].cells):
                cell.paragraphs[0].runs[0].font.size = Pt(8)
                cell.paragraphs[0].alignment = 2 if j == 1 else 0
            table.cell(0, 0).paragraphs[0].runs[0].bold = True
            # Breakdown of Cash
            doc.add_heading("Breakdown of Cash", level=2).style.font.size = Pt(10)
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            table.autofit = True
            table.columns[0].width = Inches(5.0)
            table.columns[1].width = Inches(2.5)
            for row in table.rows:
                row.height = Inches(0.2)
            table.cell(0, 0).text = "Cash in Bank"
            table.cell(0, 1).text = format_amount(self.variables['ending_cash_bank'].get())
            table.cell(1, 0).text = "Cash on Hand"
            table.cell(1, 1).text = format_amount(self.variables['ending_cash_hand'].get())
            for i, row in enumerate(table.rows):
                for j, cell in enumerate(row.cells):
                    cell.paragraphs[0].runs[0].font.size = Pt(8)
                    cell.paragraphs[0].alignment = 2 if j == 1 else 0
            doc.save(filename)
            messagebox.showinfo("Success", f"Word document saved to {filename}")
            return filename
        except Exception as e:
            messagebox.showerror("Error", f"Error saving to Word: {str(e)}\n\nMake sure you have python-docx installed by running:\npip install python-docx")
            return None

    def load_from_documentpdf(self):
        """Load data from either a Word or PDF document."""
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