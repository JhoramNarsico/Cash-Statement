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
from docx.shared import Inches, Pt
import sys
import datetime

class FileHandler:
    def __init__(self, variables, title_var, date_var, prepared_by_var, noted_by_var, checked_by_var):
        self.variables = variables
        self.title_var = title_var
        self.date_var = date_var
        self.prepared_by_var = prepared_by_var
        self.noted_by_var = noted_by_var
        self.checked_by_var = checked_by_var  # New variable

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
            # Define custom page size: 8.5 x 13 inches (612 x 936 points)
            folio_size = (8.5 * 72, 13 * 72)
            doc = SimpleDocTemplate(
                filename,
                pagesize=folio_size,
                topMargin=36,
                bottomMargin=90,  # Increased to accommodate extra footer line
                leftMargin=36,
                rightMargin=36
            )
            styles = getSampleStyleSheet()
            # Create a custom style for centered normal text
            normal_centered_style = styles['Normal']
            normal_centered_style.alignment = 1
            normal_centered_style.fontSize = 10
            # Create a custom style for bold centered text
            bold_centered_style = styles['Normal']
            bold_centered_style.alignment = 1
            bold_centered_style.fontSize = 12
            bold_centered_style.fontName = 'Helvetica-Bold'
            # Create a custom style for minimized centered text
            minimized_centered_style = styles['Normal']
            minimized_centered_style.alignment = 1
            minimized_centered_style.fontSize = 8
            # Create a custom style for footer (left-aligned for "Prepared by" and "Checked by")
            footer_style = styles['Normal']
            footer_style.alignment = 0  # Left-aligned for better control in table
            footer_style.fontSize = 8
            # Create a custom style for centered footer ("Noted by")
            footer_centered_style = styles['Normal']
            footer_centered_style.alignment = 1  # Center-aligned
            footer_centered_style.fontSize = 8
            elements = []
            # Calculate the available width for the header (page width minus margins)
            page_width = folio_size[0] - (doc.leftMargin + doc.rightMargin)  # 612 - (36 + 36) = 540 points
            # Add the header using a table to ensure perfect centering
            header_data = [
                [Paragraph("Buena Oro Homeowners Association Inc.", normal_centered_style)],
                [Paragraph("Macansandig, Cagayan de Oro City", normal_centered_style)],
                [Paragraph("CASH FLOW STATEMENT", bold_centered_style)],
                [Paragraph(f"For the Month of {self.format_date_for_display(self.date_var.get())}", minimized_centered_style)],
            ]
            header_table = Table(header_data, colWidths=[page_width], rowHeights=[20, 20, 24, 18])
            header_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ]))
            elements.append(header_table)
            elements.append(Spacer(1, 12))
            # Beginning Cash Balances
            beg_data = [
                ["Cash in Bank-beg", format_amount(self.variables['cash_bank_beg'].get())],
                ["Cash on Hand-beg", format_amount(self.variables['cash_hand_beg'].get())]
            ]
            beg_table = Table(beg_data, colWidths=[300, 150], rowHeights=[18, 18])
            beg_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))
            elements.append(beg_table)
            elements.append(Spacer(1, 6))
            elements.append(Paragraph("<b>Cash inflows:</b>", styles['Normal']))
            elements.append(Spacer(1, 4))
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
            inflows_table = Table(inflows_data, colWidths=[300, 150], rowHeights=[18]*10)
            inflows_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))
            elements.append(inflows_table)
            elements.append(Spacer(1, 6))
            elements.append(Paragraph("<b>Less:</b>", styles['Normal']))
            elements.append(Spacer(1, 4))
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
            outflows_table = Table(outflows_data, colWidths=[300, 150], rowHeights=[18]*17)
            outflows_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (0, 0), colors.lightgrey),
                ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))
            elements.append(outflows_table)
            elements.append(Spacer(1, 6))
            ending_data = [
                ["Ending cash balance", format_amount(self.variables['ending_cash'].get())]
            ]
            ending_table = Table(ending_data, colWidths=[300, 150], rowHeights=[18])
            ending_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (-1, -1), colors.lightgrey),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))
            elements.append(ending_table)
            elements.append(Spacer(1, 6))
            elements.append(Paragraph("<b>Breakdown of cash:</b>", styles['Normal']))
            elements.append(Spacer(1, 4))
            breakdown_data = [
                ["Cash in Bank", format_amount(self.variables['ending_cash_bank'].get())],
                ["Cash on Hand", format_amount(self.variables['ending_cash_hand'].get())]
            ]
            breakdown_table = Table(breakdown_data, colWidths=[300, 150], rowHeights=[18, 18])
            breakdown_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))
            elements.append(breakdown_table)
            # Add footer with user-provided names in the desired layout
            elements.append(Spacer(1, 24))
            prepared_name = self.prepared_by_var.get() or "_______________________"
            noted_name = self.noted_by_var.get() or "_______________________"
            checked_name = self.checked_by_var.get() or "_______________________"
            # Create a table for "Prepared by" and "Checked by" on the same line
            footer_top_data = [
                [f"Prepared by: {prepared_name}", f"Checked by: {checked_name}"]
            ]
            footer_top_table = Table(footer_top_data, colWidths=[225, 225], rowHeights=[18])
            footer_top_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))
            elements.append(footer_top_table)
            # Add "Noted by" on a separate line, centered
            elements.append(Paragraph(f"Noted by: {noted_name}", footer_centered_style))
            def add_page_numbers_and_footer(canvas, doc):
                page_num = canvas.getPageNumber()
                text = f"Page {page_num}"
                canvas.drawRightString(200, 20, text)
            doc.build(elements, onFirstPage=add_page_numbers_and_footer, onLaterPages=add_page_numbers_and_footer)
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
            # Set page size to 8.5 x 13 inches (Folio)
            from docx.oxml.ns import qn
            from docx.shared import Inches, Pt
            section = doc.sections[0]
            section.page_width = Inches(8.5)
            section.page_height = Inches(13)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(1.2)  # Increased to accommodate extra footer line
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
            # Add footer with user-provided names in the desired layout using paragraphs and tab stops
            footer = section.footer
            prepared_name = self.prepared_by_var.get() or "_______________________"
            noted_name = self.noted_by_var.get() or "_______________________"
            checked_name = self.checked_by_var.get() or "_______________________"
            # First line: "Prepared by:" and "Checked by:" side by side using a tab stop
            footer_para1 = footer.add_paragraph()
            footer_para1.alignment = 0  # Left-aligned
            run = footer_para1.add_run(f"Prepared by: {prepared_name}\tChecked by: {checked_name}")
            run.font.size = Pt(8)
            # Set a right-aligned tab stop to position "Checked by:" on the right
            tab_stops = footer_para1.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Inches(7.5), alignment=2)  # Right-aligned tab at 7.5 inches (page width - margins)
            # Second line: "Noted by:", centered
            footer_para2 = footer.add_paragraph()
            footer_para2.alignment = 1  # Center-aligned
            run = footer_para2.add_run(f"Noted by: {noted_name}")
            run.font.size = Pt(8)
            # Add the header
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
            p = doc.add_paragraph()
            run = p.add_run(f"For the Month of {self.format_date_for_display(self.date_var.get())}")
            run.font.size = Pt(8)
            p.alignment = 1
            # Beginning Cash Balances
            doc.add_heading("Beginning Cash Balances", level=2).style.font.size = Pt(10)
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            table.autofit = True
            # Set column widths (total width = 7.5 inches, roughly 2:1 ratio)
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
            # Set column widths
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
            # Less: Cash Outflows
            doc.add_heading("Less: Cash Outflows", level=2).style.font.size = Pt(10)
            table = doc.add_table(rows=17, cols=2)
            table.style = 'Table Grid'
            table.autofit = True
            # Set column widths
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
            # Set column widths
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
            # Set column widths
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