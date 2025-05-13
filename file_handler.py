import os
import re
import sys 
import datetime
import logging 
from decimal import Decimal
from tkinter import filedialog, messagebox 
import tempfile # Added for temporary files

# PDF Imports
import pdfplumber
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image 
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter 
from reportlab.lib.units import inch 
import camelot
import fitz
import random

# Word Imports
from docx import Document
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Inches, Pt
from docx.oxml.ns import qn 
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT  # Add import at the top of the file
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

class FileHandler:
    def __init__(self, variables, title_var, date_var, logo_path_var, address_var, prepared_by_var, noted_by_var_1, noted_by_var_2, checked_by_var):
        self.variables = variables
        self.title_var = title_var
        self.date_var = date_var
        self.logo_path_var = logo_path_var 
        self.address_var = address_var   
        self.prepared_by_var = prepared_by_var
        self.noted_by_var_1 = noted_by_var_1
        self.noted_by_var_2 = noted_by_var_2
        self.checked_by_var = checked_by_var

    def parse_amount(self, text):
        if not text or text.strip() == "":
            return ""
        try:
            cleaned_text = re.sub(r'[^\d.]', '', text.replace(',', ''))
            return str(Decimal(cleaned_text))
        except Exception as e:
            logging.warning(f"Could not parse amount '{text}': {e}")
            return text

    def format_date_for_display(self, date_str):
        try:
            date_obj = datetime.datetime.strptime(date_str, "%m/%d/%Y")
            return date_obj.strftime("%B %d, %Y")
        except ValueError:
            logging.warning(f"Invalid date format for display: {date_str}")
            return date_str

    # --- Internal PDF Content Generation ---
    def _build_pdf_elements(self):
        """Builds the list of ReportLab Platypus elements for the PDF content."""
        elements = []
        
        def format_amount_pdf(value): # Renamed to avoid conflict if used elsewhere
            if value:
                try:
                    str_value = str(value).replace(',', '')
                    if not str_value: return ""
                    amount = Decimal(str_value)
                    return f"{amount:,.2f}"
                except Exception as e:
                    logging.warning(f"Could not format amount '{value}' for PDF: {e}")
                    return str(value) 
            return ""

        styles = getSampleStyleSheet()
        folio_size = (8.5 * inch, 13 * inch) # Needed for page_width calculation
         # Assuming doc margins are 0.5 inch each side for page_width calculation here
        doc_left_margin = 0.5 * inch
        doc_right_margin = 0.5 * inch


        header_style = styles['Normal']
        header_style.alignment = 1
        header_style.fontSize = 10
        header_style.leading = 12
        header_style.fontName = 'Helvetica'

        addressStyle = ParagraphStyle(name='addressStyle', fontName='Helvetica', fontSize=10, leading=14, alignment=1)
        titleStyle = ParagraphStyle(name='titleStyle', fontName='Helvetica-Bold', fontSize=12, leading=14, alignment=1)
        dateStyle = ParagraphStyle(name='dateStyle', fontName='Helvetica', fontSize=8, leading=10, alignment=1, spaceBefore=6)
        tableboldStyle = ParagraphStyle(name='tableBoldStyle', fontName='Helvetica-Bold', fontSize=10, leading=10, spaceAfter=10, spaceBefore=4)
        footerStyle = ParagraphStyle(name='footerStyle', fontName='Helvetica', fontSize=8, leading=10, alignment=1)
        notedStyle = ParagraphStyle(name='notedStyle', fontName='Helvetica', fontSize=8, leading=12, alignment=1)

        logo_path = self.logo_path_var.get()
        address_text = self.address_var.get() or " "

        logo_img = None
        logo_placeholder_text = ""
        if logo_path and os.path.exists(logo_path):
            try:
                target_w = 1.18 * inch
                target_h = 1.18 * inch
                logo_img = Image(logo_path, width=target_w, height=target_h)
                logo_img.hAlign = 'CENTER'
                logo_img.vAlign = 'MIDDLE'
            except Exception as e:
                logging.warning(f"Could not load or process logo image '{logo_path}': {e}")
                logo_placeholder_text = "[Logo Error]"
        elif logo_path:
            logo_placeholder_text = "[Logo N/A]"
        
        logo_cell_content = Paragraph(logo_placeholder_text, styles['Italic']) if logo_img is None else logo_img

        header_text_elements = [
            Paragraph(address_text, addressStyle),
            Spacer(1, 12),
            Paragraph("CASH FLOW STATEMENT", titleStyle),
            Paragraph(f"For the Month of {self.format_date_for_display(self.date_var.get())}", dateStyle)
        ]
        
        page_width = folio_size[0] - doc_left_margin - doc_right_margin
        logo_col_width = 1.58 * inch
        text_col_width = 4.5 * inch

        header_table_data = [[logo_cell_content, header_text_elements]]
        header_table = Table(header_table_data, colWidths=[logo_col_width, text_col_width], hAlign='LEFT', vAlign='LEFT')
        header_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (1, 0), (1, 0), 'CENTER'),
            ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 0)
        ]))
        elements.append(header_table)
        elements.append(Spacer(1, 12))

        data_label_width = page_width * 0.65
        data_value_width = page_width * 0.35
        
        common_table_style = TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 5), ('RIGHTPADDING', (1,0), (1,-1), 5)
        ])
        
        # Beginning Cash Balances
        elements.append(Paragraph("Beginning Cash Balances", tableboldStyle))
        beg_data = [
            ["Cash in Bank-beg", format_amount_pdf(self.variables['cash_bank_beg'].get())],
            ["Cash on Hand-beg", format_amount_pdf(self.variables['cash_hand_beg'].get())]
        ]
        beg_table = Table(beg_data, colWidths=[data_label_width, data_value_width], rowHeights=[14 for _ in beg_data])
        beg_table.setStyle(common_table_style)
        elements.append(beg_table)
        elements.append(Spacer(1, 6))

        # Cash Inflows
        elements.append(Paragraph("Cash Inflows", tableboldStyle))
        inflows_data = [
            ["Monthly Dues Collected", format_amount_pdf(self.variables['monthly_dues'].get())],
            ["Certifications Issued", format_amount_pdf(self.variables['certifications'].get())],
            ["Membership Fee", format_amount_pdf(self.variables['membership_fee'].get())],
            ["Vehicle Stickers", format_amount_pdf(self.variables['vehicle_stickers'].get())],
            ["Rentals", format_amount_pdf(self.variables['rentals'].get())],
            ["Solicitations/Donations", format_amount_pdf(self.variables['solicitations'].get())],
            ["Interest Income on Bank Deposits", format_amount_pdf(self.variables['interest_income'].get())],
            ["Livelihood Management Fee", format_amount_pdf(self.variables['livelihood_fee'].get())],
            ["Others (Inflow)", format_amount_pdf(self.variables['inflows_others'].get())],
            ["Total Cash Receipts", format_amount_pdf(self.variables['total_receipts'].get())]
        ]
        inflows_table = Table(inflows_data, colWidths=[data_label_width, data_value_width], rowHeights=[14]*len(inflows_data))
        inflows_table.setStyle(common_table_style)
        elements.append(inflows_table)
        elements.append(Spacer(1, 6))

        # Cash Outflows
        elements.append(Paragraph("Less: Cash Outflows", tableboldStyle))
        outflows_data = [
            ["Snacks/Meals for Visitors", format_amount_pdf(self.variables['snacks_meals'].get())],
            ["Transportation Expenses", format_amount_pdf(self.variables['transportation'].get())],
            ["Office Supplies Expense", format_amount_pdf(self.variables['office_supplies'].get())],
            ["Printing and Photocopy", format_amount_pdf(self.variables['printing'].get())],
            ["Labor", format_amount_pdf(self.variables['labor'].get())],
            ["Billboard Expense", format_amount_pdf(self.variables['billboard'].get())],
            ["Clearing/Cleaning Charges", format_amount_pdf(self.variables['cleaning'].get())],
            ["Miscellaneous Expenses", format_amount_pdf(self.variables['misc_expenses'].get())],
            ["Federation Fee", format_amount_pdf(self.variables['federation_fee'].get())],
            ["HOA-BOD Uniforms", format_amount_pdf(self.variables['uniforms'].get())],
            ["BOD Meeting", format_amount_pdf(self.variables['bod_mtg'].get())],
            ["General Assembly", format_amount_pdf(self.variables['general_assembly'].get())],
            ["Cash Deposit to Bank", format_amount_pdf(self.variables['cash_deposit'].get())],
            ["Withholding Tax on Bank Deposit", format_amount_pdf(self.variables['withholding_tax'].get())],
            ["Refund", format_amount_pdf(self.variables['refund'].get())],
            ["Others (Outflow)", format_amount_pdf(self.variables['outflows_others'].get())],
            ["Cash Outflows/Disbursements", format_amount_pdf(self.variables['cash_outflows'].get())]
        ]
        outflows_table = Table(outflows_data, colWidths=[data_label_width, data_value_width], rowHeights=[14]*len(outflows_data))
        outflows_table.setStyle(common_table_style)
        elements.append(outflows_table)
        elements.append(Spacer(1, 6))
        
        # Ending Cash Balance
        elements.append(Paragraph("Ending Cash Balance", tableboldStyle))
        ending_data = [["Ending Cash Balance", format_amount_pdf(self.variables['ending_cash'].get())]]
        ending_table = Table(ending_data, colWidths=[data_label_width, data_value_width], rowHeights=[14])
        ending_table.setStyle(common_table_style)
        elements.append(ending_table)
        elements.append(Spacer(1, 6))

        # Breakdown of Cash
        elements.append(Paragraph("Breakdown of Cash", tableboldStyle))
        breakdown_data = [
            ["Cash in Bank", format_amount_pdf(self.variables['ending_cash_bank'].get())],
            ["Cash on Hand", format_amount_pdf(self.variables['ending_cash_hand'].get())]
        ]
        breakdown_table = Table(breakdown_data, colWidths=[data_label_width, data_value_width], rowHeights=[14, 14])
        breakdown_table.setStyle(common_table_style)
        elements.append(breakdown_table)
        elements.append(Spacer(1, 12))
        
        # Footer Signatories
        prepared_name = self.prepared_by_var.get() or "_______________________"
        noted_name_1 = self.noted_by_var_1.get() or "_______________________"
        noted_name_2 = self.noted_by_var_2.get() or "_______________________"
        checked_name = self.checked_by_var.get() or "_______________________"
        
        col_width = (page_width / 2) - (0.1 * inch)
        
        prep_check_data = [
            [Paragraph(f"Prepared by:<br/>{prepared_name}<br/>HOA Treasurer", footerStyle)],
            [Paragraph(f"Checked by:<br/>{checked_name}<br/>HOA Auditor", footerStyle)]
        ]
        prep_check_table = Table([prep_check_data[0] + prep_check_data[1]], colWidths=[col_width, col_width])
        prep_check_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0),(-1,-1), 'TOP'), ('LEFTPADDING', (0,0),(-1,-1),0), ('RIGHTPADDING', (0,0),(-1,-1),0)]))
        elements.append(prep_check_table)
        elements.append(Spacer(1, 12))
        
        elements.append(Paragraph("Noted by:", notedStyle))
        elements.append(Spacer(1, 6))
        
        noted_data = [
            [Paragraph(f"{noted_name_1}<br/>HOA President", footerStyle)],
            [Paragraph(f"{noted_name_2}<br/>CHUDD HCD-CORDS", footerStyle)]
        ]
        noted_table = Table([noted_data[0] + noted_data[1]], colWidths=[col_width, col_width])
        noted_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0),(-1,-1), 'TOP'), ('LEFTPADDING', (0,0),(-1,-1),0), ('RIGHTPADDING', (0,0),(-1,-1),0)]))
        elements.append(noted_table)
        
        return elements

    def _create_pdf_at_path(self, filename):
        """Internal method to generate PDF content and save to the given filename."""
        try:
            folio_size = (8.5 * inch, 13 * inch)
            doc = SimpleDocTemplate(
                filename,
                pagesize=folio_size,
                topMargin=0.5*inch,
                bottomMargin=0.5*inch,
                leftMargin=0.5*inch,
                rightMargin=0.5*inch
            )
            elements = self._build_pdf_elements()
            doc.build(elements)
            return True 
        except Exception as e:
            logging.exception(f"Error creating PDF content for {filename}")
            return False

    def export_to_pdf(self):
        """Prompts user for save location and exports data to PDF.
           Returns a dictionary with 'status' and 'message'."""
        try:
            default_filename = f"cash_flow_statement_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            filename = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile=default_filename,
                title="Save PDF As"
            )
            if not filename:
                return {"status": "cancelled", "message": "Export to PDF cancelled by user."}

            if self._create_pdf_at_path(filename):
                logging.info(f"PDF successfully exported to {filename}")
                return {"status": "success", "message": f"PDF successfully exported to {filename}", "filename": filename}
            else:
                return {"status": "error", "message": f"Failed to create PDF at {filename}.\nCheck logs for details and ensure ReportLab is installed."}
        except Exception as e: 
            logging.exception("Error during PDF export process")
            return {"status": "error", "message": f"An unexpected error occurred during PDF export: {str(e)}"}

    def generate_temp_pdf(self):
        """Generates a PDF in a temporary location.
           Returns a dictionary with 'status', 'filename', and 'message'."""
        try:
            # Create a temporary file that is not deleted on close, so ReportLab can write to it by name
            fd, temp_filename = tempfile.mkstemp(suffix=".pdf", prefix="cash_flow_")
            os.close(fd) # Close the file descriptor, ReportLab will open it by name

            if self._create_pdf_at_path(temp_filename):
                logging.info(f"Temporary PDF for email generated at {temp_filename}")
                return {"status": "success", "filename": temp_filename}
            else:
                if os.path.exists(temp_filename):
                    try: os.remove(temp_filename)
                    except Exception as e_del: logging.error(f"Could not remove temp PDF {temp_filename} after generation failure: {e_del}")
                return {"status": "error", "message": "Failed to generate temporary PDF."}
        except Exception as e:
            logging.exception("Error generating temporary PDF")
            return {"status": "error", "message": f"An unexpected error occurred while generating temporary PDF: {str(e)}"}

    # --- Internal DOCX Content Generation ---
    def _build_docx_document(self):
        """Builds and returns a python-docx Document object."""
        doc = Document()
        
        def format_amount_docx(value): # Renamed to avoid conflict
            if value:
                try:
                    str_value = str(value).replace(',', '')
                    if not str_value: return ""
                    amount = Decimal(str_value)
                    return f"{amount:,.2f}"
                except Exception as e:
                    logging.warning(f"Could not format amount '{value}' for Word: {e}")
                    return str(value)
            return ""
        
        # Page Setup
        section = doc.sections[0]
        section.page_width = Inches(8.5)
        section.page_height = Inches(13)
        section.top_margin = Inches(0.5); section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.5); section.right_margin = Inches(0.5)

            # --- Header Section ---
        header = section.header
        header.is_linked_to_previous = False  # Ensure header is unique to this section
# Clear existing default paragraph in header
        if header.paragraphs:
            ht = header.paragraphs[0]._element
            ht.getparent().remove(ht)

# Get logo path and address
        logo_path = self.logo_path_var.get()
        address_text = self.address_var.get() or " "

# Create a 1x2 table in the header
        header_table = header.add_table(rows=1, cols=2, width=Inches(6.08))  # Width = Page Width - Margins
        header_table.autofit = False
        header_table.allow_autofit = False
        header_table.alignment = WD_TABLE_ALIGNMENT.LEFT

# Set column widths
        header_table.columns[0].width = Inches(1.58)  # Width for logo
        header_table.columns[1].width = Inches(4.5)   # Width for text

# Set cell widths explicitly
        logo_cell = header_table.cell(0, 0)
        text_cell = header_table.cell(0, 1)
        logo_cell.width = Inches(1.58)
        text_cell.width = Inches(4.5)

# Set vertical alignment
        logo_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        text_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

# Left Cell (Logo)
# Remove default paragraph in cell
        if logo_cell.paragraphs:
            p = logo_cell.paragraphs[0]._element
            p.getparent().remove(p)
# Add logo if path exists
        logo_para = logo_cell.add_paragraph()
        logo_para.alignment = 1  # Center logo in its cell
        if logo_path and os.path.exists(logo_path):
            try:
                logo_run = logo_para.add_run()
                    # Scale logo to fit within 1.58-inch column, preserving aspect ratio
                logo_run.add_picture(logo_path, width=Inches(1.18))  # Slightly less than column width
                logging.info(f"Included logo in Word header: {logo_path}, scaled to width 1.5 inches")
            except Exception as e:
                logging.warning(f"Could not add logo picture to Word: {e}")
                logo_para.add_run("[Logo Error]").italic = True
        elif logo_path:
            logging.warning(f"Logo path specified but not found for Word: {logo_path}")
            logo_para.add_run("[Logo N/A]").italic = True
# Else: leave the cell empty

# Right Cell (Text)
# Remove default paragraph
        if text_cell.paragraphs:
            p = text_cell.paragraphs[0]._element
            p.getparent().remove(p)

# Add Header Text Lines
        p = text_cell.add_paragraph()
        run = p.add_run(address_text)  # Add address
        run.font.name = 'Helvetica'
        run.font.size = Pt(10)
        p.alignment = 1  # Center
  # Small spacer
        p.add_run("\n")
        run = p.add_run("\nCASH FLOW STATEMENT")
        run.font.name = 'Helvetica'
        run.bold = True
        run.font.size = Pt(12)
        p.alignment = 1  # Center

        p = text_cell.add_paragraph()
        run = p.add_run(f"For the Month of {self.format_date_for_display(self.date_var.get())}")
        run.font.name = 'Helvetica'
        run.font.size = Pt(8)
        p.alignment = 1  # Center

        def set_cell_style(cell, text, size=8, bold=False, align='left', font='Helvetica'):
                para = cell.paragraphs[0]
                # Clear existing runs if needed, or just set text on the first paragraph
                if not para.runs:
                    para.add_run(text)
                else:
                    para.text = text # Overwrite if text exists

                # Ensure there's at least one run after setting text
                if not para.runs:
                     para.add_run(text) # Add run again if setting para.text cleared it

                run = para.runs[0]
                run.font.name = font
                run.font.size = Pt(size)
                run.bold = bold
                # Alignment: 0=left, 1=center, 2=right
                if align == 'right': para.alignment = 2
                elif align == 'center': para.alignment = 1
                else: para.alignment = 0
                
        # Content tables (width 7.5 inches)
        table_col1_width = Inches(6.0)
        table_col2_width = Inches(1.5)

        # Beginning Cash Balances
        p_beg = doc.add_paragraph(); run_beg = p_beg.add_run("Beginning Cash Balances")
        run_beg.font.name = 'Helvetica'; run_beg.bold = True; run_beg.font.size = Pt(10)
        beg_table = doc.add_table(rows=2, cols=2)
        beg_table.style = 'Table Grid'; beg_table.autofit = False
        beg_table.columns[0].width = table_col1_width; beg_table.columns[1].width = table_col2_width
        set_cell_style(beg_table.cell(0,0), "Cash in Bank-beg")
        set_cell_style(beg_table.cell(0,1), format_amount_docx(self.variables['cash_bank_beg'].get()), align='right')
        set_cell_style(beg_table.cell(1,0), "Cash on Hand-beg")
        set_cell_style(beg_table.cell(1,1), format_amount_docx(self.variables['cash_hand_beg'].get()), align='right')

        # Cash Inflows
        p_in = doc.add_paragraph(); run_in = p_in.add_run("\nCash Inflows")
        run_in.font.name = 'Helvetica'; run_in.bold = True; run_in.font.size = Pt(10)
        inflow_items = [
            ("Monthly Dues Collected", self.variables['monthly_dues']), ("Certifications Issued", self.variables['certifications']),
            ("Membership Fee", self.variables['membership_fee']), ("Vehicle Stickers", self.variables['vehicle_stickers']),
            ("Rentals", self.variables['rentals']), ("Solicitations/Donations", self.variables['solicitations']),
            ("Interest Income on Bank Deposits", self.variables['interest_income']), ("Livelihood Management Fee", self.variables['livelihood_fee']),
            ("Others (Inflow)", self.variables['inflows_others']), ("Total Cash Receipts", self.variables['total_receipts'])
        ]
        in_table = doc.add_table(rows=len(inflow_items), cols=2)
        in_table.style = 'Table Grid'; in_table.autofit = False
        in_table.columns[0].width = table_col1_width; in_table.columns[1].width = table_col2_width
        for i, (label, var) in enumerate(inflow_items):
            is_total = (label == "Total Cash Receipts")
            set_cell_style(in_table.cell(i,0), label, bold=is_total)
            set_cell_style(in_table.cell(i,1), format_amount_docx(var.get()), align='right', bold=is_total)

        # Cash Outflows
        p_out = doc.add_paragraph(); run_out = p_out.add_run("\nLess: Cash Outflows")
        run_out.font.name = 'Helvetica'; run_out.bold = True; run_out.font.size = Pt(10)
        outflow_items = [
            ("Snacks/Meals for Visitors", self.variables['snacks_meals']), ("Transportation Expenses", self.variables['transportation']),
            ("Office Supplies Expense", self.variables['office_supplies']), ("Printing and Photocopy", self.variables['printing']),
            ("Labor", self.variables['labor']), ("Billboard Expense", self.variables['billboard']),
            ("Clearing/Cleaning Charges", self.variables['cleaning']), ("Miscellaneous Expenses", self.variables['misc_expenses']),
            ("Federation Fee", self.variables['federation_fee']), ("HOA-BOD Uniforms", self.variables['uniforms']),
            ("BOD Meeting", self.variables['bod_mtg']), ("General Assembly", self.variables['general_assembly']),
            ("Cash Deposit to Bank", self.variables['cash_deposit']), ("Withholding Tax on Bank Deposit", self.variables['withholding_tax']),
            ("Refund", self.variables['refund']), ("Others (Outflow)", self.variables['outflows_others']),
            ("Cash Outflows/Disbursements", self.variables['cash_outflows'])
        ]
        out_table = doc.add_table(rows=len(outflow_items), cols=2)
        out_table.style = 'Table Grid'; out_table.autofit = False
        out_table.columns[0].width = table_col1_width; out_table.columns[1].width = table_col2_width
        for i, (label, var) in enumerate(outflow_items):
            is_total = (label == "Cash Outflows/Disbursements")
            set_cell_style(out_table.cell(i,0), label, bold=is_total)
            set_cell_style(out_table.cell(i,1), format_amount_docx(var.get()), align='right', bold=is_total)

        # Ending Cash Balance
        p_end = doc.add_paragraph(); run_end = p_end.add_run("\nEnding Cash Balance")
        run_end.font.name = 'Helvetica'; run_end.bold = True; run_end.font.size = Pt(10)
        end_table = doc.add_table(rows=1, cols=2)
        end_table.style = 'Table Grid'; end_table.autofit = False
        end_table.columns[0].width = table_col1_width; end_table.columns[1].width = table_col2_width
        set_cell_style(end_table.cell(0,0), "Ending Cash Balance", bold=True)
        set_cell_style(end_table.cell(0,1), format_amount_docx(self.variables['ending_cash'].get()), align='right', bold=True)

        # Breakdown of Cash
        p_brk = doc.add_paragraph(); run_brk = p_brk.add_run("\nBreakdown of Cash")
        run_brk.font.name = 'Helvetica'; run_brk.bold = True; run_brk.font.size = Pt(10)
        brk_table = doc.add_table(rows=2, cols=2)
        brk_table.style = 'Table Grid'; brk_table.autofit = False
        brk_table.columns[0].width = table_col1_width; brk_table.columns[1].width = table_col2_width
        set_cell_style(brk_table.cell(0,0), "Cash in Bank")
        set_cell_style(brk_table.cell(0,1), format_amount_docx(self.variables['ending_cash_bank'].get()), align='right')
        set_cell_style(brk_table.cell(1,0), "Cash on Hand")
        set_cell_style(brk_table.cell(1,1), format_amount_docx(self.variables['ending_cash_hand'].get()), align='right')
        
        doc.add_paragraph() # Spacer

        # Signatories
        prepared_name = self.prepared_by_var.get() or "_______________________"
        noted_name_1 = self.noted_by_var_1.get() or "_______________________"
        noted_name_2 = self.noted_by_var_2.get() or "_______________________"
        checked_name = self.checked_by_var.get() or "_______________________"

        # Prepared by / Checked by table
        prep_check_table = doc.add_table(rows=3, cols=2)
        prep_check_table.autofit = False
        # Set column widths for a 2-column layout within the page margins
        col_width_sign = Inches(3.75) # (7.5 / 2)
        prep_check_table.columns[0].width = col_width_sign
        prep_check_table.columns[1].width = col_width_sign
        prep_check_table.alignment = WD_TABLE_ALIGNMENT.CENTER


        set_cell_style(prep_check_table.cell(0,0), "Prepared by:", align='center')
        set_cell_style(prep_check_table.cell(0,1), "Checked by:", align='center')
        set_cell_style(prep_check_table.cell(1,0), prepared_name, align='center', bold=True) # Names can be bold
        set_cell_style(prep_check_table.cell(1,1), checked_name, align='center', bold=True)
        set_cell_style(prep_check_table.cell(2,0), "HOA Treasurer", align='center')
        set_cell_style(prep_check_table.cell(2,1), "HOA Auditor", align='center')
        
        # "Noted by" title - centered paragraph
        p_noted_title = doc.add_paragraph()
        p_noted_title.add_run("Noted by:").font.name = 'Helvetica'; p_noted_title.runs[0].font.size = Pt(8)
        p_noted_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_noted_title.paragraph_format.space_before = Pt(12)
        p_noted_title.paragraph_format.space_after = Pt(6)

        # Noted by table
        noted_table = doc.add_table(rows=2, cols=2)
        noted_table.autofit = False
        noted_table.columns[0].width = col_width_sign
        noted_table.columns[1].width = col_width_sign
        noted_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        set_cell_style(noted_table.cell(0,0), noted_name_1, align='center', bold=True)
        set_cell_style(noted_table.cell(0,1), noted_name_2, align='center', bold=True)
        set_cell_style(noted_table.cell(1,0), "HOA President", align='center')
        set_cell_style(noted_table.cell(1,1), "CHUDD HCD-CORDS", align='center')
        
        return doc

    def _create_docx_at_path(self, filename):
        """Internal method to generate DOCX content and save to the given filename."""
        try:
            doc = self._build_docx_document()
            doc.save(filename)
            return True 
        except Exception as e:
            logging.exception(f"Error creating DOCX content for {filename}")
            return False

    def save_to_docx(self):
        """Prompts user for save location and saves data to a Word document.
           Returns a dictionary with 'status' and 'message'."""
        try:
            default_filename = f"cash_flow_statement_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            filename = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx")],
                initialfile=default_filename,
                title="Save Word Document As"
            )
            if not filename:
                return {"status": "cancelled", "message": "Save to Word cancelled by user."}

            if self._create_docx_at_path(filename):
                logging.info(f"Word document successfully saved to {filename}")
                return {"status": "success", "message": f"Word document saved to {filename}", "filename": filename}
            else:
                return {"status": "error", "message": f"Failed to create Word document at {filename}.\nCheck logs and ensure python-docx is installed."}
        except Exception as e:
            logging.exception("Error during Word save process")
            return {"status": "error", "message": f"An unexpected error occurred during Word save: {str(e)}"}

    def generate_temp_docx(self):
        """Generates a DOCX in a temporary location.
           Returns a dictionary with 'status', 'filename', and 'message'."""
        try:
            fd, temp_filename = tempfile.mkstemp(suffix=".docx", prefix="cash_flow_")
            os.close(fd)

            if self._create_docx_at_path(temp_filename):
                logging.info(f"Temporary DOCX for email generated at {temp_filename}")
                return {"status": "success", "filename": temp_filename}
            else:
                if os.path.exists(temp_filename):
                    try: os.remove(temp_filename)
                    except Exception as e_del: logging.error(f"Could not remove temp DOCX {temp_filename} after generation failure: {e_del}")
                return {"status": "error", "message": "Failed to generate temporary DOCX."}
        except Exception as e:
            logging.exception("Error generating temporary DOCX")
            return {"status": "error", "message": f"An unexpected error occurred while generating temporary DOCX: {str(e)}"}

    # load_from_docx, load_from_pdf, load_from_documentpdf remain the same
    def load_from_docx(self, filename):
        """Load data from a Word document into the application variables."""
        try:
            logging.info(f"Loading DOCX: {filename}")
            doc = Document(filename)
            current_directory = os.getcwd()
            for section in doc.sections:
                header = section.header
                for rel in header.part.rels.values():
                    if "image" in rel.reltype:
                        image_data = rel.target_part.blob
                        image_ext = rel.target_part.content_type.split("/")[-1]
                        image_filename = os.path.join(current_directory, f"Logo_{random.random()}.{image_ext}")
                        with open(image_filename, "wb") as f:
                            f.write(image_data)
                            print(f"Saved header image: {image_filename}")
                            self.variables['logo_path_var'].set(image_filename)
            
            self.address_var.set("")
            address_found = False
            for section in doc.sections:
                headers = [section.header, section.first_page_header, section.even_page_header]
                for header in headers:
                    if not header: continue
                    for table in header.tables:
                        if len(table.rows) > 0 and len(table.rows[0].cells) >= 2:
                            cell_content = table.rows[0].cells[1].text.strip()
                            lines = [line.strip() for line in cell_content.splitlines() if line.strip()]
                            unwanted_patterns = [r"CASH FLOW STATEMENT", r"For the Month of \w+\s+\d{1,2},\s+\d{4}"]
                            address_lines = [line for line in lines if not any(re.search(pattern, line, re.IGNORECASE) for pattern in unwanted_patterns)]
                            if address_lines:
                                self.address_var.set(address_lines[0])
                                address_found = True; break
                    if address_found: break
                if address_found: break
            if not address_found: logging.warning("No address found in any header table.")

            label_to_var = {
                "Cash in Bank-beg": self.variables['cash_bank_beg'],
                "Cash on Hand-beg": self.variables['cash_hand_beg'],
                "Monthly Dues Collected": self.variables['monthly_dues'],
                "Certifications Issued": self.variables['certifications'],
                "Membership Fee": self.variables['membership_fee'],
                "Vehicle Stickers": self.variables['vehicle_stickers'],
                "Rentals": self.variables['rentals'],
                "Solicitations/Donations": self.variables['solicitations'],
                "Interest Income on Bank Deposits": self.variables['interest_income'],
                "Livelihood Management Fee": self.variables['livelihood_fee'],
                "Others (Inflow)": self.variables['inflows_others'],
                "Total Cash Receipts": self.variables['total_receipts'],
                "Cash Outflows/Disbursements": self.variables['cash_outflows'],
                "Snacks/Meals for Visitors": self.variables['snacks_meals'],
                "Transportation Expenses": self.variables['transportation'],
                "Office Supplies Expense": self.variables['office_supplies'],
                "Printing and Photocopy": self.variables['printing'],
                "Labor": self.variables['labor'],
                "Billboard Expense": self.variables['billboard'],
                "Clearing/Cleaning Charges": self.variables['cleaning'],
                "Miscellaneous Expenses": self.variables['misc_expenses'],
                "Federation Fee": self.variables['federation_fee'],
                "HOA-BOD Uniforms": self.variables['uniforms'],
                "BOD Meeting": self.variables['bod_mtg'],
                "General Assembly": self.variables['general_assembly'],
                "Cash Deposit to Bank": self.variables['cash_deposit'],
                "Withholding Tax on Bank Deposit": self.variables['withholding_tax'],
                "Refund": self.variables['refund'],
                "Others (Outflow)": self.variables['outflows_others']
            }

            # Extract table data
            for table in doc.tables:
                for row in table.rows:
                    if len(row.cells) >= 2:
                        label = row.cells[0].text.strip()
                        value = row.cells[1].text.strip()
                        # logging.debug(f"Extracted Label: '{label}', Value: '{value}'") # Debugging
                        if label.lower().startswith("for the month of"): # Adjusted keyword
                            try:
                                lines = value.splitlines()
                                lines = [line.strip() for line in lines if line.strip()]
                                # Extract date string after "For the Month of "
                                date_str_part = label.split("For the Month of", 1)[1].strip()
                                date_obj = datetime.datetime.strptime(date_str_part, "%B %d, %Y")
                                self.date_var.set(date_obj.strftime("%m/%d/%Y"))
                                logging.info(f"Loaded date: {self.date_var.get()}")
                            except (IndexError, ValueError, TypeError) as e:
                                logging.warning(f"Could not parse date from header: '{label}', Error: {e}")
                                # Try parsing just the value if label didn't work
                                try:
                                    date_obj = datetime.datetime.strptime(value, "%B %d, %Y")
                                    self.date_var.set(date_obj.strftime("%m/%d/%Y"))
                                    logging.info(f"Loaded date from value: {self.date_var.get()}")
                                except: pass # Ignore if value also fails
                        if label.lower().startswith("prepared by:"): # Adjusted keyword
                            treasurer = table.rows[1].cells[0].text.strip()
                            auditor = table.rows[1].cells[1].text.strip()
                            self.prepared_by_var.set(treasurer)
                            self.checked_by_var.set(auditor)

                        if label.lower().startswith("hoa president"): # Adjusted keyword
                            president = table.rows[0].cells[0].text.strip()
                            Chords = table.rows[0].cells[1].text.strip()
                            self.noted_by_var_1.set(president)
                            self.noted_by_var_2.set(Chords)
                        if label in label_to_var:
                            parsed_value = self.parse_amount(value)
                            label_to_var[label].set(parsed_value)
                            # logging.debug(f"Set {label} to {parsed_value}") # Debugging
                
            messagebox.showinfo("Success", "DOCX data loaded successfully")
            return True

        except Exception as e:
            logging.exception(f"Error loading Word document: {filename}") # Log full traceback
            messagebox.showerror("Error", f"Error loading Word document:\n{str(e)}")
            return False

    def load_from_pdf(self, filename):
        """Load data from a PDF document into the application variables."""
        try:
            logging.info(f"Attempting to load PDF: {filename}")
            self.address_var.set(""); self.prepared_by_var.set(""); self.checked_by_var.set("")
            self.noted_by_var_1.set(""); self.noted_by_var_2.set("")

            label_to_var = {
                "Cash in Bank-beg": self.variables['cash_bank_beg'], "Cash on Hand-beg": self.variables['cash_hand_beg'],
                "Monthly Dues Collected": self.variables['monthly_dues'], "Certifications Issued": self.variables['certifications'],
                "Membership Fee": self.variables['membership_fee'], "Vehicle Stickers": self.variables['vehicle_stickers'],
                "Rentals": self.variables['rentals'], "Solicitations/Donations": self.variables['solicitations'],
                "Interest Income on Bank Deposits": self.variables['interest_income'], "Livelihood Management Fee": self.variables['livelihood_fee'],
                "Others (Inflow)": self.variables['inflows_others'], "Total Cash Receipts": self.variables['total_receipts'],
                "Cash Outflows/Disbursements": self.variables['cash_outflows'], "Snacks/Meals for Visitors": self.variables['snacks_meals'],
                "Transportation Expenses": self.variables['transportation'], "Office Supplies Expense": self.variables['office_supplies'],
                "Printing and Photocopy": self.variables['printing'], "Labor": self.variables['labor'],
                "Billboard Expense": self.variables['billboard'], "Clearing/Cleaning Charges": self.variables['cleaning'],
                "Miscellaneous Expenses": self.variables['misc_expenses'], "Federation Fee": self.variables['federation_fee'],
                "HOA-BOD Uniforms": self.variables['uniforms'], "BOD Meeting": self.variables['bod_mtg'],
                "General Assembly": self.variables['general_assembly'], "Cash Deposit to Bank": self.variables['cash_deposit'],
                "Withholding Tax on Bank Deposit": self.variables['withholding_tax'], "Refund": self.variables['refund'],
                "Others (Outflow)": self.variables['outflows_others']
            }
            address_found = False; date_found = False
            
            with pdfplumber.open(filename) as pdf:
                first_page = pdf.pages[0]
                page_text = first_page.extract_text()
                if page_text: # Try extracting address from raw text first
                    lines = [line.strip() for line in page_text.splitlines() if line.strip()]
                    for i, line in enumerate(lines):
                        if re.search(r"CASH FLOW STATEMENT", line, re.IGNORECASE):
                            if i > 0 and not re.search(r"For the Month of", lines[i-1], re.IGNORECASE) and not lines[i-1].startswith("[Logo"):
                                self.address_var.set(lines[i-1])
                                address_found = True; break
                
                if not address_found: # Fallback to table extraction for address
                    tables = first_page.extract_tables()
                    for table in tables:
                        if table and len(table) > 0 and len(table[0]) >= 2:
                            cell_content = str(table[0][1]).strip() # Assuming address is in the second column of a header-like table
                            lines = [line.strip() for line in cell_content.splitlines() if line.strip()]
                            unwanted_patterns = [r"CASH FLOW STATEMENT", r"For the Month of \w+\s+\d{1,2},\s+\d{4}"]
                            address_lines = [line for line in lines if not any(re.search(pattern, line, re.IGNORECASE) for pattern in unwanted_patterns)]
                            if address_lines:
                                self.address_var.set(address_lines[0]); address_found = True; break
                if not address_found: self.address_var.set("Default Address - Change Me")

                # Extract cash flow data
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table_data in tables:
                        if not table_data: continue
                        for row in table_data:
                            if not row or len(row) < 2 or not row[0]: continue
                            label = str(row[0]).replace('\n', ' ').strip()
                            value = str(row[1]).replace('\n', ' ').strip() if len(row) > 1 and row[1] else ""
                            if label in label_to_var:
                                label_to_var[label].set(self.parse_amount(value))
                
                # Extract signatories using Camelot (more robust for complex table structures if present)
                try:
                    camelot_tables = camelot.read_pdf(filename, flavor="stream", pages="all", table_areas=["0,200,612,0"], row_tol=15, column_tol=15, split_text=True, strip_text='\n') # Bottom part of page
                    for table in camelot_tables:
                        df = table.df
                        for r_idx in range(len(df)):
                            for c_idx in range(len(df.columns)):
                                cell_text = str(df.iloc[r_idx, c_idx]).strip()
                                if "Prepared by:" in cell_text:
                                    self.prepared_by_var.set(cell_text.replace("Prepared by:", "").strip())
                                elif "Checked by:" in cell_text:
                                    self.checked_by_var.set(cell_text.replace("Checked by:", "").strip())
                                elif "HOA President" in cell_text and "Noted by:" not in cell_text : # Avoid title
                                    self.noted_by_var_1.set(cell_text.replace("HOA President", "").strip())
                                elif "CHUDD HCD-CORDS" in cell_text and "Noted by:" not in cell_text :
                                    self.noted_by_var_2.set(cell_text.replace("CHUDD HCD-CORDS", "").strip())
                except Exception as e_camelot:
                    logging.warning(f"Camelot extraction for signatories failed or found no tables: {e_camelot}")
                    # Fallback to simple text search if Camelot fails or specific structure not found
                    full_text_for_names = "".join([p.extract_text() or "" for p in pdf.pages])
                    # Simplified regex, might need adjustment
                    prepared_match = re.search(r"Prepared by:\s*([^\n]+)\s*HOA Treasurer", full_text_for_names, re.IGNORECASE)
                    if prepared_match: self.prepared_by_var.set(prepared_match.group(1).strip())
                    checked_match = re.search(r"Checked by:\s*([^\n]+)\s*HOA Auditor", full_text_for_names, re.IGNORECASE)
                    if checked_match: self.checked_by_var.set(checked_match.group(1).strip())
                    noted1_match = re.search(r"([^\n]+)\s*HOA President", full_text_for_names, re.IGNORECASE) # Be careful this doesn't match "Noted by:" itself
                    if noted1_match and "noted by" not in noted1_match.group(1).lower(): self.noted_by_var_1.set(noted1_match.group(1).strip())
                    noted2_match = re.search(r"([^\n]+)\s*CHUDD HCD-CORDS", full_text_for_names, re.IGNORECASE)
                    if noted2_match and "noted by" not in noted2_match.group(1).lower(): self.noted_by_var_2.set(noted2_match.group(1).strip())


                # Extract date
                full_text_for_date = "".join([p.extract_text() or "" for p in pdf.pages])
                date_match = re.search(r"For the Month of\s+(\w+\s+\d{1,2},\s+\d{4})", full_text_for_date, re.IGNORECASE)
                if date_match:
                    try:
                        date_obj = datetime.datetime.strptime(date_match.group(1).strip(), "%B %d, %Y")
                        self.date_var.set(date_obj.strftime("%m/%d/%Y"))
                        date_found = True
                    except ValueError: pass
                if not date_found: logging.warning("Date not found in PDF.")

                # Extract logo
                doc_fitz = fitz.open(filename)
                for page_num in range(len(doc_fitz)):
                    page_fitz = doc_fitz[page_num]
                    images = page_fitz.get_images(full=True)
                    if images: # Take the first image as logo (heuristic)
                        img_info = images[0]
                        xref = img_info[0]
                        base_image = doc_fitz.extract_image(xref)
                        image_bytes = base_image["image"]
                        image_ext = base_image["ext"]
                        image_filename = os.path.join(os.getcwd(), f"Logo_extracted_{random.random()}.{image_ext}")
                        with open(image_filename, "wb") as f: f.write(image_bytes)
                        self.variables['logo_path_var'].set(image_filename)
                        break # Assume first image found is the logo
            
            messagebox.showinfo("Success", "PDF data loaded successfully!")
            if hasattr(self.variables.get('calculator'), 'calculate_totals'):
                self.variables['calculator'].calculate_totals()
            return True
        except Exception as e:
            logging.exception(f"Error loading PDF: {filename}")
            messagebox.showerror("Error", f"Error loading PDF:\n{str(e)}")
            return False
        
    def load_from_documentpdf(self):
        try:
            filename = filedialog.askopenfilename(
                filetypes=[("Documents", "*.docx *.pdf"), ("Word Documents", "*.docx"), ("PDF Files", "*.pdf"), ("All Files", "*.*")],
                title="Select a Document (DOCX or PDF)"
            )
            if not filename: return

            if filename.lower().endswith('.pdf'): self.load_from_pdf(filename)
            elif filename.lower().endswith('.docx'): self.load_from_docx(filename)
            else: messagebox.showerror("Error", "Unsupported file format. Please select a PDF or DOCX file.")
        except Exception as e:
             logging.exception("Error loading document")
             messagebox.showerror("Error", f"Error loading document: {str(e)}")

# --- END OF FILE file_handler.py ---