# --- START OF FILE file_handler.py ---

import os
import re
import sys # Keep sys if used elsewhere, not strictly needed for changes here
import datetime
import logging # Use logging
from decimal import Decimal
from tkinter import filedialog, messagebox

# PDF Imports
import pdfplumber
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image # Added Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter # Keep for reference, but using custom size
from reportlab.lib.units import inch # Import inch for easier sizing
import camelot

# Word Imports
from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.shared import Inches, Pt
from docx.oxml.ns import qn # Import qn for column settings
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT  # Add import at the top of the file
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn  # For column settings

# Configure logging if not already done elsewhere
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class FileHandler:
    # --- MODIFIED __init__ ---
    def __init__(self, variables, title_var, date_var, logo_path_var, address_var, prepared_by_var, noted_by_var_1, noted_by_var_2, checked_by_var):
        self.variables = variables
        self.title_var = title_var
        self.date_var = date_var
        self.logo_path_var = logo_path_var # Store logo path variable
        self.address_var = address_var   # Store address variable
        self.prepared_by_var = prepared_by_var
        self.noted_by_var_1 = noted_by_var_1
        self.noted_by_var_2 = noted_by_var_2
        self.checked_by_var = checked_by_var
    # --- END MODIFIED __init__ ---

    def parse_amount(self, text):
        """Parse text to extract numerical amount, removing non-numeric characters except decimal."""
        if not text or text.strip() == "":
            return ""
        try:
            # Remove commas before attempting conversion
            cleaned_text = re.sub(r'[^\d.]', '', text.replace(',', ''))
            return str(Decimal(cleaned_text))
        except Exception as e:
            logging.warning(f"Could not parse amount '{text}': {e}")
            return text # Return original text if parsing fails

    def format_date_for_display(self, date_str):
        """Convert mm/dd/yyyy to MMMM dd, yyyy for display."""
        try:
            date_obj = datetime.datetime.strptime(date_str, "%m/%d/%Y")
            return date_obj.strftime("%B %d, %Y")
        except ValueError:
            logging.warning(f"Invalid date format for display: {date_str}")
            return date_str # Return original if format is wrong

    def load_from_docx(self, filename):
        """Load data from a Word document into the application variables."""
        # --- NOTE: This does NOT load logo/address from the docx ---
        # --- It loads cash flow data and signatory names only ---
        try:
            logging.info(f"Loading DOCX: {filename}")
            doc = Document(filename)

        # Initialize address_var
            self.address_var.set("")

        # Step 1: Extract address from header table
            address_found = False
            for section in doc.sections:
                headers = [
                section.header,
                section.first_page_header,
                section.even_page_header
            ]
                for header in headers:
                    for table in header.tables:
                        if len(table.rows) > 0 and len(table.rows[0].cells) >= 2:
                            cell_content = table.rows[0].cells[1].text.strip()
                            logging.debug(f"Header table cell content: {cell_content}")
                            lines = cell_content.splitlines()
                            lines = [line.strip() for line in lines if line.strip()]

                        # Filter out unwanted lines
                        # Use a static date pattern since self.date_var may not be set yet
                            unwanted_patterns = [
                            r"CASH FLOW STATEMENT",
                            r"For the Month of \w+\s+\d{1,2},\s+\d{4}"  # Matches "For the Month of May 05, 2025"
                        ]
                            address_lines = [
                            line for line in lines
                            if not any(re.search(pattern, line, re.IGNORECASE) for pattern in unwanted_patterns)
                        ]
                            if address_lines:
                                self.address_var.set(address_lines[0])  # e.g., "Default Address - Change Me"
                                address_found = True
                                logging.info(f"Extracted address from header: {self.address_var.get()}")
                                break  # Stop after finding the first valid address
                    if address_found:
                        break
                if address_found:
                    break

            if not address_found:
                logging.warning("No address found in any header table.")
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
            # Initialize variables
            self.address_var.set("")
            self.prepared_by_var.set("")
            self.checked_by_var.set("")
            self.noted_by_var_1.set("")
            self.noted_by_var_2.set("")

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

            address_found = False
            date_found = False
            names_found = {'prepared': False, 'checked': False, 'noted1': False, 'noted2': False}

            with pdfplumber.open(filename) as pdf:
                # Step 1: Extract address from first page text (primary method)
                first_page = pdf.pages[0]
                page_text = first_page.extract_text()
                if page_text:
                    lines = page_text.splitlines()
                    lines = [line.strip() for line in lines if line.strip()]
                    logging.debug(f"First page text lines: {lines}")
                    for i, line in enumerate(lines):
                        if re.search(r"CASH FLOW STATEMENT", line, re.IGNORECASE):
                            # Look for the address in the line(s) before "CASH FLOW STATEMENT"
                            for j in range(i - 1, -1, -1):
                                if lines[j] and not re.search(r"^\[Logo", lines[j]) and not re.search(r"For the Month of", lines[j], re.IGNORECASE):
                                    self.address_var.set(lines[j])
                                    address_found = True
                                    logging.info(f"Extracted address from text: {self.address_var.get()}")
                                    break
                            break

                # Step 2: Fallback to table extraction for address
                if not address_found:
                    tables = first_page.extract_tables()
                    logging.debug(f"First page tables: {tables}")
                    for table in tables:
                        if not table or len(table) == 0 or len(table[0]) < 2:
                            continue
                        cell_content = str(table[0][1]).strip()
                        logging.debug(f"Table cell content: {cell_content}")
                        lines = cell_content.splitlines()
                        lines = [line.strip() for line in lines if line.strip()]
                        unwanted_patterns = [
                            r"CASH FLOW STATEMENT",
                            r"For the Month of \w+\s+\d{1,2},\s+\d{4}"
                        ]
                        address_lines = [
                            line for line in lines
                            if not any(re.search(pattern, line, re.IGNORECASE) for pattern in unwanted_patterns)
                        ]
                        if address_lines:
                            self.address_var.set(address_lines[0])
                            address_found = True
                            logging.info(f"Extracted address from table: {self.address_var.get()}")
                            break

                if not address_found:
                    logging.warning("No address found in PDF. Setting default.")
                    self.address_var.set("Default Address - Change Me")

                # Step 3: Extract cash flow data from tables
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table:
                            continue
                        for row in table:
                            if not row or len(row) < 2 or not row[0]:
                                continue
                            label = str(row[0]).replace('\n', ' ').strip() if row[0] else ""
                            value = str(row[1]).replace('\n', ' ').strip() if row[1] else ""
                            logging.debug(f"Table row: Label='{label}', Value='{value}'")
                            if label in label_to_var:
                                try:
                                    parsed_value = self.parse_amount(value)
                                    label_to_var[label].set(parsed_value)
                                    logging.debug(f"Set {label} to {parsed_value}")
                                except Exception as e:
                                    logging.warning(f"Error setting {label}: {e}")
                #signatories
                try:
    # Assume 8.5x11 page (612x792 points), footer in bottom 200 points
                    tables = camelot.read_pdf(
        filename,
        flavor="stream",
        pages="all",
        table_areas=["0,300,612,0"],  # Adjust y1 to include all footer text
        split_text=True,  # Handle multi-line cells
        strip_text='\n',  # Remove newlines
        row_tol=15,  # Loose row grouping
        column_tol=15  # Loose column grouping
    )
                    for table in tables:
                        print("Footer table found:")
                        print(table.df)
        # Check for "Prepared by:"
                        print(table.df.iloc[3, 3].strip())
                        if table.df.iloc[3, 1].strip().startswith("Prepared by:"):
                            prepared_by = table.df.iloc[3, 1].strip() if len(table.df) > 1 else ""
                            newPrepared = prepared_by.replace("Prepared by:", "").strip()
                            self.prepared_by_var.set(newPrepared)
                        if table.df.iloc[3, 3].strip().startswith("Checked by:"):
                            checked_by = table.df.iloc[3, 3].strip() if len(table.df) > 1 else ""
                            newChecked = checked_by.replace("Checked by:", "").strip()
                            self.checked_by_var.set(newChecked)
                        if table.df.iloc[6, 1].strip().endswith("HOA President"):
                            noted_by_1 = table.df.iloc[6, 1].strip() if len(table.df) > 1 else ""
                            newNoted1 = noted_by_1.replace("HOA President", "").strip()
                            self.noted_by_var_1.set(newNoted1)
                        if table.df.iloc[6, 3].strip().endswith("CHUDD HCD-CORDS"):
                            noted_by_2 = table.df.iloc[6, 3].strip() if len(table.df) > 1 else ""
                            newNoted2 = noted_by_2.replace("CHUDD HCD-CORDS", "").strip()
                            self.noted_by_var_2.set(newNoted2)
                        if not tables:
                            print("No footer tables detected with Camelot.")
                except Exception as e:
                    print(f"Camelot error: {str(e)}")
                    if not any(names_found.values()):
                        logging.debug("Signatory table not found. Falling back to text extraction.")
                    
                full_text = ""
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        full_text += page_text + "\n"
                    #
                    #print(full_text)  # Debugging: Print full text for analysis
                
                # Step 6: Extract date
                date_match = re.search(r"For the Month of\s+(\w+\s+\d{1,2},\s+\d{4})", full_text, re.IGNORECASE)
                if date_match:
                    try:
                        date_str = date_match.group(1).strip()
                        date_obj = datetime.datetime.strptime(date_str, "%B %d, %Y")
                        self.date_var.set(date_obj.strftime("%m/%d/%Y"))
                        logging.info(f"Extracted date: {self.date_var.get()}")
                        date_found = True
                    except ValueError as e:
                        logging.warning(f"Could not parse date: {date_str}, Error: {e}")

                if not date_found:
                    logging.warning("Date not found in PDF.")
                if not any(names_found.values()):
                    logging.warning("No signatory names found in PDF.")

            messagebox.showinfo("Success", "PDF data loaded successfully!")
            if hasattr(self.variables.get('calculator'), 'calculate_totals'):
                self.variables['calculator'].calculate_totals()
            return True

        except Exception as e:
            logging.exception(f"Error loading PDF: {filename}")
            messagebox.showerror("Error", f"Error loading PDF:\n{str(e)}")
            return False
        
    # --- MODIFIED save_to_docx ---
    def save_to_docx(self):
        """Save data to a Word document with logo, address, and two Noted by fields in the footer."""
        try:
            def format_amount(value): # Inner function for formatting
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

            # --- Page Setup (Folio 8.5 x 13) ---
            section = doc.sections[0]
            section.page_width = Inches(8.5)
            section.page_height = Inches(13)
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.7) # Slightly more space for footer if needed
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

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

            # --- Body Content (Tables) ---
            # Helper to set cell font and alignment
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
                # Optional: Adjust cell margins?
                # cell.margin_left = Inches(0.05)
                # cell.margin_right = Inches(0.05)


            # Add a spacer paragraph after header content (in main body)

            # Beginning Cash Balances
            p = doc.add_paragraph()
            run = p.add_run("Beginning Cash Balances")
            run.font.name = 'Helvetica'
            run.bold = True
            run.font.size = Pt(10)
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            table.autofit = False # Use manual widths
            table.columns[0].width = Inches(6.0) # Adjust widths based on 7.5" content area
            table.columns[1].width = Inches(1.5)
            set_cell_style(table.cell(0, 0), "Cash in Bank-beg")
            set_cell_style(table.cell(0, 1), format_amount(self.variables['cash_bank_beg'].get()), align='right')
            set_cell_style(table.cell(1, 0), "Cash on Hand-beg")
            set_cell_style(table.cell(1, 1), format_amount(self.variables['cash_hand_beg'].get()), align='right')
            

            # Cash Inflows
            p = doc.add_paragraph()
            run = p.add_run("\nCash Inflows")
            run.font.name = 'Helvetica'
            run.bold = True
            run.font.size = Pt(10)
            inflow_items = [
                ("Monthly Dues Collected", self.variables['monthly_dues']),
                ("Certifications Issued", self.variables['certifications']),
                ("Membership Fee", self.variables['membership_fee']),
                ("Vehicle Stickers", self.variables['vehicle_stickers']),
                ("Rentals", self.variables['rentals']),
                ("Solicitations/Donations", self.variables['solicitations']),
                ("Interest Income on Bank Deposits", self.variables['interest_income']),
                ("Livelihood Management Fee", self.variables['livelihood_fee']),
                ("Others (Inflow)", self.variables['inflows_others']),
                ("Total Cash Receipts", self.variables['total_receipts'])
            ]
            table = doc.add_table(rows=len(inflow_items), cols=2)
            table.style = 'Table Grid'
            table.autofit = False
            table.columns[0].width = Inches(6.0)
            table.columns[1].width = Inches(1.5)
            for i, (label, var) in enumerate(inflow_items):
                 is_total = (i == len(inflow_items) - 1)
                 set_cell_style(table.cell(i, 0), label, bold=is_total)
                 set_cell_style(table.cell(i, 1), format_amount(var.get()), align='right', bold=is_total)
            

            # Cash Outflows
            p = doc.add_paragraph()
            run = p.add_run("\nLess: Cash Outflows")
            run.font.name = 'Helvetica'
            run.bold = True
            run.font.size = Pt(10)
            outflow_items = [
                ("Snacks/Meals for Visitors", self.variables['snacks_meals']),
                ("Transportation Expenses", self.variables['transportation']),
                ("Office Supplies Expense", self.variables['office_supplies']),
                ("Printing and Photocopy", self.variables['printing']),
                ("Labor", self.variables['labor']),
                ("Billboard Expense", self.variables['billboard']),
                ("Clearing/Cleaning Charges", self.variables['cleaning']),
                ("Miscellaneous Expenses", self.variables['misc_expenses']),
                ("Federation Fee", self.variables['federation_fee']),
                ("HOA-BOD Uniforms", self.variables['uniforms']),
                ("BOD Meeting", self.variables['bod_mtg']),
                ("General Assembly", self.variables['general_assembly']),
                ("Cash Deposit to Bank", self.variables['cash_deposit']),
                ("Withholding Tax on Bank Deposit", self.variables['withholding_tax']),
                ("Refund", self.variables['refund']),
                ("Others (Outflow)", self.variables['outflows_others']),
                ("Cash Outflows/Disbursements", self.variables['cash_outflows'])
            ]
            table = doc.add_table(rows=len(outflow_items), cols=2)
            table.style = 'Table Grid'
            table.autofit = False
            table.columns[0].width = Inches(6.0)
            table.columns[1].width = Inches(1.5)
            for i, (label, var) in enumerate(outflow_items):
                 is_total = (i == len(outflow_items) - 1)
                 set_cell_style(table.cell(i, 0), label, bold=is_total)
                 set_cell_style(table.cell(i, 1), format_amount(var.get()), align='right', bold=is_total)

            # Ending Cash Balance
            p = doc.add_paragraph()
            run = p.add_run("\nEnding Cash Balance")
            run.font.name = 'Helvetica'
            run.bold = True
            run.font.size = Pt(10)
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            table.autofit = False
            table.columns[0].width = Inches(6.0)
            table.columns[1].width = Inches(1.5)
            set_cell_style(table.cell(0, 0), "Ending Cash Balance", bold=True)
            set_cell_style(table.cell(0, 1), format_amount(self.variables['ending_cash'].get()), align='right', bold=True)

            # Breakdown of Cash
            p = doc.add_paragraph()
            run = p.add_run("\nBreakdown of Cash")
            run.font.name = 'Helvetica'
            run.bold = True
            run.font.size = Pt(10)
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            table.autofit = False
            table.columns[0].width = Inches(6.0)
            table.columns[1].width = Inches(1.5)
            set_cell_style(table.cell(0, 0), "Cash in Bank")
            set_cell_style(table.cell(0, 1), format_amount(self.variables['ending_cash_bank'].get()), align='right')
            set_cell_style(table.cell(1, 0), "Cash on Hand")
            set_cell_style(table.cell(1, 1), format_amount(self.variables['ending_cash_hand'].get()), align='right')
            doc.add_paragraph() # Spacer before footer section

            # --- Footer / Signatories ---
             # --- Footer / Signatories ---
            # (Using section breaks and column changes as before)
            prepared_name = self.prepared_by_var.get() or "_______________________"
            noted_name_1 = self.noted_by_var_1.get() or "_______________________"
            noted_name_2 = self.noted_by_var_2.get() or "_______________________"
            checked_name = self.checked_by_var.get() or "_______________________"

            # Helper for signatory paragraph
            def add_signatory_para(text, size=8, alignment=1, space_after=2):
                p = doc.add_paragraph()
                run = p.add_run(text)
                run.font.name = 'Helvetica'
                run.font.size = Pt(size)
                p.alignment = alignment
                p.paragraph_format.space_after = Pt(space_after) # Add small space after

            footertable1 = doc.add_table(rows=3, cols=2)
            footertable1.autofit = False # Use manual widths
            set_cell_style(footertable1.cell(0, 0), "Prepared by:", align = 'center')
            set_cell_style(footertable1.cell(0, 1), "Checked by:", align = 'center')
            set_cell_style(footertable1.cell(1, 0), prepared_name, align = 'center')
            set_cell_style(footertable1.cell(1, 1), checked_name, align= 'center')
            set_cell_style(footertable1.cell(2, 0), "HOA Treasurer", align = 'center')
            set_cell_style(footertable1.cell(2, 1), "HOA Auditor", align= 'center')

            # Switch back to single column for "Noted by:" title
            doc.add_section(WD_SECTION.CONTINUOUS)
            noted_section = doc.sections[-1]
            sectPr = noted_section._sectPr
            cols = sectPr.xpath("./w:cols")[0]
            cols.set(qn("w:num"), "1")
            # Adjust margins/spacing for this single-column section if needed
            # noted_section.top_margin = Inches(0.1) # Example: reduce top margin
            # noted_section.bottom_margin = Inches(0.1)

            # Add "Noted by:" title centrally
            p7 = doc.add_paragraph()
            p7.add_run("Noted by:")
            p7.alignment = 1 # Center
            p7.runs[0].font.name = 'Helvetica'
            p7.runs[0].font.size = Pt(8)
            p7.paragraph_format.space_before = Pt(10) # Add space before
            p7.paragraph_format.space_after = Pt(5)   # Add space after

            # Noted by - HOA President (left column)
            footertable = doc.add_table(rows=2, cols=2)
            footertable.autofit = False # Use manual widths
            set_cell_style(footertable.cell(0, 0), noted_name_1, align = 'center')
            set_cell_style(footertable.cell(0, 1), noted_name_2, align = 'center')
            set_cell_style(footertable.cell(1, 0), "HOA President", align = 'center')
            set_cell_style(footertable.cell(1, 1), "CHUDD HCD-CORDS", align= 'center')

            # --- Save Document ---
            doc.save(filename)
            logging.info(f"Word document successfully saved to {filename}")
            messagebox.showinfo("Success", f"Word document saved to {filename}")
            return filename
        except Exception as e:
            logging.exception("Error saving to Word") # Log full traceback
            messagebox.showerror("Error", f"Error saving to Word: {str(e)}\n\nMake sure you have python-docx installed.")
            return None

    # --- MODIFIED export_to_pdf ---
    def export_to_pdf(self):
        """Export data to a single-page PDF matching the Word document format, including logo and address."""
        try:
            def format_amount(value):
                if value:
                    try:
                        # Ensure value is string, remove commas, then format
                        str_value = str(value).replace(',', '')
                        if not str_value: return ""
                        amount = Decimal(str_value)
                        return f"{amount:,.2f}"
                    except Exception as e:
                        logging.warning(f"Could not format amount '{value}' for PDF: {e}")
                        return str(value) # Return as string if conversion fails
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

            # Folio size (8.5 x 13 inches)
            folio_size = (8.5 * inch, 13 * inch)
            doc = SimpleDocTemplate(
                filename,
                pagesize=folio_size,
                topMargin=0.5*inch,
                bottomMargin=0.5*inch,
                leftMargin=0.5*inch,
                rightMargin=0.5*inch
            )
            styles = getSampleStyleSheet()
            elements = []

            # --- Create Custom Styles ---
            # Base Header Style (for Address) - HOA Name Removed
            header_style = styles['Normal']
            header_style.alignment = 1  # Center
            header_style.fontSize = 10
            header_style.leading = 12
            header_style.fontName = 'Helvetica'

            # Title Style (CASH FLOW STATEMENT)
            addressStyle = ParagraphStyle(name = ' normal', fontName = 'Helvetica', fontSize = 10, leading = 14, alignment =1)
            titleStyle = ParagraphStyle(name = ' normal', fontName = 'Helvetica-Bold', fontSize = 12, leading = 14, alignment =1)
            dateStyle = ParagraphStyle(name = ' normal', fontName = 'Helvetica', fontSize = 8, leading = 10, alignment =1, spaceBefore=6)  
            tablestyle = ParagraphStyle(name = ' normal', fontName = 'Helvetica', fontSize = 8, leading = 10)
            tableboldStyle = ParagraphStyle(name = ' normal', fontName = 'Helvetica-Bold', fontSize = 10, leading = 10, spaceAfter=10, spaceBefore=4)
            footerStyle = ParagraphStyle(name = ' normal', fontName = 'Helvetica', fontSize = 8, leading = 10, alignment =1)
            notedStyle = ParagraphStyle(name = ' normal', fontName = 'Helvetica', fontSize = 8, leading = 12, alignment =1)

            # --- Header with Logo and Address ---
            logo_path = self.logo_path_var.get()
            address_text = self.address_var.get() or " " # Use space if empty
            # hoa_name = "Buena Oro Homeowners Association Inc." # ** REMOVED **

            logo_img = None
            logo_placeholder_text = ""
            if logo_path and os.path.exists(logo_path):
                try:
                    # --- OPTIMIZED LOGO SIZE ---
                    # Set target 2x2 inch size. ReportLab will scale proportionally
                    # to fit *within* this box, preserving aspect ratio. It won't distort.
                    target_w = 1.18 * inch
                    target_h = 1.18 * inch
                    logo_img = Image(logo_path, width=target_w, height=target_h)
                    logo_img.hAlign = 'CENTER' # Align within its cell
                    logo_img.vAlign = 'MIDDLE'
                    logging.info(f"Included logo from: {logo_path}, scaled within {target_w/inch:.1f}x{target_h/inch:.1f} inches")
                except Exception as e:
                    logging.warning(f"Could not load or process logo image '{logo_path}': {e}")
                    logo_placeholder_text = "[Logo Error]"
            elif logo_path:
                logging.warning(f"Logo path specified but not found: {logo_path}")
                logo_placeholder_text = "[Logo N/A]"
            else:
                 logo_placeholder_text = "" # Empty if no logo path

            # If image failed or not provided, use placeholder
            if logo_img is None:
                logo_cell_content = Paragraph(logo_placeholder_text, styles['Italic'])
            else:
                logo_cell_content = logo_img
    
            # Assemble header text elements (HOA Name removed)
            header_text_elements = [
                # Paragraph(hoa_name, header_style), # ** REMOVED **
                Paragraph(address_text, addressStyle), # Add the address
                Spacer(1, 12), # Small space
                Paragraph("CASH FLOW STATEMENT", titleStyle),
                Paragraph(f"For the Month of {self.format_date_for_display(self.date_var.get())}", dateStyle)
            ]

            # Use a table for layout: Logo | Text
            # Calculate available width for text column
            page_width = folio_size[0] - doc.leftMargin - doc.rightMargin
            # --- OPTIMIZED COLUMN WIDTHS ---
            logo_col_width = 1.58 * inch # Assign 2 inches for logo column
            text_col_width = 4.5 * inch # Remaining width for text

            header_table_data = [[logo_cell_content, header_text_elements]]
            header_table = Table(header_table_data, colWidths=[logo_col_width, text_col_width], hAlign = 'LEFT', vAlign = 'LEFT')
            header_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0)
                #('GRID', (0, 0), (-1, -1), 0.5, colors.grey),  # Add border
            ]))

            elements.append(header_table)
            elements.append(Spacer(1, 12)) # Space after header section

            # --- Cash Flow Sections (Tables) ---
            # (Table generation code remains largely the same, ensure format_amount is used)
            # Calculate table column widths based on page width
            data_label_width = page_width * 0.65 # Approx 65% for description
            data_value_width = page_width * 0.35 # Approx 35% for amount

            # Beginning Cash Balances
            elements.append(Paragraph("Beginning Cash Balances", tableboldStyle)) # Use bold style for section titles
            beg_data = [
                ["Cash in Bank-beg", format_amount(self.variables['cash_bank_beg'].get())],
                ["Cash on Hand-beg", format_amount(self.variables['cash_hand_beg'].get())]
            ]
            beg_table = Table(beg_data, colWidths=[data_label_width, data_value_width], rowHeights=[14 for _ in beg_data]) # Adjusted row height
            beg_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black), # Thinner grid lines
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5), # Added padding
                ('RIGHTPADDING', (1, 0), (1, -1), 5), # Added padding
            ]))
            elements.append(beg_table)
            elements.append(Spacer(1, 6))

            # Cash Inflows
            elements.append(Paragraph("Cash Inflows", tableboldStyle))
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
                ["Total Cash Receipts", format_amount(self.variables['total_receipts'].get())] # Bold total row
            ]
            inflows_table = Table(inflows_data, colWidths=[data_label_width, data_value_width], rowHeights=[14]*len(inflows_data))
            inflows_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'), # Default font
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (1, 0), (1, -1), 5),
            ]))
            elements.append(inflows_table)
            elements.append(Spacer(1, 6))

            # Cash Outflows
            elements.append(Paragraph("Less: Cash Outflows", tableboldStyle))
            outflows_data = [
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
                ["Others (Outflow)", format_amount(self.variables['outflows_others'].get())],
                ["Cash Outflows/Disbursements", format_amount(self.variables['cash_outflows'].get())] # Bold total row
            ]
            outflows_table = Table(outflows_data, colWidths=[data_label_width, data_value_width], rowHeights=[14]*len(outflows_data))
            outflows_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (1, 0), (1, -1), 5),
            ]))
            elements.append(outflows_table)
            elements.append(Spacer(1, 6))

            # Ending Cash Balance
            elements.append(Paragraph("Ending Cash Balance", tableboldStyle))
            ending_data = [
                 ["Ending Cash Balance", format_amount(self.variables['ending_cash'].get())]
            ]
            ending_table = Table(ending_data, colWidths=[data_label_width, data_value_width], rowHeights=[14])
            ending_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'), # Already bold via Paragraph
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (1, 0), (1, -1), 5),
            ]))
            elements.append(ending_table)
            elements.append(Spacer(1, 6))

            # Breakdown of Cash
            elements.append(Paragraph("Breakdown of Cash", tableboldStyle))
            breakdown_data = [
                ["Cash in Bank", format_amount(self.variables['ending_cash_bank'].get())],
                ["Cash on Hand", format_amount(self.variables['ending_cash_hand'].get())]
            ]
            breakdown_table = Table(breakdown_data, colWidths=[data_label_width, data_value_width], rowHeights=[14, 14])
            breakdown_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (1, 0), (1, -1), 5),
            ]))
            elements.append(breakdown_table)
            elements.append(Spacer(1, 12))

            # --- Footer Signatories ---
            # (Signatory layout remains the same, using footer_style and noted_style)
            prepared_name = self.prepared_by_var.get() or "_______________________"
            noted_name_1 = self.noted_by_var_1.get() or "_______________________"
            noted_name_2 = self.noted_by_var_2.get() or "_______________________"
            checked_name = self.checked_by_var.get() or "_______________________"

            # Use a table for 2-column layout of signatories
            col_width = (page_width / 2) - (0.1 * inch) # Slightly less than half for padding

            # Prepared by / Checked by row
            prep_check_data = [
                [Paragraph(f"Prepared by:<br/>{prepared_name}<br/>HOA Treasurer", footerStyle)],
                [Paragraph(f"Checked by:<br/>{checked_name}<br/>HOA Auditor", footerStyle)]
            ]
            prep_check_table = Table([prep_check_data[0] + prep_check_data[1]], colWidths=[col_width, col_width]) # Combine into one row
            prep_check_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0,0), (-1,-1), 0),
                ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ]))
            elements.append(prep_check_table)
            elements.append(Spacer(1, 12))

            # Noted by Title
            elements.append(Paragraph("Noted by:", notedStyle))
            elements.append(Spacer(1, 6))

            # Noted by Names row
            noted_data = [
                 [Paragraph(f"{noted_name_1}<br/>HOA President", footerStyle)],
                 [Paragraph(f"{noted_name_2}<br/>CHUDD HCD-CORDS", footerStyle)]
            ]
            noted_table = Table([noted_data[0] + noted_data[1]], colWidths=[col_width, col_width])
            noted_table.setStyle(TableStyle([
                 ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                 ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                 ('LEFTPADDING', (0,0), (-1,-1), 0),
                 ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ]))
            elements.append(noted_table)

            # --- Build PDF ---
            doc.build(elements)
            logging.info(f"PDF successfully exported to {filename}")
            messagebox.showinfo("Success", f"PDF successfully exported to {filename}")
            return filename

        except Exception as e:
            logging.exception("Error exporting to PDF") # Log full traceback
            messagebox.showerror("Error", f"Error exporting to PDF: {str(e)}\n\nMake sure you have ReportLab installed.")
            return None
           
    
    def load_from_documentpdf(self):
        """Load data from either a Word or PDF document."""
        # (This function remains the same as it doesn't load logo/address)
        try:
            filename = filedialog.askopenfilename(
                filetypes=[
                ("Documents", "*.docx *.pdf"), # Combined filter
                ("Word Documents", "*.docx"),
                ("PDF Files", "*.pdf"),
                ("All Files", "*.*"),
            ],
             title="Select a Document (DOCX or PDF)"
            )
            if not filename:
                return

            if filename.lower().endswith('.pdf'):
                self.load_from_pdf(filename)
            elif filename.lower().endswith('.docx'):
                 self.load_from_docx(filename)
            else:
                messagebox.showerror("Error", "Unsupported file format. Please select a PDF or DOCX file.")
                return # Stop if unsupported

            # messagebox.showinfo("Success", f"Loaded data from {os.path.basename(filename)}") # Already shown in specific load functions

        except Exception as e:
             logging.exception("Error loading document") # Log full traceback
             messagebox.showerror("Error", f"Error loading document: {str(e)}")

# --- END OF FILE file_handler.py ---a