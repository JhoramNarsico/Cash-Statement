import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from tkinter import messagebox
import logging # Import logging

class EmailSender:
    def __init__(self, sender_email, sender_password, recipient_emails_var, file_handler):
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.recipient_emails_var = recipient_emails_var
        self.file_handler = file_handler

    def _safe_remove_file(self, filepath, context=""):
        """Safely removes a file if it exists."""
        if filepath and os.path.exists(filepath):
            try:
                os.remove(filepath)
                logging.info(f"Successfully removed {context}file: {filepath}")
            except Exception as e_remove:
                logging.error(f"Failed to remove {context}file {filepath}: {e_remove}")
        elif filepath:
            logging.warning(f"Attempted to remove {context}file, but it does not exist: {filepath}")
        # else: filepath is None, nothing to do

    def send_email(self):
        pdf_filename = None  # Initialize to None
        docx_filename = None # Initialize to None
        try:
            recipient_emails = [email.strip() for email in self.recipient_emails_var.get().split(',') if email.strip()]

            if not recipient_emails:
                messagebox.showerror("Error", "Please fill in the recipient email field.")
                return

            pdf_filename = self.file_handler.export_to_pdf()
            if not pdf_filename:
                # export_to_pdf should show its own error, so just return
                return
            
            docx_filename = self.file_handler.save_to_docx()
            if not docx_filename:
                # save_to_docx should show its own error
                self._safe_remove_file(pdf_filename, "PDF (after DOCX gen failed) ")
                return

            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = ", ".join(recipient_emails)
            msg['Subject'] = f"Cash Flow Statement - {self.file_handler.format_date_for_display(self.file_handler.date_var.get())}"

            body = f"Attached is the cash flow statement for {self.file_handler.format_date_for_display(self.file_handler.date_var.get())} in both PDF and Word formats.\n\nRegards,\nYour's truly"
            msg.attach(MIMEText(body, 'plain'))

            with open(pdf_filename, 'rb') as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(pdf_filename))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_filename)}"'
                msg.attach(part)

            with open(docx_filename, 'rb') as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(docx_filename))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(docx_filename)}"'
                msg.attach(part)

            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.sender_email, self.sender_password)
            server.send_message(msg)
            server.quit()

            messagebox.showinfo("Success", f"Email with PDF and Word files sent to {', '.join(recipient_emails)}!")
            self._safe_remove_file(pdf_filename, "PDF (after success) ")
            self._safe_remove_file(docx_filename, "DOCX (after success) ")

        except smtplib.SMTPAuthenticationError as e:
            messagebox.showerror("Error", f"Authentication failed: {str(e)}\nCheck your email and app password in settings.")
            self._safe_remove_file(pdf_filename, "PDF (on auth error) ")
            self._safe_remove_file(docx_filename, "DOCX (on auth error) ")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email: {str(e)}\nEnsure your email credentials are correct, you have an internet connection, and the generated files are accessible.")
            logging.exception("Email sending failed") # Log full traceback for this generic error
            self._safe_remove_file(pdf_filename, "PDF (on general error) ")
            self._safe_remove_file(docx_filename, "DOCX (on general error) ")
