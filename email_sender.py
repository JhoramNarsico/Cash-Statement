import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import logging 

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

    def send_email(self):
        """Generates temporary PDF and DOCX files, sends email with these attachments, then cleans them up.
           Returns a dictionary with 'status' and 'message'."""
        temp_pdf_path = None
        temp_docx_path = None
        try:
            recipient_emails = [email.strip() for email in self.recipient_emails_var.get().split(',') if email.strip()]

            if not recipient_emails:
                return {"status": "error", "message": "Please fill in the recipient email field."}

            # Generate temporary PDF for emailing
            pdf_gen_result = self.file_handler.generate_temp_pdf()
            if pdf_gen_result.get("status") != "success":
                # Error message already in pdf_gen_result['message']
                return {"status": "error", "message": f"Failed to prepare PDF for email: {pdf_gen_result.get('message', 'Unknown PDF generation error')}"}
            temp_pdf_path = pdf_gen_result.get("filename")
            
            # Generate temporary DOCX for emailing
            docx_gen_result = self.file_handler.generate_temp_docx()
            if docx_gen_result.get("status") != "success":
                 # Error message already in docx_gen_result['message']
                return {"status": "error", "message": f"Failed to prepare Word document for email: {docx_gen_result.get('message', 'Unknown DOCX generation error')}"}
            temp_docx_path = docx_gen_result.get("filename")

            if not temp_pdf_path or not temp_docx_path: # Should be caught by status checks above
                error_parts = []
                if not temp_pdf_path: error_parts.append("Temporary PDF was not created.")
                if not temp_docx_path: error_parts.append("Temporary Word document was not created.")
                return {"status": "error", "message": " ".join(error_parts)}

            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = ", ".join(recipient_emails)
            # Use a clean date format for the subject
            formatted_display_date = self.file_handler.format_date_for_display(self.file_handler.date_var.get())
            msg['Subject'] = f"Cash Flow Statement - {formatted_display_date}"

            body = f"Attached is the cash flow statement for {formatted_display_date} in both PDF and Word formats.\n\nRegards,\nYour's truly"
            msg.attach(MIMEText(body, 'plain'))

            # Use a generic filename for the attachment as seen by the recipient
            attachment_date_str = self.file_handler.date_var.get().replace("/", "-") # e.g., 12-31-2023
            
            with open(temp_pdf_path, 'rb') as f:
                part_pdf = MIMEApplication(f.read(), Name=f"CashFlowStatement_{attachment_date_str}.pdf")
                part_pdf['Content-Disposition'] = f'attachment; filename="CashFlowStatement_{attachment_date_str}.pdf"'
                msg.attach(part_pdf)

            with open(temp_docx_path, 'rb') as f:
                part_docx = MIMEApplication(f.read(), Name=f"CashFlowStatement_{attachment_date_str}.docx")
                part_docx['Content-Disposition'] = f'attachment; filename="CashFlowStatement_{attachment_date_str}.docx"'
                msg.attach(part_docx)

            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.sender_email, self.sender_password)
            server.send_message(msg)
            server.quit()
            
            return {"status": "success", "message": f"Email with PDF and Word files sent to {', '.join(recipient_emails)}!"}

        except smtplib.SMTPAuthenticationError as e_auth:
            logging.error(f"SMTP Authentication Error: {e_auth}")
            return {"status": "error", "message": f"Authentication failed: {str(e_auth)}\nCheck your email and app password in settings."}
        except Exception as e_general:
            logging.exception("Email sending process failed")
            return {"status": "error", "message": f"Failed to send email: {str(e_general)}\nEnsure credentials, connection, and file access are correct."}
        finally:
            # Always attempt to clean up temporary files
            if temp_pdf_path:
                self._safe_remove_file(temp_pdf_path, "Temporary PDF ")
            if temp_docx_path:
                self._safe_remove_file(temp_docx_path, "Temporary DOCX ")
