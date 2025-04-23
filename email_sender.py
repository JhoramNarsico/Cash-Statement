import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from tkinter import messagebox

class EmailSender:
    def __init__(self, sender_email, sender_password, recipient_emails_var, file_handler):
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.recipient_emails_var = recipient_emails_var
        self.file_handler = file_handler

    def send_email(self):
        try:
            recipient_emails = [email.strip() for email in self.recipient_emails_var.get().split(',') if email.strip()]

            if not recipient_emails:
                messagebox.showerror("Error", "Please fill in the recipient email field.")
                return

            pdf_filename = self.file_handler.export_to_pdf()
            if not pdf_filename:
                return
            docx_filename = self.file_handler.save_to_docx()
            if not docx_filename:
                os.remove(pdf_filename)
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
            os.remove(pdf_filename)
            os.remove(docx_filename)

        except smtplib.SMTPAuthenticationError as e:
            messagebox.showerror("Error", f"Authentication failed: {str(e)}\nCheck your hardcoded email and app password.")
            if 'pdf_filename' in locals():
                os.remove(pdf_filename)
            if 'docx_filename' in locals():
                os.remove(docx_filename)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email: {str(e)}\nEnsure your email credentials are correct and you have an internet connection.")
            if 'pdf_filename' in locals():
                os.remove(pdf_filename)
            if 'docx_filename' in locals():
                os.remove(docx_filename)