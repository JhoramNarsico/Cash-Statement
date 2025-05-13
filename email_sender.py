# --- START OF FILE email_sender.py ---

import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
import mimetypes
import logging

class EmailSender:
    def __init__(self, settings_manager, recipient_emails_var, file_handler):
        self.settings_manager = settings_manager
        self.recipient_emails_var = recipient_emails_var
        self.file_handler = file_handler # Still needed for date formatting maybe?
        self.attachments = []

    def _safe_remove_file(self, filepath, context=""):
        """Safely removes a file if it exists."""
        # This function remains useful if temporary files were ever needed again.
        if filepath and os.path.exists(filepath):
            try:
                os.remove(filepath)
                logging.info(f"Successfully removed {context}file: {filepath}")
            except Exception as e_remove:
                logging.error(f"Failed to remove {context}file {filepath}: {e_remove}")
        elif filepath:
            logging.warning(f"Attempted to remove {context}file, but it does not exist: {filepath}")

    def send_email(self):
        """Sends email with only the additional attachments managed by the user.
           Returns a dictionary with 'status' and 'message'."""
        # temp_pdf_path = None # No longer generating these
        # temp_docx_path = None # No longer generating these
        try:
            # Get current credentials using self.settings_manager
            current_sender_email = self.settings_manager.get_setting("sender_email")
            current_sender_password = self.settings_manager.get_setting("sender_password")
            if not current_sender_email or not current_sender_password:
                 return {"status": "error", "message": "Sender email or password not configured in settings."}

            recipient_emails = [email.strip() for email in self.recipient_emails_var.get().split(',') if email.strip()]

            if not recipient_emails:
                return {"status": "error", "message": "Please fill in the recipient email field."}

            # Check if there are any user-added attachments
            if not self.attachments:
                return {"status": "error", "message": "No attachments selected. Please add files using 'Manage Attachments'."}


            msg = MIMEMultipart()
            msg['From'] = current_sender_email
            msg['To'] = ", ".join(recipient_emails)
            # Subject can be more generic now, or keep the date if relevant contextually
            # formatted_display_date = self.file_handler.format_date_for_display(self.file_handler.date_var.get()) # Might still want date context?
            # msg['Subject'] = f"Cash Flow Statement Related Files - {formatted_display_date}"
            msg['Subject'] = f"Requested File Attachments" # More generic subject

            # Update body text
            body_lines = [
                f"Please find the requested file(s) attached."
                # f"Attached are the files selected via the 'Manage Attachments' feature."
            ]
            body_lines.append("\n\nRegards,\nYours truly")
            body = "".join(body_lines)
            msg.attach(MIMEText(body, 'plain'))

            # Attach additional files from self.attachments
            attached_count = 0
            for filepath in self.attachments:
                if not os.path.exists(filepath):
                    logging.warning(f"Skipping attachment: File not found at '{filepath}'")
                    continue

                ctype, encoding = mimetypes.guess_type(filepath)
                if ctype is None or encoding is not None:
                    ctype = 'application/octet-stream'
                maintype, subtype = ctype.split('/', 1)

                try:
                    filename = os.path.basename(filepath)
                    with open(filepath, 'rb') as fp:
                        if maintype == 'text':
                            try:
                                file_content = fp.read().decode('utf-8')
                                part = MIMEText(file_content, _subtype=subtype, _charset='utf-8')
                            except UnicodeDecodeError:
                                logging.warning(f"Could not decode text file {filename} as UTF-8, sending as binary.")
                                fp.seek(0)
                                part = MIMEBase(maintype, subtype)
                                part.set_payload(fp.read())
                                encoders.encode_base64(part)
                        elif maintype == 'image':
                            part = MIMEBase(maintype, subtype)
                            part.set_payload(fp.read())
                            encoders.encode_base64(part)
                        elif maintype == 'application':
                             part = MIMEApplication(fp.read(), _subtype=subtype)
                        else:
                            part = MIMEBase(maintype, subtype)
                            part.set_payload(fp.read())
                            encoders.encode_base64(part)

                        part.add_header('Content-Disposition', 'attachment', filename=filename)
                        msg.attach(part)
                        logging.info(f"Attached additional file: {filename}")
                        attached_count += 1

                except FileNotFoundError:
                     logging.warning(f"Attachment file disappeared before attaching: {filepath}")
                except Exception as e_attach:
                    logging.error(f"Failed to attach file '{filepath}': {e_attach}")
            # --- End of attachment section ---

            if attached_count == 0:
                 # This case might happen if files existed when added but were deleted before sending
                 return {"status": "error", "message": "No valid attachments found to send. Please check the files in 'Manage Attachments'."}


            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(current_sender_email, current_sender_password)
            server.send_message(msg)
            server.quit()

            # Update success message
            success_message = f"Email with {attached_count} file(s) sent to {', '.join(recipient_emails)}!"

            return {"status": "success", "message": success_message}

        except smtplib.SMTPAuthenticationError as e_auth:
            logging.error(f"SMTP Authentication Error: {e_auth}")
            return {"status": "error", "message": f"Authentication failed: {str(e_auth)}\nCheck your email and app password in settings."}
        except Exception as e_general:
            logging.exception("Email sending process failed")
            return {"status": "error", "message": f"Failed to send email using '{current_sender_email}': {str(e_general)}\nEnsure credentials, connection, and file access are correct."}
        # finally: # No temporary files generated, so cleanup isn't strictly needed here anymore
            # self._safe_remove_file(temp_pdf_path, "Temporary PDF ")
            # self._safe_remove_file(temp_docx_path, "Temporary DOCX ")
# --- END OF FILE email_sender.py ---