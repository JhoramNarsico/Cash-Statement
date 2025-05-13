# --- START OF FILE email_attachments_window.py ---

import customtkinter as ctk
import tkinter as tk # Need tkinter for Listbox and Scrollbar
from tkinter import filedialog, messagebox
import os
import logging
import time # Added for potential delay simulation or button disabling feedback

class EmailAttachmentsWindow(ctk.CTkToplevel):
    """GUI for managing additional email attachments and sending the email."""
    # --- Updated __init__ signature ---
    def __init__(self, parent, email_sender, recipient_emails_var):
        super().__init__(parent)
        self.parent = parent
        self.email_sender = email_sender # Reference to the EmailSender instance
        self.recipient_emails_var = recipient_emails_var # Store the shared StringVar

        self.title("Manage & Send Email Attachments")
        # Increased height slightly for recipient field
        self.geometry("550x500")
        self.resizable(True, True)
        self.minsize(400, 350) # Adjusted min height
        self.transient(parent)
        self.grab_set()

        self.center_window()

        # --- UI Elements ---
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.pack(padx=15, pady=15, fill="both", expand=True)

        # --- Adjust Grid ---
        self.main_frame.grid_rowconfigure(0, weight=0) # Recipient frame (fixed)
        self.main_frame.grid_rowconfigure(1, weight=0) # Attachments Label (fixed)
        self.main_frame.grid_rowconfigure(2, weight=1) # Listbox area expands
        self.main_frame.grid_rowconfigure(3, weight=0) # Button frame (fixed)
        self.main_frame.grid_columnconfigure(0, weight=1) # Main column expands

        # --- New Frame for Recipients ---
        recipient_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        recipient_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        recipient_frame.grid_columnconfigure(1, weight=1) # Make entry expand

        ctk.CTkLabel(
            recipient_frame,
            text="Recipients:",
            font=("Roboto", 12),
            anchor="w"
        ).grid(row=0, column=0, sticky="w", padx=(0, 5))

        recipient_entry = ctk.CTkEntry(
            recipient_frame,
            textvariable=self.recipient_emails_var, # Use the passed StringVar
            font=("Roboto", 12),
            corner_radius=6,
            fg_color="#FAFAFA",
            text_color="#333333",
            border_color="#B0BEC5",
            placeholder_text="Enter emails, comma-separated"
        )
        recipient_entry.grid(row=0, column=1, sticky="ew")
        # -------------------------------

        ctk.CTkLabel(
            self.main_frame,
            text="Additional Email Attachments",
            font=("Roboto", 16, "bold")
        ).grid(row=1, column=0, pady=(5, 5), sticky="ew") # Adjusted row

        # --- Listbox Frame ---
        list_frame = ctk.CTkFrame(self.main_frame, fg_color="#EAEAEA", corner_radius=5)
        list_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 10)) # Adjusted row
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_columnconfigure(1, weight=0)

        self.listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            font=("Roboto", 26, "bold"),
            bg="#FFFFFF",
            fg="#333333",
            selectbackground="#1A73E8",
            selectforeground="#FFFFFF",
            borderwidth=0,
            highlightthickness=0,
            activestyle='none'
        )
        self.listbox.grid(row=0, column=0, sticky="nsew", padx=(5,0), pady=5)

        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(0,5), pady=5)
        self.listbox.configure(yscrollcommand=scrollbar.set)

        # --- Button Frame ---
        button_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        button_frame.grid(row=3, column=0, sticky="ew", pady=(5, 0)) # Adjusted row
        button_frame.grid_columnconfigure((0, 1, 2), weight=1)

        self.add_button = ctk.CTkButton(
            button_frame,
            text="Add File(s)",
            command=self.add_attachments,
            font=("Roboto", 12),
            corner_radius=6,
            fg_color="#4CAF50",
            hover_color="#388E3C"
        )
        self.add_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.remove_button = ctk.CTkButton(
            button_frame,
            text="Remove Selected",
            command=self.remove_selected_attachment,
            font=("Roboto", 12),
            corner_radius=6,
            fg_color="#D32F2F",
            hover_color="#C62828"
        )
        self.remove_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.send_button = ctk.CTkButton(
            button_frame,
            text="Send Email",
            command=self.send_and_close,
            font=("Roboto", 12, "bold"),
            corner_radius=6,
            fg_color="#2196F3",
            hover_color="#1976D2"
        )
        self.send_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        # Populate the listbox initially
        self.populate_listbox()

        self.bind("<Escape>", lambda e: self.close_window())

    def center_window(self):
        """Centers the window on the parent."""
        self.update_idletasks()
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        win_width = self.winfo_width()
        win_height = self.winfo_height()
        x = parent_x + (parent_width // 2) - (win_width // 2)
        y = parent_y + (parent_height // 2) - (win_height // 2)
        self.geometry(f"+{x}+{y}")

    def populate_listbox(self):
        """Clears and refills the listbox with current attachments."""
        self.listbox.delete(0, tk.END)
        if not self.email_sender.attachments:
            self.listbox.insert(tk.END, " No additional attachments added.")
            self.listbox.itemconfig(0, fg="grey") # Style the placeholder text
        else:
            for filepath in self.email_sender.attachments:
                filename = os.path.basename(filepath)
                self.listbox.insert(tk.END, f" {filename}")

    def add_attachments(self):
        """Opens file dialog to select files and adds them to the list."""
        filetypes = [("All files", "*.*")]
        filepaths = filedialog.askopenfilenames(
            title="Select Attachments",
            filetypes=filetypes
        )

        added_count = 0
        if filepaths:
            for path in filepaths:
                if os.path.exists(path):
                    if path not in self.email_sender.attachments:
                        self.email_sender.attachments.append(path)
                        logging.info(f"Attachment added: {path}")
                        added_count += 1
                    else:
                        logging.info(f"Attachment already in list, skipped: {path}")
                else:
                    messagebox.showwarning("File Not Found", f"The selected file could not be found:\n{path}", parent=self)
                    logging.warning(f"User tried to add non-existent file: {path}")

            if added_count > 0:
                self.populate_listbox()

    def remove_selected_attachment(self):
        """Removes the selected attachment(s) from the list."""
        selected_indices = self.listbox.curselection()

        if not selected_indices:
            messagebox.showinfo("No Selection", "Please select an attachment from the list to remove.", parent=self)
            return

        removed_count = 0
        indices_to_remove = sorted(selected_indices, reverse=True)

        for index in indices_to_remove:
            try:
                if not self.email_sender.attachments and index == 0:
                    logging.debug("Attempted to remove placeholder text.")
                    continue

                if 0 <= index < len(self.email_sender.attachments):
                    removed_path = self.email_sender.attachments.pop(index)
                    logging.info(f"Attachment removed: {removed_path}")
                    removed_count += 1
                else:
                     logging.warning(f"Attempted to remove invalid listbox index: {index}")
            except IndexError:
                 logging.error(f"IndexError removing attachment at index {index}. List length: {len(self.email_sender.attachments)}")

        if removed_count > 0:
            self.populate_listbox()

    def send_and_close(self):
        """Sends the email using EmailSender and then closes the window."""
        logging.info("Attempting to send email from attachments window...")

        if not self.email_sender:
            messagebox.showerror("Error", "Email sending component is not available.", parent=self)
            logging.error("send_and_close called but self.email_sender is None.")
            return

        # --- Check recipients field directly from the StringVar ---
        if not self.recipient_emails_var.get().strip():
             messagebox.showerror("Missing Information", "Please enter recipient email addresses.", parent=self)
             return # Don't proceed if recipients are empty
        # -------------------------------------------------------


        try:
            self.send_button.configure(state="disabled", text="Sending...")
            self.update_idletasks()
        except tk.TclError:
            logging.warning("Could not disable send button (window might be closing).")
            pass

        try:
            # EmailSender reads recipient_emails_var internally
            result = self.email_sender.send_email()

            if result.get("status") == "success":
                messagebox.showinfo("Email Sent", result.get("message", "Email sent successfully!"), parent=self)
            else:
                messagebox.showerror("Email Error", result.get("message", "Failed to send email."), parent=self)

        except Exception as e:
            logging.exception("Unexpected error during send_and_close.")
            messagebox.showerror("Error", f"An unexpected error occurred: {e}", parent=self)

        finally:
            try:
                if self.winfo_exists():
                    self.send_button.configure(state="normal", text="Send Email")
            except tk.TclError:
                pass

            # Close ONLY if email was successful (optional, keeps window open on error)
            if result and result.get("status") == "success":
                 self.close_window()
            # OR: Always close
            # self.close_window()

    def close_window(self):
        """Closes the attachments window."""
        logging.debug("Closing EmailAttachmentsWindow.")
        self.grab_release()
        self.destroy()

# --- END OF FILE email_attachments_window.py ---