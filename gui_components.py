# --- START OF FILE gui_components.py ---

# --- START OF FILE gui_components.py ---

import tkinter
import customtkinter as ctk
from tkinter import Tk, messagebox, filedialog
import datetime
import time
import logging
import os
import sys # Added for resource_path
from setting import SettingsWindow
from email_attachments_window import EmailAttachmentsWindow # <-- Import the new window
from PIL import Image  # Requires Pillow
import webbrowser # Added for hyperlink functionality
import threading # Added for loading indicator
import queue     # Added for loading indicator

try:
    from hover_calendar import HoverCalendar
    logging.info("HoverCalendar imported successfully.")
except ImportError:
    HoverCalendar = None
    logging.warning("HoverCalendar not found. Calendar functionality will be disabled.")

# --- Helper function for PyInstaller asset bundling ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
# --- End helper function ---

class LoadingWindow(ctk.CTkToplevel):
    # ... (LoadingWindow code remains the same) ...
    def __init__(self, parent, title="Loading..."):
        super().__init__(parent)
        self.title(title)
        # Make sure parent is valid and drawn
        parent.update_idletasks()
        self.geometry("300x150")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set() # Make it modal

        # Center the window relative to the parent
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        win_width = 300
        win_height = 150
        x = parent_x + (parent_width // 2) - (win_width // 2)
        y = parent_y + (parent_height // 2) - (win_height // 2)
        self.geometry(f"{win_width}x{win_height}+{x}+{y}")

        self.protocol("WM_DELETE_WINDOW", lambda: None) # Prevent closing by user

        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.message_label = ctk.CTkLabel(self.main_frame, text="Processing...", font=("Roboto", 12))
        self.message_label.pack(pady=(20, 10))

        self.progress_bar = ctk.CTkProgressBar(self.main_frame, mode='indeterminate')
        self.progress_bar.pack(pady=10, padx=20, fill="x")
        self.progress_bar.start()

    def update_message(self, message):
        self.message_label.configure(text=message)
        self.update_idletasks() # Ensure message update is visible

    def close_window(self):
        if self.winfo_exists(): # Check if window still exists
            self.progress_bar.stop()
            self.grab_release()
            self.destroy()

class GUIComponents:
    def __init__(self, root, variables, title_var, date_var, display_date, calculator, file_handler, email_sender, settings_manager):
        self.root = root
        self.variables = variables
        self.title_var = title_var
        self.date_var = date_var
        self.display_date = display_date
        self.calculator = calculator
        self.file_handler = file_handler
        self.email_sender = email_sender
        self.settings_manager = settings_manager

        # ... (Initialization code remains the same) ...
        self.required_vars = [
            'logo_path_var', 'address_var',
            'recipient_emails_var', # Still required here for sharing
            'prepared_by_var', 'noted_by_var_1',
            'noted_by_var_2', 'checked_by_var', 'cash_bank_beg', 'cash_hand_beg',
            'monthly_dues', 'certifications', 'membership_fee', 'vehicle_stickers',
            'rentals', 'solicitations', 'interest_income', 'livelihood_fee',
            'inflows_others', 'snacks_meals', 'transportation', 'office_supplies',
            'printing', 'labor', 'billboard', 'cleaning', 'misc_expenses',
            'federation_fee', 'uniforms', 'bod_mtg', 'general_assembly',
            'cash_deposit', 'withholding_tax', 'refund', 'outflows_others',
            'ending_cash_bank', 'ending_cash_hand', 'total_receipts',
            'cash_outflows', 'ending_cash',
            'footer_image1_var', 'footer_image2_var'
        ]
        self._initialize_missing_variables()

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.BG_COLOR = "#F5F5F5"
        self.FRAME_COLOR = "#FFFFFF"
        self.BORDER_COLOR = "#E0E0E0"
        self.TEXT_COLOR = "#333333"
        self.ENTRY_BG_COLOR = "#FAFAFA"
        self.ENTRY_BORDER_COLOR = "#B0BEC5"
        self.DISABLED_BG_COLOR = "#ECEFF1"
        self.BUTTON_FG_COLOR = "#2196F3"
        self.BUTTON_HOVER_COLOR = "#1976D2"
        self.BUTTON_TEXT_COLOR = "#FFFFFF"
        self.DATE_BTN_FG = "#E3F2FD"
        self.DATE_BTN_HOVER = "#BBDEFB"
        self.DATE_BTN_TEXT = "#0D47A1"
        self.FOOTER_BG = "#192337"  # Solid blue for footer
        self.REMOVE_BTN_FG_COLOR = "#D32F2F" # Reddish for remove
        self.REMOVE_BTN_HOVER_COLOR = "#C62828"
        self.ATTACH_BTN_FG_COLOR = "#FF9800" # Orange for Attachments
        self.ATTACH_BTN_HOVER_COLOR = "#FB8C00"


        # ... (Font size calculations, validation, etc. remain the same) ...
        self.root.update_idletasks()
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        logging.info(f"Screen dimensions: {self.screen_width}x{self.screen_height}")
        # --- MODIFICATION HERE ---
        # Changed divisor from 65 to 70 to make base font smaller
        self.base_font_size = max(9, min(15, int(self.screen_height / 70))) # Also slightly reduced max from 16 to 15
        # --- END MODIFICATION ---
        self.title_font_size = int(self.base_font_size * 1.2)
        self.button_font_size = self.base_font_size
        self.label_font_size = self.base_font_size
        self.entry_font_size = self.base_font_size
        self.base_pad_x = int(self.base_font_size * 0.6)
        self.base_pad_y = int(self.base_font_size * 0.3)
        self.section_pad_x = self.base_pad_x * 2
        self.section_pad_y = self.base_pad_y * 2
        self.min_input_width = 160 # Kept from previous change
        self.max_input_width = 300
        self.min_column_width = 200
        self.layout_debounce_delay_ms = 100
        self.debounce_id = None

        # Register validation command for numeric entries
        self.vcmd_numeric = (self.root.register(self._validate_numeric_input), '%P')

        self.create_widgets()
        self.setup_keyboard_shortcuts()
        self.date_var.trace_add('write', self._update_display_date)
        if 'logo_path_var' in self.variables: # Ensure var exists before tracing
            self.variables['logo_path_var'].trace_add('write', self._update_logo_display)
        self._update_display_date()
        self._update_logo_display() # Initial update for logo display
        self.root.after(150, self.update_layout)
    # --- Methods (_validate_numeric_input, _initialize_missing_variables, etc. remain the same) ---

    def _validate_numeric_input(self, P):
        """Validates that the input P is a valid numeric string (allowing digits, one decimal, and commas)."""
        if P == "":
            return True  # Allow empty string (clearing the field)

        allowed_chars = "0123456789.,"
        decimal_points = 0
        for char in P:
            if char not in allowed_chars:
                return False
            if char == '.':
                decimal_points += 1

        if decimal_points > 1:
            return False

        return True

    def _initialize_missing_variables(self):
        initialized_count = 0
        missing_vars = []
        for var_key in self.required_vars:
            if var_key not in self.variables:
                default_value = ""
                if var_key == 'address_var':
                    default_value = "Default Address - Configure Me"
                elif var_key == 'footer_image1_var':
                    default_value = resource_path("chud logo.png")
                elif var_key == 'footer_image2_var':
                    default_value = resource_path("xu logo.png")
                elif var_key == 'logo_path_var': # Ensure logo_path_var is initialized
                    default_value = ""
                elif var_key == 'recipient_emails_var': # Default recipient if needed
                    default_value = ""

                self.variables[var_key] = ctk.StringVar(value=default_value)
                initialized_count += 1
                missing_vars.append(var_key)

        if initialized_count > 0:
            logging.warning(f"Initialized {initialized_count} missing StringVars: {missing_vars}")
        elif not self.variables:
            logging.error("Variables dictionary is empty!")

    def _update_display_date(self, *args):
        raw_date = self.date_var.get()
        try:
            date_obj = datetime.datetime.strptime(raw_date, "%m/%d/%Y")
            self.display_date.set(date_obj.strftime("%b %d, %Y"))
        except ValueError:
            if raw_date:
                logging.warning(f"Invalid date format entered: {raw_date}. Expected MM/DD/YYYY.")
            self.display_date.set("Select Date")

    def _update_logo_display(self, *args):
        logo_path = self.variables['logo_path_var'].get()
        display_text = "No logo selected"
        if logo_path:
            if os.path.exists(logo_path):
                filename = os.path.basename(logo_path)
                display_text = f"Selected: {filename}" if len(filename) < 40 else f"Selected: ...{filename[-37:]}"
            else:
                display_text = "Logo file not found!" # Indicate error clearly

        if hasattr(self, 'logo_path_display') and self.logo_path_display.winfo_exists():
            self.logo_path_display.configure(text=display_text)
        else:
            logging.debug("logo_path_display not available or destroyed during _update_logo_display")

    def _execute_with_loading(self, func, action_name, loading_message, *args, **kwargs):
        if not self.root.winfo_viewable():
            logging.warning(f"Root window not viewable when trying to show loading for {action_name}.")
            return

        loading_window = LoadingWindow(self.root, title=f"{action_name} in Progress")
        loading_window.update_message(loading_message)
        self.root.update_idletasks()

        result_queue = queue.Queue()

        def task_wrapper():
            try:
                result = func(*args, **kwargs)
                result_queue.put(result)
            except Exception as e:
                logging.exception(f"Unhandled exception in task_wrapper for {action_name}")
                result_queue.put({"status": "error", "message": f"Unexpected error in {action_name}: {str(e)}"})

        thread = threading.Thread(target=task_wrapper, daemon=True)
        thread.start()

        def check_queue():
            try:
                response = result_queue.get_nowait()

                if not loading_window.winfo_exists():
                    logging.warning(f"Loading window for {action_name} closed before task completion.")
                    return

                loading_window.close_window()

                if isinstance(response, dict):
                    status = response.get("status")
                    message = response.get("message")

                    if status == "success":
                        if message:
                            messagebox.showinfo("Success", message)
                        logging.info(f"Action '{action_name}' completed: {message or 'OK'}")
                    elif status == "error":
                        if message:
                            messagebox.showerror("Error", message)
                        else:
                            messagebox.showerror("Error", f"An error occurred during {action_name}.")
                        logging.error(f"Action '{action_name}' failed: {message or 'Unknown error'}")
                    elif status == "cancelled":
                        logging.info(f"Action '{action_name}' was cancelled by the user.")
                    else:
                        logging.warning(f"Unexpected response structure from {action_name}: {response}")
                        if message:
                             messagebox.showinfo("Info", f"{action_name} finished.\nStatus: {status}\nMessage: {message}")
                        else:
                             messagebox.showinfo("Info", f"{action_name} finished.\nStatus: {status}")
                else:
                    logging.warning(f"Legacy or unexpected return type from {action_name}: {response}")
                    if response not in [None, False]:
                         messagebox.showinfo("Completed", f"{action_name} finished successfully.")
            except queue.Empty:
                if loading_window.winfo_exists():
                    self.root.after(100, check_queue)
            except Exception as e:
                if loading_window.winfo_exists():
                    loading_window.close_window()
                logging.exception(f"Error in loading indicator's check_queue for {action_name}: {e}")
                messagebox.showerror("Error", f"An error occurred while monitoring {action_name}: {e}")

        self.root.after(100, check_queue)

    def setup_keyboard_shortcuts(self):
        self.root.bind('<Control-l>', lambda event: self._safe_call(self.file_handler.load_from_documentpdf, "Load"))
        self.root.bind('<Control-e>', lambda event: self._execute_with_loading(self.file_handler.export_to_pdf, "Export PDF", "Exporting to PDF..."))
        self.root.bind('<Control-w>', lambda event: self._execute_with_loading(self.file_handler.save_to_docx, "Save Word", "Saving to Word document..."))
        self.root.bind('<Control-s>', lambda e: messagebox.showinfo("Not Implemented", "Save functionality (Ctrl+S) is not yet implemented."))
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        logging.info("Keyboard shortcuts set up.")

    def _safe_call(self, func, action_name):
        try:
            func()
            logging.info(f"Action '{action_name}' executed (synchronously).")
        except AttributeError as e:
            logging.error(f"Action '{action_name}' failed: Method not found or attribute missing: {e}")
            messagebox.showerror("Error", f"Could not perform '{action_name}'. Feature might be misconfigured or an object is missing.")
        except FileNotFoundError as e:
            logging.error(f"Action '{action_name}' failed: File not found: {e}")
            messagebox.showerror("File Error", f"File not found during {action_name}:\n{e}")
        except Exception as e:
            logging.exception(f"Error during '{action_name}' action.")
            messagebox.showerror("Error", f"An unexpected error occurred during {action_name}:\n{e}")

    def show_calendar(self):
        # ... (show_calendar method remains the same) ...
        if not HoverCalendar:
            logging.warning("Attempted to open calendar, but HoverCalendar is not available.")
            messagebox.showerror("Error", "Calendar functionality is not available (HoverCalendar library missing).")
            return
        popup = ctk.CTkToplevel(self.root)
        popup.title("Select Date")
        popup_width = min(max(350, int(self.screen_width * 0.25)), 550)
        popup_height = min(max(380, int(self.screen_height * 0.35)), 600)
        self.root.update_idletasks()
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()
        x = main_x + (main_width - popup_width) // 2
        y = main_y + (main_height - popup_height) // 2
        popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")
        popup.transient(self.root)
        popup.grab_set()
        popup.attributes("-topmost", True)
        try:
            cal_font_size = max(10, int(self.base_font_size * 1.1))
            cal = HoverCalendar(popup, font=("Roboto", cal_font_size))
            cal.pack(padx=self.section_pad_x, pady=self.section_pad_y, fill="both", expand=True)
            try:
                current_date_str = self.date_var.get()
                if current_date_str:
                    current_date = datetime.datetime.strptime(current_date_str, "%m/%d/%Y")
                    cal.selection_set(current_date)
            except ValueError:
                logging.warning(f"Could not pre-select date '{current_date_str}' in calendar.")
            def on_date_select():
                selected_date = cal.selection_get()
                if selected_date:
                    self.date_var.set(selected_date.strftime("%m/%d/%Y"))
                    logging.info(f"Date selected from calendar: {self.date_var.get()}")
                else:
                    logging.warning("Calendar closed without selection.")
                popup.destroy()
            confirm_button_width = min(max(100, int(popup_width * 0.3)), 160)
            confirm_button = ctk.CTkButton(
                popup, text="Confirm", command=on_date_select, font=("Roboto", self.button_font_size),
                width=confirm_button_width, corner_radius=8, fg_color=self.BUTTON_FG_COLOR,
                hover_color=self.BUTTON_HOVER_COLOR, text_color=self.BUTTON_TEXT_COLOR,
            )
            confirm_button.pack(pady=(0, self.section_pad_y))
            popup.protocol("WM_DELETE_WINDOW", popup.destroy)
        except Exception as e:
            logging.exception("Failed to create or display the calendar popup.")
            messagebox.showerror("Calendar Error", f"Could not open the calendar: {e}")
            if popup and popup.winfo_exists():
                popup.destroy()

    def _select_logo(self):
        # ... (_select_logo method remains the same) ...
        filetypes = [("Image files", "*.png *.jpg *.jpeg *.gif *.bmp")]
        filepath = filedialog.askopenfilename(
            title="Select Logo Image",
            filetypes=filetypes
        )
        if filepath:
            if os.path.exists(filepath):
                self.variables['logo_path_var'].set(filepath) # Trace will update display
                logging.info(f"Logo selected: {filepath}")
            else:
                messagebox.showerror("Error", f"Selected file does not exist:\n{filepath}")
                logging.error(f"Selected logo file path does not exist: {filepath}")
                # No change to logo_path_var, display will reflect previous state or "No logo selected" via trace.
        else:
            logging.info("Logo selection cancelled.")

    def _remove_logo(self):
        # ... (_remove_logo method remains the same) ...
        self.variables['logo_path_var'].set("") # Trace will update display to "No logo selected"
        logging.info("Logo removed by user.")

    def show_settings(self):
        """Open the settings window."""
        SettingsWindow(self.root, self.settings_manager)

    # --- Method to Show Attachments Window (unchanged signature) ---
    def show_email_attachments_window(self):
        """Opens the window to manage additional email attachments."""
        if not hasattr(self, 'email_sender'):
             messagebox.showerror("Error", "Email functionality is not available.")
             logging.error("Attempted to open attachments window, but email_sender is not initialized.")
             return
        if 'recipient_emails_var' not in self.variables:
             messagebox.showerror("Error", "Internal error: Recipient variable missing.")
             logging.error("Attempted to open attachments window, but recipient_emails_var is missing from self.variables.")
             return

        # Pass the main root window, email_sender instance, AND the recipient_emails_var
        EmailAttachmentsWindow(
            self.root,
            self.email_sender,
            self.variables['recipient_emails_var'] # Pass the actual StringVar
        )
    # --- End Method ---

    def create_widgets(self):
        # Main frame with footer space
        self.main_frame = ctk.CTkFrame(self.root, corner_radius=0, fg_color=self.BG_COLOR)
        self.main_frame.pack(fill="both", expand=True)
        # Configure rows for content_frame and footer_frame
        self.main_frame.grid_rowconfigure(0, weight=1)  # content_frame takes available vertical space
        self.main_frame.grid_rowconfigure(1, weight=0)  # footer_frame has fixed height
        self.main_frame.grid_columnconfigure(0, weight=1)

        # --- MODIFICATION: Use CTkScrollableFrame for content_frame ---
        self.content_frame = ctk.CTkScrollableFrame(self.main_frame, corner_radius=0, fg_color=self.BG_COLOR)
        # --- END MODIFICATION ---
        self.content_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 10)) # Increased top padding for better spacing
        self.content_frame.grid_columnconfigure(0, weight=1) # Ensure content within scrollable frame can expand horizontally

        # Footer (remains unchanged, placed in self.main_frame)
        footer_height = 80
        self.footer_frame = ctk.CTkFrame(self.main_frame, height=footer_height, corner_radius=0, fg_color=self.FOOTER_BG)
        self.footer_frame.grid(row=1, column=0, sticky="ew")
        self.footer_frame.grid_propagate(False)

        # Footer text container (aligned right)
        footer_text_frame = ctk.CTkFrame(self.footer_frame, fg_color="transparent")
        footer_text_frame.pack(side="right", padx=self.base_pad_x)

        # Container for Copyright and GitHub link
        copyright_github_container = ctk.CTkFrame(footer_text_frame, fg_color="transparent")
        copyright_github_container.pack(anchor="e", padx=(0, 20))

        # Copyright text
        compliance_label = ctk.CTkLabel(
            copyright_github_container,
            text="© 2025 All Rights Reserved",
            font=("Roboto", 10), text_color="#FFFFFF", justify="right"
        )
        compliance_label.pack(side="left")

        # GitHub link and logo
        github_url = "https://github.com/JhoramNarsico/Cash-Statement"
        try:
            gh_logo_path = resource_path("github logo.png")
            if os.path.exists(gh_logo_path):
                gh_logo_pil = Image.open(gh_logo_path).convert("RGBA")
                gh_logo_size = (16, 16); gh_logo_pil = gh_logo_pil.resize(gh_logo_size, Image.Resampling.LANCZOS)
                gh_logo_ctk = ctk.CTkImage(light_image=gh_logo_pil, dark_image=gh_logo_pil, size=gh_logo_size)
                gh_logo_label = ctk.CTkLabel(copyright_github_container, image=gh_logo_ctk, text="", cursor="hand2")
                gh_logo_label.pack(side="left", padx=(8, 0))
                gh_logo_label.bind("<Button-1>", lambda e, url=github_url: webbrowser.open_new_tab(url))
            else:
                logging.warning(f"GitHub logo not found: {gh_logo_path}. Displaying text link.")
                github_fallback_label = ctk.CTkLabel(copyright_github_container, text="GitHub", font=("Roboto", 10), text_color="#A9D1F7", cursor="hand2")
                github_fallback_label.pack(side="left", padx=(5, 0))
                github_fallback_label.bind("<Button-1>", lambda e, url=github_url: webbrowser.open_new_tab(url))
        except Exception as e_gh_logo:
            logging.error(f"Error loading GitHub logo: {e_gh_logo}. Displaying text link.")
            github_fallback_label = ctk.CTkLabel(copyright_github_container, text="GitHub", font=("Roboto", 10), text_color="#A9D1F7", cursor="hand2")
            github_fallback_label.pack(side="left", padx=(5, 0))
            github_fallback_label.bind("<Button-1>", lambda e, url=github_url: webbrowser.open_new_tab(url))

        # Footer images
        image1_size = (70, 70); image2_size = (273, 70)
        try:
            footer_image1_var = self.variables.setdefault('footer_image1_var', ctk.StringVar(value=resource_path("chud logo.png")))
            footer_image2_var = self.variables.setdefault('footer_image2_var', ctk.StringVar(value=resource_path("xu logo.png")))
            image1_path = footer_image1_var.get(); image2_path = footer_image2_var.get()
            if os.path.exists(image1_path):
                img1 = Image.open(image1_path); img1 = img1.resize(image1_size, Image.Resampling.LANCZOS)
                ctk_img1 = ctk.CTkImage(light_image=img1, dark_image=img1, size=image1_size)
                img1_label = ctk.CTkLabel(self.footer_frame, image=ctk_img1, text="")
                img1_label.pack(side="left", padx=(20, 10), pady=self.base_pad_y)
                logging.info(f"Footer image 1 loaded: {image1_path}")
            else:
                logging.warning(f"Footer image 1 not found: {image1_path}")
                ctk.CTkLabel(self.footer_frame, text="Img1 N/A", font=("Roboto", self.label_font_size-2), text_color="#FFFFFF").pack(side="left", padx=(20, 10), pady=self.base_pad_y)
            if os.path.exists(image2_path):
                img2 = Image.open(image2_path); img2 = img2.resize(image2_size, Image.Resampling.LANCZOS)
                ctk_img2 = ctk.CTkImage(light_image=img2, dark_image=img2, size=image2_size)
                img2_label = ctk.CTkLabel(self.footer_frame, image=ctk_img2, text="")
                img2_label.pack(side="left", padx=(10, 20), pady=self.base_pad_y)
                logging.info(f"Footer image 2 loaded: {image2_path}")
            else:
                logging.warning(f"Footer image 2 not found: {image2_path}")
                ctk.CTkLabel(self.footer_frame, text="Img2 N/A", font=("Roboto", self.label_font_size-2), text_color="#FFFFFF").pack(side="left", padx=(10, 20), pady=self.base_pad_y)
        except Exception as e:
            logging.exception("Failed to load footer images.")
            ctk.CTkLabel(self.footer_frame, text="Img Err", font=("Roboto", self.label_font_size-2), text_color="#FFFFFF").pack(side="right", padx=self.base_pad_x, pady=self.base_pad_y)

        # --- Subdivide content frame (now the CTkScrollableFrame) ---
        # Note: For CTkScrollableFrame, direct children are packed/gridded into its internal canvas.
        # The grid_columnconfigure and grid_rowconfigure below apply to the layout *within* the scrollable area.
        self.content_frame.grid_columnconfigure(0, weight=1) # Single column for content stack
        self.content_frame.grid_rowconfigure(0, weight=0)  # header_config_frame (row 0)
        self.content_frame.grid_rowconfigure(1, weight=0)  # button_frame (action buttons) (row 1)
        # For CTkScrollableFrame, it manages its own vertical expansion based on content.
        # So, row 2 for columns_frame doesn't need 'weight=1' in the same way,
        # but its content will determine the scrollable height.
        self.content_frame.grid_rowconfigure(2, weight=0) # columns_frame (data sections) (row 2)
        self.content_frame.grid_rowconfigure(3, weight=0)  # names_frame (row 3)


        # --- Header configuration (Address, Logo, Settings/Attachments, Date) ---
        # These are now children of self.content_frame (the CTkScrollableFrame)
        header_config_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        header_config_frame.grid(row=0, column=0, sticky="ew", padx=self.section_pad_x, pady=(0, 0))
        header_config_frame.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="header_group") # 4 columns

        # Address (Column 0)
        address_frame = ctk.CTkFrame(header_config_frame, fg_color="transparent")
        address_frame.grid(row=0, column=0, sticky="ew", padx=(0, self.base_pad_x))
        ctk.CTkLabel(address_frame, text="Header Address:", font=("Roboto", self.label_font_size), text_color=self.TEXT_COLOR, anchor="w").pack(side="top", fill="x", pady=(0, self.base_pad_y // 2))
        address_entry = ctk.CTkEntry(address_frame, textvariable=self.variables['address_var'], font=("Roboto", self.entry_font_size), corner_radius=6, fg_color=self.ENTRY_BG_COLOR, text_color=self.TEXT_COLOR, border_color=self.ENTRY_BORDER_COLOR)
        address_entry.pack(side="top", fill="x")

        # Logo (Column 1)
        logo_frame = ctk.CTkFrame(header_config_frame, fg_color="transparent")
        logo_frame.grid(row=0, column=1, sticky="ew", padx=(self.base_pad_x, self.base_pad_x))
        ctk.CTkLabel(logo_frame, text="Header Logo:", font=("Roboto", self.label_font_size), text_color=self.TEXT_COLOR, anchor="w").pack(side="top", fill="x", pady=(0, self.base_pad_y // 2))
        logo_buttons_frame = ctk.CTkFrame(logo_frame, fg_color="transparent")
        logo_buttons_frame.pack(side="top", fill="x")
        select_logo_button = ctk.CTkButton(logo_buttons_frame, text="Select Logo", font=("Roboto", self.button_font_size), command=self._select_logo, corner_radius=6, fg_color=self.BUTTON_FG_COLOR, hover_color=self.BUTTON_HOVER_COLOR, text_color=self.BUTTON_TEXT_COLOR)
        select_logo_button.pack(side="left", padx=(0, self.base_pad_x // 2))
        remove_logo_button = ctk.CTkButton(logo_buttons_frame, text="Remove Logo", font=("Roboto", self.button_font_size), command=self._remove_logo, corner_radius=6, fg_color=self.REMOVE_BTN_FG_COLOR, hover_color=self.REMOVE_BTN_HOVER_COLOR, text_color=self.BUTTON_TEXT_COLOR)
        remove_logo_button.pack(side="left", padx=(0, self.base_pad_x))
        self.logo_path_display = ctk.CTkLabel(logo_frame, text="No logo selected", font=("Roboto", max(9, self.label_font_size-2)), text_color=self.TEXT_COLOR, anchor="w", wraplength=200)
        self.logo_path_display.pack(side="top", fill="x", pady=(self.base_pad_y // 2, 0))

        # Settings & Attachments (Column 2)
        settings_attach_frame = ctk.CTkFrame(header_config_frame, fg_color="transparent")
        settings_attach_frame.grid(row=0, column=2, sticky="ew", padx=(self.base_pad_x, self.base_pad_x))
        # Label for the whole section (optional)
        ctk.CTkLabel(settings_attach_frame, text="Actions:", font=("Roboto", self.label_font_size), text_color=self.TEXT_COLOR, anchor="w").pack(side="top", fill="x", pady=(0, self.base_pad_y // 2))

        # Subframe to hold buttons side-by-side
        buttons_subframe = ctk.CTkFrame(settings_attach_frame, fg_color="transparent")
        buttons_subframe.pack(side="top", fill="x", anchor="w") # Align buttons left

        # --- Account Settings Button ---
        settings_button_width = min(max(120, int(self.screen_width * 0.08)), 180) # Keep calculated width
        settings_button = ctk.CTkButton(
            buttons_subframe, # Parent is the subframe
            text="Account Settings", # Renamed
            font=("Roboto", self.button_font_size),
            command=self.show_settings,
            corner_radius=6,
            fg_color=self.BUTTON_FG_COLOR,
            hover_color=self.BUTTON_HOVER_COLOR,
            text_color=self.BUTTON_TEXT_COLOR,
            width=settings_button_width
        )
        settings_button.pack(side="left", padx=(0, self.base_pad_x // 2)) # Pack left

        # --- Manage Attachments Button ---
        attachments_button_width = min(max(120, int(self.screen_width * 0.1)), 200) # Slightly wider maybe
        attachments_button = ctk.CTkButton(
            buttons_subframe, # Parent is the subframe
            text="Manage Attachments & Send",
            font=("Roboto", self.button_font_size),
            command=self.show_email_attachments_window,
            corner_radius=6,
            fg_color=self.ATTACH_BTN_FG_COLOR, # Orange color
            hover_color=self.ATTACH_BTN_HOVER_COLOR,
            text_color=self.BUTTON_TEXT_COLOR,
            width=attachments_button_width
        )
        attachments_button.pack(side="left", padx=(self.base_pad_x // 2, 0)) # Pack left, next to settings

        # Date (Column 3)
        date_frame = ctk.CTkFrame(header_config_frame, fg_color="transparent")
        date_frame.grid(row=0, column=3, sticky="ew", padx=(self.base_pad_x, 0))
        ctk.CTkLabel(date_frame, text="Report Date:", font=("Roboto", self.label_font_size), text_color=self.TEXT_COLOR, anchor="w").pack(side="top", fill="x", pady=(0, self.base_pad_y // 2))
        date_button_width = min(max(120, int(self.screen_width * 0.08)), 180)
        date_button = ctk.CTkButton(date_frame, textvariable=self.display_date, font=("Roboto", self.button_font_size), command=self.show_calendar, corner_radius=6, fg_color=self.DATE_BTN_FG, hover_color=self.DATE_BTN_HOVER, text_color=self.DATE_BTN_TEXT, width=date_button_width)
        date_button.pack(side="top", anchor="w") # Kept anchor="w"

        # --- Action buttons (bottom of content) ---
        button_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        button_frame.grid(row=1, column=0, sticky="ew", padx=self.section_pad_x, pady=(self.base_pad_y, self.section_pad_y // 2)) # row 1 for buttons
        buttons_data = [
            ("Load (Ctrl+L)", self.file_handler.load_from_documentpdf, "Load data from DOCX/PDF"),
            ("Clear Fields", self.clear_fields, "Clear all input fields"),
            ("Export PDF (Ctrl+E)", self.file_handler.export_to_pdf, "Export as PDF"),
            ("Save Word (Ctrl+W)", self.file_handler.save_to_docx, "Save as DOCX"),
        ]
        num_buttons = len(buttons_data)
        if num_buttons > 0:
            button_frame.grid_columnconfigure(tuple(range(num_buttons)), weight=1, uniform="button_group")

        for i, (text, command, tooltip_text) in enumerate(buttons_data):
            action_name_simple = text.split('(')[0].strip()
            loading_message = f"{action_name_simple}..."
            fg_color = self.BUTTON_FG_COLOR
            hover_color = self.BUTTON_HOVER_COLOR
            if command in [self.file_handler.export_to_pdf, self.file_handler.save_to_docx]:
                actual_command = lambda f=command, name=action_name_simple, msg=loading_message: self._execute_with_loading(f, name, msg)
            else:
                actual_command = lambda cmd=command, name=action_name_simple: self._safe_call(cmd, name)
            btn = ctk.CTkButton(
                button_frame, text=text,
                command=actual_command,
                font=("Roboto", self.button_font_size), corner_radius=8,
                fg_color=fg_color, hover_color=hover_color,
                text_color=self.BUTTON_TEXT_COLOR, height=int(self.base_font_size * 2.5)
            )
            btn.grid(row=0, column=i, sticky="ew", padx=self.base_pad_x // 4, pady=self.base_pad_y // 4)


        # --- Form sections (middle tables) ---
        self.columns_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.columns_frame.grid(row=2, column=0, sticky="ew", padx=self.section_pad_x // 2, pady=(self.base_pad_y, self.section_pad_y)) # row 2 for columns
        num_data_cols = 5
        # The min_column_width for data columns is now dynamically tied to self.min_input_width
        min_col_width = int(self.min_input_width * 1.7) # Ensure columns are wide enough
        for i in range(num_data_cols):
            self.columns_frame.grid_columnconfigure(i, weight=1, uniform="data_cols", minsize=min_col_width)

        self.beg_frame = self._create_section_frame("Beginning Cash Balances")
        self.inflow_frame = self._create_section_frame("Cash Inflows")
        self.outflow_frame = self._create_section_frame("Cash Outflows")
        self.end_frame = self._create_section_frame("Ending Cash Balances (Calculated)")
        self.totals_frame = self._create_section_frame("Totals (Calculated)")
        self.beg_frame.grid(row=0, column=0, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.inflow_frame.grid(row=0, column=1, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.outflow_frame.grid(row=0, column=2, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.end_frame.grid(row=0, column=3, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.totals_frame.grid(row=0, column=4, sticky="nsew", padx=self.base_pad_x//2, pady=self.base_pad_y)
        self.populate_columns()


        # --- Names section ---
        names_frame = ctk.CTkFrame(
            self.content_frame, corner_radius=8, fg_color=self.FRAME_COLOR,
            border_width=1, border_color=self.BORDER_COLOR,
        )
        names_frame.grid( row=3, column=0, sticky="ew", padx=self.section_pad_x, pady=(0, self.base_pad_y) ) # Added some bottom padding
        name_fields_data = [
            ("Prepared by (Treasurer):", 'prepared_by_var', "Name of HOA Treasurer"),
            ("Noted by (President):", 'noted_by_var_1', "Name of HOA President"),
            ("Noted by (CHUDD HCD-CORDS):", 'noted_by_var_2', "Name of CHUDD HCD-CORDS rep"),
            ("Checked by (Auditor):", 'checked_by_var', "Name of HOA Auditor")
        ]
        num_name_fields = len(name_fields_data)
        # The min_name_col_width is also dynamically tied to self.min_input_width
        min_name_col_width = int(self.min_input_width * 1.5) # Ensure name columns are wide enough
        for i in range(num_name_fields):
            names_frame.grid_columnconfigure(i, weight=1, uniform="name_group_adjusted", minsize=min_name_col_width)
        for i, (label_text, var_key, tooltip_text) in enumerate(name_fields_data):
            frame = ctk.CTkFrame(names_frame, fg_color="transparent")
            frame.grid(row=0, column=i, sticky="nsew", padx=self.base_pad_x, pady=self.base_pad_y)
            frame.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(frame, text=label_text, font=("Roboto", self.label_font_size), text_color=self.TEXT_COLOR, anchor="w").grid(row=0, column=0, sticky="ew", pady=(0, self.base_pad_y // 2))
            entry = ctk.CTkEntry(frame, textvariable=self.variables[var_key], font=("Roboto", self.entry_font_size), corner_radius=6, fg_color=self.ENTRY_BG_COLOR, text_color=self.TEXT_COLOR, border_color=self.ENTRY_BORDER_COLOR)
            entry.grid(row=1, column=0, sticky="ew")

        # Bind configure to main_frame still, as this controls the overall window size changes
        self.main_frame.bind("<Configure>", self.debounce_layout, add="+")


    def _create_section_frame(self, title):
        # This frame is a child of self.columns_frame, which is a child of self.content_frame (scrollable)
        frame = ctk.CTkFrame(self.columns_frame, corner_radius=8, fg_color=self.FRAME_COLOR, border_width=1, border_color=self.BORDER_COLOR)
        ctk.CTkLabel(
            frame, text=title, font=("Roboto", self.title_font_size, "bold"),
            text_color=self.TEXT_COLOR, anchor="w"
        ).pack(fill="x", padx=self.base_pad_x * 1.5, pady=(self.base_pad_y * 1.5, self.base_pad_y))
        content_frame = ctk.CTkFrame(frame, fg_color="transparent") # This content_frame is local to the section
        content_frame.pack(fill="both", expand=True, padx=self.base_pad_x, pady=(0, self.base_pad_y * 1.5))
        return frame


    def _create_entry_pair(self, parent_frame, label_text, var, is_disabled=False, tooltip_text=None, input_width=None, is_numeric=False):
        # ... (Remains the same from previous step, using sticky="e" for entry) ...
        item_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
        item_frame.pack(fill="x", padx=self.base_pad_x // 2, pady=self.base_pad_y // 2)
        item_frame.grid_columnconfigure(0, weight=0) # Column for Label
        item_frame.grid_columnconfigure(1, weight=1) # Column for Entry (this column will expand)
        
        ctk.CTkLabel(
            item_frame, text=label_text, font=("Roboto", self.label_font_size),
            text_color=self.TEXT_COLOR, anchor="w"
        ).grid(row=0, column=0, sticky="w", padx=(0, self.base_pad_x))
        
        preferred_entry_width = input_width or self.min_input_width
        entry_config = {
            "textvariable": var, "width": preferred_entry_width,
            "font": ("Roboto", self.entry_font_size), "corner_radius": 6,
            "fg_color": self.DISABLED_BG_COLOR if is_disabled else self.ENTRY_BG_COLOR,
            "text_color": self.TEXT_COLOR, "border_color": self.ENTRY_BORDER_COLOR,
            "state": "disabled" if is_disabled else "normal", "justify": "right"
        }
        if is_numeric and not is_disabled:
            entry_config["validate"] = "key"; entry_config["validatecommand"] = self.vcmd_numeric
        
        entry = ctk.CTkEntry(item_frame, **entry_config)
        entry.grid(row=0, column=1, sticky="e") # Aligned to the right
        
        if not is_disabled and is_numeric:
            if hasattr(self.calculator, 'format_entry'):
                try: self.calculator.format_entry(var, entry)
                except Exception as e: logging.exception(f"Error applying format_entry to '{label_text}': {e}")
            else: logging.warning(f"Calculator object missing 'format_entry' method for '{label_text}'.")


    def populate_columns(self):
        # ... (Remains the same, but will use the new self.min_input_width and scaled font) ...
        beg_content = self.beg_frame.winfo_children()[1]
        beg_items = [("Cash in Bank:", 'cash_bank_beg', "Starting bank balance"), ("Cash on Hand:", 'cash_hand_beg', "Starting physical cash")]
        for label, var_key, tooltip in beg_items: self._create_entry_pair(beg_content, label, self.variables[var_key], tooltip_text=tooltip, is_numeric=True)
        
        inflow_content = self.inflow_frame.winfo_children()[1]
        inflow_items = [ ("Monthly dues collected:", 'monthly_dues'), ("Certifications issued:", 'certifications'), ("Membership fee:", 'membership_fee'), ("Vehicle stickers:", 'vehicle_stickers'), ("Rentals:", 'rentals'), ("Solicitations/Donations:", 'solicitations'), ("Interest Income:", 'interest_income'), ("Livelihood Fee:", 'livelihood_fee'), ("Others:", 'inflows_others', "Other income sources")]
        for label, var_key, *tooltip_arg in inflow_items: self._create_entry_pair( inflow_content, label, self.variables[var_key], tooltip_text=tooltip_arg[0] if tooltip_arg else None, input_width=int(self.min_input_width / 1.5), is_numeric=True ) 
        
        outflow_content = self.outflow_frame.winfo_children()[1]
        outflow_items = [ ("Snacks/Meals:", 'snacks_meals'), ("Transportation:", 'transportation'), ("Office supplies:", 'office_supplies'), ("Printing/Photocopy:", 'printing'), ("Labor:", 'labor'), ("Billboard expense:", 'billboard'), ("Cleaning charges:", 'cleaning'), ("Misc expenses:", 'misc_expenses'), ("Federation fee:", 'federation_fee'), ("Uniforms:", 'uniforms'), ("BOD Mtg:", 'bod_mtg', "Board meeting expenses"), ("General Assembly:", 'general_assembly', "Assembly expenses"), ("Cash Deposit:", 'cash_deposit', "Cash moved hand to bank"), ("Withholding tax:", 'withholding_tax'), ("Refund:", 'refund'), ("Others:", 'outflows_others', "Other expenses")]
        for label, var_key, *tooltip_arg in outflow_items: self._create_entry_pair( outflow_content, label, self.variables[var_key], tooltip_text=tooltip_arg[0] if tooltip_arg else None, input_width=int(self.min_input_width / 1.5), is_numeric=True )
        
        end_content = self.end_frame.winfo_children()[1]
        end_items = [("Cash in Bank:", 'ending_cash_bank', "Calculated ending bank balance"), ("Cash on Hand:", 'ending_cash_hand', "Calculated ending cash on hand")]
        for label, var_key, tooltip in end_items: self._create_entry_pair(end_content, label, self.variables[var_key], is_disabled=True, tooltip_text=tooltip, is_numeric=True)
        
        total_content = self.totals_frame.winfo_children()[1]
        total_items = [("Total Receipts:", 'total_receipts', "Calculated total inflows"), ("Total Outflows:", 'cash_outflows', "Calculated total outflows"), ("Ending Balance:", 'ending_cash', "Calculated total ending cash")]
        for label, var_key, tooltip in total_items: self._create_entry_pair(total_content, label, self.variables[var_key], is_disabled=True, tooltip_text=tooltip, is_numeric=True)
        logging.info("GUI columns populated into fixed horizontal layout.")

    def debounce_layout(self, event=None):
        # ... (Remains the same) ...
        if event and event.widget != self.main_frame: return
        if self.debounce_id: self.root.after_cancel(self.debounce_id)
        self.debounce_id = self.root.after(self.layout_debounce_delay_ms, self.update_layout)

    def update_layout(self):
        # When self.content_frame is a CTkScrollableFrame, its internal scrollbars
        # will adjust based on its content size vs its allocated size.
        # This method might not need to do as much explicit recalculation for elements
        # *within* the scrollable frame, but updating idletasks is still good.
        if self.main_frame.winfo_exists():
            self.main_frame.update_idletasks()
        if self.content_frame.winfo_exists():
            self.content_frame.update_idletasks() # For the scrollable frame itself
        logging.debug("Layout updated/recalculated due to window configure event.")
        self.debounce_id = None


    def clear_fields(self):
        # ... (Remains the same - already updated) ...
        if not self.variables:
            logging.error("Cannot clear fields: variables dictionary missing.")
            messagebox.showerror("Error", "Internal error: Cannot access data fields.")
            return
        if not messagebox.askyesno("Confirm Clear", "Clear all input and calculated fields?\n(Logo, Address, Date, and Signatories will remain)"):
            return
        cleared_count = 0
        fields_to_keep = {
            'logo_path_var', 'address_var', 'date_var', 'display_date',
            'recipient_emails_var',
            'prepared_by_var', 'noted_by_var_1', 'noted_by_var_2', 'checked_by_var', 'title_var',
            'footer_image1_var', 'footer_image2_var'
        }
        for key, var in self.variables.items():
            if key not in fields_to_keep and isinstance(var, ctk.StringVar):
                try:
                    var.set("")
                    cleared_count += 1
                except tkinter.TclError as e:
                    logging.warning(f"Could not clear variable '{key}' (might be destroyed): {e}")
        if 'recipient_emails_var' in self.variables:
             self.variables['recipient_emails_var'].set("")
             logging.info("Cleared recipient email variable.")
        if hasattr(self, 'email_sender') and hasattr(self.email_sender, 'attachments'):
            if self.email_sender.attachments:
                self.email_sender.attachments.clear()
                logging.info("Cleared additional email attachments list.")
        messagebox.showinfo("Success", "Cash flow data fields, recipient list, and additional attachments have been cleared.")
        try:
            if hasattr(self.calculator, 'calculate_totals'):
                self.calculator.calculate_totals()
            else:
                logging.warning("Calculator has no 'calculate_totals' method to call after clearing.")
        except Exception as e:
            logging.exception("Error recalculating totals after clearing fields.")

# --- Example usage (if running gui_components.py directly) ---
if __name__ == '__main__':
    # ... (Mocks remain the same) ...
    class MockCalculator:
        def format_entry(self, var, entry): print(f"Mock format_entry called for var linked to entry {entry}")
        def calculate_totals(self): print("Mock calculate_totals called")

    class MockFileHandler:
        def load_from_documentpdf(self): messagebox.showinfo("Mock", "Load called")
        def export_to_pdf(self):
            time.sleep(1); return {"status": "success", "message": "Mock PDF Exported!"}
        def save_to_docx(self):
            time.sleep(1); return {"status": "success", "message": "Mock Word Saved!"}

    class MockEmailSender:
        def __init__(self, settings_manager, recipient_emails_var, file_handler):
             self.settings_manager = settings_manager
             self.recipient_emails_var = recipient_emails_var
             self.file_handler = file_handler
             self.attachments = []
        def send_email(self):
            time.sleep(1.5)
            recipients = self.recipient_emails_var.get()
            print(f"Mock Send Email called for recipients: {recipients} with attachments: {self.attachments}")
            if not recipients: return {"status": "error", "message": "Mock Error: No recipients."}
            # Simulate no attachments error if needed for testing EmailAttachmentsWindow
            # if not self.attachments: return {"status": "error", "message": "Mock Error: No attachments."}
            return {"status": "success", "message": f"Mock Email Sent with {len(self.attachments)} extra files!"}

    from setting import SettingsManager
    root = ctk.CTk()
    root.title("HOA Cash Flow (Test)")
    root.geometry("1200x750") # Start with a larger size for testing
    # root.geometry("800x500") # Test with smaller size to see scrollbars
    root.resizable(True, True)
    root.minsize(700, 500) # Reduced minsize to test scrolling earlier

    settings_manager = SettingsManager()
    variables = {}
    title_var = ctk.StringVar(value="HOA Cash Flow Statement")
    date_var = ctk.StringVar(value=datetime.date.today().strftime("%m/%d/%Y"))
    display_date = ctk.StringVar()
    calculator = MockCalculator()
    file_handler = MockFileHandler()

    # Instantiate GUI first
    gui = GUIComponents(root, variables, title_var, date_var, display_date, calculator, file_handler, None, settings_manager)

    # Instantiate mock email_sender after GUIComponents populated variables
    if 'recipient_emails_var' in gui.variables:
        email_sender = MockEmailSender(settings_manager, gui.variables['recipient_emails_var'], file_handler)
        gui.email_sender = email_sender # Assign back to gui
    else:
        messagebox.showerror("Test Init Error", "recipient_emails_var was not initialized!")
        root.destroy(); exit()

    root.after(200, gui.update_layout)
    root.mainloop()

# --- END OF FILE gui_components.py ---