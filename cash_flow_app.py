import tkinter as tk
import customtkinter as ctk
import datetime
from setting import SettingsManager

class CashFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cash Flow Statement Generator with Email")
        self.root.geometry("1920x1080")

        # Initialize settings
        self.settings_manager = SettingsManager()

        # Initialize variables
        self.variables = {
            'recipient_emails_var': tk.StringVar(),
            'title_var': tk.StringVar(value="Statement Of Cash Flows"),
            'date_var': tk.StringVar(value=datetime.datetime.now().strftime("%m/%d/%Y")),
            'display_date': tk.StringVar(value=datetime.datetime.now().strftime("%b %d, %Y")),
            'prepared_by_var': tk.StringVar(),
            'noted_by_var_1': tk.StringVar(),
            'noted_by_var_2': tk.StringVar(),
            'checked_by_var': tk.StringVar(),
            'cash_bank_beg': tk.StringVar(),
            'cash_hand_beg': tk.StringVar(),
            'monthly_dues': tk.StringVar(),
            'certifications': tk.StringVar(),
            'membership_fee': tk.StringVar(),
            'vehicle_stickers': tk.StringVar(),
            'rentals': tk.StringVar(),
            'solicitations': tk.StringVar(),
            'interest_income': tk.StringVar(),
            'livelihood_fee': tk.StringVar(),
            'inflows_others': tk.StringVar(),
            'total_receipts': tk.StringVar(),
            'cash_outflows': tk.StringVar(),
            'snacks_meals': tk.StringVar(),
            'transportation': tk.StringVar(),
            'office_supplies': tk.StringVar(),
            'printing': tk.StringVar(),
            'labor': tk.StringVar(),
            'billboard': tk.StringVar(),
            'cleaning': tk.StringVar(),
            'misc_expenses': tk.StringVar(),
            'federation_fee': tk.StringVar(),
            'uniforms': tk.StringVar(),
            'bod_mtg': tk.StringVar(),
            'general_assembly': tk.StringVar(),
            'cash_deposit': tk.StringVar(),
            'withholding_tax': tk.StringVar(),
            'refund': tk.StringVar(),
            'outflows_others': tk.StringVar(),
            'ending_cash': tk.StringVar(),
            'ending_cash_bank': tk.StringVar(),
            'ending_cash_hand': tk.StringVar(),
            'logo_path_var': tk.StringVar(),
            'address_var': tk.StringVar(value="Default Address - Change Me"),
        }

        # Initialize components
        from cash_flow_calculator import CashFlowCalculator
        from file_handler import FileHandler
        from email_sender import EmailSender
        from gui_components import GUIComponents

        self.calculator = CashFlowCalculator(self.variables)
        self.file_handler = FileHandler(
            self.variables,
            self.variables['title_var'],
            self.variables['date_var'],
            self.variables['logo_path_var'],
            self.variables['address_var'],
            self.variables['prepared_by_var'],
            self.variables['noted_by_var_1'],
            self.variables['noted_by_var_2'],
            self.variables['checked_by_var']
        )
        self.email_sender = EmailSender(
            sender_email=self.settings_manager.get_setting("sender_email"),
            sender_password=self.settings_manager.get_setting("sender_password"),
            recipient_emails_var=self.variables['recipient_emails_var'],
            file_handler=self.file_handler
        )
        self.gui = GUIComponents(
            self.root,
            self.variables,
            self.variables['title_var'],
            self.variables['date_var'],
            self.variables['display_date'],
            self.calculator,
            self.file_handler,
            self.email_sender,
            self.settings_manager
        )

      