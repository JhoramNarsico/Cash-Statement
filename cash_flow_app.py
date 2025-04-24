import customtkinter as ctk
import datetime

class CashFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cash Flow Statement Generator with Email")
        self.root.geometry("800x700")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Initialize variables
        self.variables = {
            'recipient_emails_var': ctk.StringVar(),
            'title_var': ctk.StringVar(value="Statement Of Cash Flows"),
            'date_var': ctk.StringVar(value=datetime.datetime.now().strftime("%m/%d/%Y")),
            'display_date': ctk.StringVar(value=datetime.datetime.now().strftime("%b %d, %Y")),
            'prepared_by_var': ctk.StringVar(),
            'noted_by_var': ctk.StringVar(),
            'checked_by_var': ctk.StringVar(),  # New variable for Checked by
            'cash_bank_beg': ctk.StringVar(),
            'cash_hand_beg': ctk.StringVar(),
            'monthly_dues': ctk.StringVar(),
            'certifications': ctk.StringVar(),
            'membership_fee': ctk.StringVar(),
            'vehicle_stickers': ctk.StringVar(),
            'rentals': ctk.StringVar(),
            'solicitations': ctk.StringVar(),
            'interest_income': ctk.StringVar(),
            'livelihood_fee': ctk.StringVar(),
            'inflows_others': ctk.StringVar(),
            'total_receipts': ctk.StringVar(),
            'cash_outflows': ctk.StringVar(),
            'snacks_meals': ctk.StringVar(),
            'transportation': ctk.StringVar(),
            'office_supplies': ctk.StringVar(),
            'printing': ctk.StringVar(),
            'labor': ctk.StringVar(),
            'billboard': ctk.StringVar(),
            'cleaning': ctk.StringVar(),
            'misc_expenses': ctk.StringVar(),
            'federation_fee': ctk.StringVar(),
            'uniforms': ctk.StringVar(),
            'bod_mtg': ctk.StringVar(),
            'general_assembly': ctk.StringVar(),
            'cash_deposit': ctk.StringVar(),
            'withholding_tax': ctk.StringVar(),
            'refund': ctk.StringVar(),
            'outflows_others': ctk.StringVar(),
            'ending_cash': ctk.StringVar(),
            'ending_cash_bank': ctk.StringVar(),
            'ending_cash_hand': ctk.StringVar()
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
            self.variables['prepared_by_var'],
            self.variables['noted_by_var'],
            self.variables['checked_by_var']  # Pass new variable
        )
        self.email_sender = EmailSender(
            sender_email="chuddcdo@gmail.com",
            sender_password="jfyb eoog ukxr hhiq",
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
            self.email_sender
        )