# --- START OF FILE cash_flow_calculator.py ---

from decimal import Decimal
import customtkinter as ctk

class CashFlowCalculator:
    def __init__(self, variables):
        self.variables = variables
        self.input_vars = [
            variables['cash_bank_beg'], variables['cash_hand_beg'], variables['monthly_dues'],
            variables['certifications'], variables['membership_fee'], variables['vehicle_stickers'],
            variables['rentals'], variables['solicitations'], variables['interest_income'],
            variables['livelihood_fee'], variables['withdrawal_from_bank'], variables['inflows_others'], # ADDED withdrawal_from_bank
            variables['snacks_meals'], variables['transportation'], variables['office_supplies'],
            variables['printing'], variables['labor'], variables['billboard'], variables['cleaning'],
            variables['misc_expenses'], variables['federation_fee'], variables['uniforms'],
            variables['bod_mtg'], variables['general_assembly'], variables['cash_deposit'],
            variables['withholding_tax'], variables['refund'], variables['outflows_others']
        ]
        for var in self.input_vars:
            var.trace_add('write', lambda *args: self.calculate_totals())

    def safe_decimal(self, var):
        val = var.get().strip()
        if not val:
            return Decimal("0")
        try:
            val = val.replace(",", "")
            return Decimal(val)
        except:
            return Decimal("0")

    def calculate_totals(self):
        try:
            inflow_total = sum([
                self.safe_decimal(self.variables['monthly_dues']),
                self.safe_decimal(self.variables['certifications']),
                self.safe_decimal(self.variables['membership_fee']),
                self.safe_decimal(self.variables['vehicle_stickers']),
                self.safe_decimal(self.variables['rentals']),
                self.safe_decimal(self.variables['solicitations']),
                self.safe_decimal(self.variables['interest_income']),
                self.safe_decimal(self.variables['livelihood_fee']),
                self.safe_decimal(self.variables['withdrawal_from_bank']), # ADDED
                self.safe_decimal(self.variables['inflows_others'])
            ])

            outflow_total = sum([
                self.safe_decimal(self.variables['snacks_meals']),
                self.safe_decimal(self.variables['transportation']),
                self.safe_decimal(self.variables['office_supplies']),
                self.safe_decimal(self.variables['printing']),
                self.safe_decimal(self.variables['labor']),
                self.safe_decimal(self.variables['billboard']),
                self.safe_decimal(self.variables['cleaning']),
                self.safe_decimal(self.variables['misc_expenses']),
                self.safe_decimal(self.variables['federation_fee']),
                self.safe_decimal(self.variables['uniforms']),
                self.safe_decimal(self.variables['bod_mtg']),
                self.safe_decimal(self.variables['general_assembly']),
                self.safe_decimal(self.variables['cash_deposit']), # cash_deposit is part of total outflows
                self.safe_decimal(self.variables['withholding_tax']),
                self.safe_decimal(self.variables['refund']),
                self.safe_decimal(self.variables['outflows_others'])
            ])

            beginning_total = self.safe_decimal(self.variables['cash_bank_beg']) + self.safe_decimal(self.variables['cash_hand_beg'])
            ending_balance = beginning_total + inflow_total - outflow_total

            self.variables['total_receipts'].set(f"{inflow_total:,.2f}")
            self.variables['cash_outflows'].set(f"{outflow_total:,.2f}")
            self.variables['ending_cash'].set(f"{ending_balance:,.2f}")

            # --- MODIFIED CALCULATION for ending_cash_bank and ending_cash_hand ---
            ending_cash_bank_calculated = (
                self.safe_decimal(self.variables['cash_bank_beg']) +
                self.safe_decimal(self.variables['cash_deposit']) -      # "Cash deposit must plus Cash in bank"
                self.safe_decimal(self.variables['withdrawal_from_bank']) # "Withdrawal must minus Cash in bank"
            )
            self.variables['ending_cash_bank'].set(f"{ending_cash_bank_calculated:,.2f}")

            ending_cash_hand_calculated = ending_balance - ending_cash_bank_calculated
            self.variables['ending_cash_hand'].set(f"{ending_cash_hand_calculated:,.2f}")
            # --- END OF MODIFICATION ---

        except Exception:
            self.variables['total_receipts'].set("")
            self.variables['cash_outflows'].set("")
            self.variables['ending_cash'].set("")
            self.variables['ending_cash_bank'].set("")
            self.variables['ending_cash_hand'].set("")

    def format_entry(self, var, entry_widget):
        def on_change(*args):
            value = var.get()
            if value:
                try:
                    formatted = f"{Decimal(value.replace(',', '')):,.2f}"
                    if formatted != value:
                        var.set(formatted)
                except:
                    pass
        var.trace_add('write', on_change)
        entry_widget.configure(justify="right")

# --- END OF FILE cash_flow_calculator.py ---