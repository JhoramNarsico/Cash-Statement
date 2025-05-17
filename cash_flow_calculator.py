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
            variables['livelihood_fee'], variables['withdrawal_from_bank'], variables['inflows_others'],
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
            # Corrected: Inflows are external receipts into the business
            inflow_total = sum([
                self.safe_decimal(self.variables['monthly_dues']),
                self.safe_decimal(self.variables['certifications']),
                self.safe_decimal(self.variables['membership_fee']),
                self.safe_decimal(self.variables['vehicle_stickers']),
                self.safe_decimal(self.variables['rentals']),
                self.safe_decimal(self.variables['solicitations']),
                self.safe_decimal(self.variables['interest_income']),
                self.safe_decimal(self.variables['livelihood_fee']),
                # self.safe_decimal(self.variables['withdrawal_from_bank']), # REMOVED from total inflows
                self.safe_decimal(self.variables['inflows_others'])
            ])

            # Corrected: Outflows are external disbursements from the business
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
                # self.safe_decimal(self.variables['cash_deposit']), # REMOVED from total outflows
                self.safe_decimal(self.variables['withholding_tax']),
                self.safe_decimal(self.variables['refund']),
                self.safe_decimal(self.variables['outflows_others'])
            ])

            beginning_total = self.safe_decimal(self.variables['cash_bank_beg']) + self.safe_decimal(self.variables['cash_hand_beg'])
            
            # This ending_balance is now the TRUE total cash of the entity
            ending_balance = beginning_total + inflow_total - outflow_total

            self.variables['total_receipts'].set(f"{inflow_total:,.2f}")
            self.variables['cash_outflows'].set(f"{outflow_total:,.2f}")
            self.variables['ending_cash'].set(f"{ending_balance:,.2f}") # This is total ending cash

            # This calculation for ending_cash_bank IS correct based on its specific purpose:
            # to track how deposits and withdrawals affect the bank balance.
            ending_cash_bank_calculated = (
                self.safe_decimal(self.variables['cash_bank_beg']) +
                self.safe_decimal(self.variables['cash_deposit']) -      # Money moved TO bank (internal transfer)
                self.safe_decimal(self.variables['withdrawal_from_bank']) # Money moved FROM bank (internal transfer)
            )
            self.variables['ending_cash_bank'].set(f"{ending_cash_bank_calculated:,.2f}")

            # Ending cash on hand is the total cash less what's in the bank.
            # This remains the most reliable way to calculate it.
            ending_cash_hand_calculated = ending_balance - ending_cash_bank_calculated
            self.variables['ending_cash_hand'].set(f"{ending_cash_hand_calculated:,.2f}")

        except Exception as e: # It's good practice to catch specific exceptions or log the error
            # print(f"Error during calculation: {e}") # Optional: for debugging
            self.variables['total_receipts'].set("")
            self.variables['cash_outflows'].set("")
            self.variables['ending_cash'].set("")
            self.variables['ending_cash_bank'].set("")
            self.variables['ending_cash_hand'].set("")

    def format_entry(self, var, entry_widget):
        def on_change(*args):
            # Store cursor position
            entry = entry_widget
            cursor_pos = entry.index(ctk.INSERT)
            original_len = len(var.get())

            value = var.get()
            if value:
                try:
                    # Remove existing commas for calculation
                    numeric_value_str = value.replace(',', '')
                    # Check if it's just a negative sign or ends with a decimal point
                    if numeric_value_str == "-" or numeric_value_str.endswith('.'):
                        # Don't format yet, allow user to continue typing
                        pass
                    elif numeric_value_str.count('.') > 1: # Prevent multiple decimal points
                        # Revert to a state before multiple decimal points if possible or just remove last char
                        var.set(value[:-1])
                        entry.icursor(cursor_pos -1 if cursor_pos > 0 else 0)
                        return

                    else:
                        # Handle numbers that might be in scientific notation temporarily if very large/small
                        # though Decimal usually handles this.
                        # Ensure we are only formatting if it's a valid number that can be Decimal
                        dec_val = Decimal(numeric_value_str)
                        
                        # Preserve decimal places during typing
                        if '.' in numeric_value_str:
                            integer_part, decimal_part = numeric_value_str.split('.', 1)
                            formatted_integer = f"{Decimal(integer_part):,}" if integer_part else "0" # Handle case like ".5"
                            if not decimal_part: # e.g. "123."
                                formatted = formatted_integer + "."
                            else: # e.g. "123.45"
                                formatted = f"{formatted_integer}.{decimal_part}"
                        else: # No decimal point yet
                            formatted = f"{dec_val:,}" # Add .00 only on blur or finalization

                        # Only update if the formatted value is different to prevent recursion
                        # and to avoid issues if the user is typing a decimal point
                        if formatted != value:
                            var.set(formatted)
                            
                            # Adjust cursor position after formatting
                            new_len = len(formatted)
                            len_diff = new_len - original_len
                            new_cursor_pos = cursor_pos + len_diff
                            
                            # Basic adjustment for comma additions/removals
                            # More sophisticated cursor management might be needed for all edge cases
                            num_commas_orig = value[:cursor_pos].count(',')
                            num_commas_new = formatted[:new_cursor_pos].count(',')
                            cursor_adjustment = num_commas_new - num_commas_orig
                            
                            final_cursor_pos = cursor_pos + len_diff - cursor_adjustment
                            if final_cursor_pos < 0: final_cursor_pos = 0
                            if final_cursor_pos > len(formatted): final_cursor_pos = len(formatted)
                            
                            entry.icursor(final_cursor_pos)


                except Exception as e:
                    # print(f"Formatting error: {e} with value: {value}") # for debugging
                    # Potentially revert to a safe value or do nothing,
                    # to avoid clearing user input aggressively
                    pass

        def on_focus_out(event):
            value_str = var.get()
            if value_str:
                try:
                    # Remove commas for conversion, then format to 2 decimal places
                    cleaned_value = value_str.replace(',', '')
                    if cleaned_value == "-" or cleaned_value == ".": # Handle incomplete entries
                        var.set("0.00")
                        return
                    if not cleaned_value: # Handle empty string after cleaning
                        var.set("0.00")
                        return

                    # Convert to Decimal and format
                    dec_value = Decimal(cleaned_value)
                    formatted_value = f"{dec_value:,.2f}"
                    if formatted_value != value_str:
                        var.set(formatted_value)
                except Exception:
                     # If it's not a valid number, set to 0.00 or leave as is, or clear
                    # For now, let's try to set to 0.00 if it's not a parseable number
                    var.set("0.00")


        var.trace_add('write', on_change)
        entry_widget.bind("<FocusOut>", on_focus_out)
        entry_widget.configure(justify="right")

# --- END OF FILE cash_flow_calculator.py ----