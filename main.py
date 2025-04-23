import customtkinter as ctk
from cash_flow_app import CashFlowApp

if __name__ == "__main__":
    root = ctk.CTk()
    app = CashFlowApp(root)
    root.mainloop()