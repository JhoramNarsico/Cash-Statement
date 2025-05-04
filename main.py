import customtkinter as ctk
from login import LoginWindow
from setting import SettingsManager

def start_main_app(root, settings_manager):
    from cash_flow_app import CashFlowApp
    app = CashFlowApp(root)

if __name__ == "__main__":
    root = ctk.CTk()
    settings_manager = SettingsManager()
    login_window = LoginWindow(root, lambda: start_main_app(root, settings_manager), settings_manager)
    root.mainloop()