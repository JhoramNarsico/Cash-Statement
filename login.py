import customtkinter as ctk
from tkinter import messagebox

class LoginWindow:
    """
    A login window that requires a username and password to access the main application.
    Uses credentials from SettingsManager for authentication, defaulting to user123/chud123.
    """
    def __init__(self, root, on_success, settings_manager):
        """
        Initializes the login window.
        
        Args:
            root: The customtkinter root window.
            on_success: Callback function to call when login is successful.
            settings_manager: SettingsManager instance for accessing credentials.
        """
        self.root = root
        self.on_success = on_success
        self.settings_manager = settings_manager

        # Set window properties
        self.root.title("Login - HOA Cash Flow")
        self.root.geometry("400x300")
        self.root.resizable(False, False)

        # Center the window
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 400) // 2
        y = (screen_height - 300) // 2
        self.root.geometry(f"400x300+{x}+{y}")

        # Create main frame
        self.main_frame = ctk.CTkFrame(self.root, corner_radius=10, fg_color="#F5F5F5")
        self.main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # Title
        ctk.CTkLabel(
            self.main_frame,
            text="Login to HOA Cash Flow",
            font=("Roboto", 16, "bold"),
            text_color="#333333"
        ).pack(pady=(20, 10))

        # Username field
        ctk.CTkLabel(
            self.main_frame,
            text="Username:",
            font=("Roboto", 12),
            text_color="#333333"
        ).pack(pady=(10, 2))
        self.username_entry = ctk.CTkEntry(
            self.main_frame,
            placeholder_text="Enter username",
            font=("Roboto", 12),
            width=200,
            corner_radius=6,
            fg_color="#FAFAFA",
            text_color="#333333",
            border_color="#B0BEC5"
        )
        self.username_entry.pack(pady=(0, 10))

        # Password field
        ctk.CTkLabel(
            self.main_frame,
            text="Password:",
            font=("Roboto", 12),
            text_color="#333333"
        ).pack(pady=(0, 2))
        self.password_entry = ctk.CTkEntry(
            self.main_frame,
            placeholder_text="Enter password",
            show="*",
            font=("Roboto", 12),
            width=200,
            corner_radius=6,
            fg_color="#FAFAFA",
            text_color="#333333",
            border_color="#B0BEC5"
        )
        self.password_entry.pack(pady=(0, 20))

        # Login button
        self.login_button = ctk.CTkButton(
            self.main_frame,
            text="Login",
            command=self.validate_login,
            font=("Roboto", 12),
            corner_radius=8,
            fg_color="#2196F3",
            hover_color="#1976D2",
            text_color="#FFFFFF",
            height=35
        )
        self.login_button.pack(pady=10)

        # Bind Enter key to login
        self.root.bind("<Return>", lambda event: self.validate_login())

    def validate_login(self):
        """Validates the entered username and password against settings credentials."""
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        expected_username = self.settings_manager.get_setting("username")
        expected_password = self.settings_manager.get_setting("password")

        if username == expected_username and password == expected_password:
            self.main_frame.destroy()
            self.root.unbind("<Return>")
            self.on_success()
        else:
            messagebox.showerror(
                "Login Failed",
                "Incorrect username or password. Please try again."
            )
            self.username_entry.delete(0, "end")
            self.password_entry.delete(0, "end")
            self.username_entry.focus()