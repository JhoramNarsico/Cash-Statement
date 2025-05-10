import customtkinter as ctk
from tkinter import messagebox

class LoginWindow:
    """Login window for HOA Cash Flow system."""

    def __init__(self, root, on_success, settings_manager):
        self.root = root
        self.on_success = on_success
        self.settings_manager = settings_manager

        self._setup_window()
        self._create_widgets()

    def _setup_window(self):
        self.root.title("HOA Cash Flow - Login")
        self.root.geometry("420x430")
        self.root.resizable(False, False)

        # Center the window
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 420) // 2
        y = (screen_height - 430) // 2
        self.root.geometry(f"420x430+{x}+{y}")

    def _create_widgets(self):
        self.main_frame = ctk.CTkFrame(self.root, corner_radius=16, fg_color="#F4F6F8")
        self.main_frame.pack(padx=30, pady=30, fill="both", expand=True)

        ctk.CTkLabel(
            self.main_frame,
            text="Welcome Back",
            font=("Segoe UI", 20, "bold"),
            text_color="#212121"
        ).pack(pady=(20, 4))

        ctk.CTkLabel(
            self.main_frame,
            text="Sign in to continue",
            font=("Segoe UI", 12),
            text_color="#616161"
        ).pack(pady=(0, 20))

        form_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        form_frame.pack(pady=(0, 10))

        self._create_labeled_entry("Username", "Enter your username", form_frame, is_password=False)
        self._create_labeled_entry("Password", "Enter your password", form_frame, is_password=True)

        # Remember Me + Forgot Password
        options_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        options_frame.pack(fill='x', padx=10, pady=(0, 10))

        self.remember_me_var = ctk.BooleanVar(value=False)
        remember_cb = ctk.CTkCheckBox(
            options_frame,
            text="Remember Me",
            variable=self.remember_me_var,
            font=("Segoe UI", 10),
            text_color="#424242"
        )
        remember_cb.pack(side='left')

        forgot_pw = ctk.CTkLabel(
            options_frame,
            text="Forgot Password?",
            font=("Segoe UI", 10, "underline"),
            text_color="#1A73E8",
            cursor='hand2'
        )
        forgot_pw.pack(side='right')
        forgot_pw.bind("<Button-1>", self._on_forgot_password)

        # Login Button
        self.login_button = ctk.CTkButton(
            form_frame,
            text="Sign In",
            command=self.validate_login,
            font=("Segoe UI", 12, "bold"),
            corner_radius=10,
            fg_color="#1976D2",
            hover_color="#1565C0",
            text_color="white",
            height=40,
            width=240
        )
        self.login_button.pack(pady=(8, 10))

        self.root.bind("<Return>", lambda event: self.validate_login())

    def _create_labeled_entry(self, label_text, placeholder, parent, is_password=False):
        ctk.CTkLabel(
            parent,
            text=label_text,
            font=("Segoe UI", 11),
            text_color="#424242"
        ).pack(anchor="w", padx=10, pady=(0, 2))

        entry = ctk.CTkEntry(
            parent,
            placeholder_text=placeholder,
            font=("Segoe UI", 11),
            width=240,
            corner_radius=6,
            fg_color="white",
            text_color="#212121",
            border_color="#BDBDBD",
            show="*" if is_password else ""
        )
        entry.pack(pady=(0, 10))

        if is_password:
            self.password_entry = entry
            self.show_pw_var = ctk.BooleanVar(value=False)
            show_pw_check = ctk.CTkCheckBox(
                parent,
                text="Show Password",
                variable=self.show_pw_var,
                font=("Segoe UI", 10),
                command=self._toggle_password_visibility,
                text_color="#424242"
            )
            show_pw_check.pack(anchor="w", padx=10, pady=(0, 6))
        else:
            self.username_entry = entry

    def _toggle_password_visibility(self):
        self.password_entry.configure(show="" if self.show_pw_var.get() else "*")

    def _on_forgot_password(self, event=None):
        messagebox.showinfo("Forgot Password", "Please contact the administrator to reset your password.")

    def validate_login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        expected_username = self.settings_manager.get_setting("username")
        expected_password = self.settings_manager.get_setting("password")

        if username == expected_username and password == expected_password:
            self.root.unbind("<Return>")
            self.main_frame.destroy()
            self.on_success()
        else:
            messagebox.showerror("Authentication Failed", "Incorrect username or password.")
            self.username_entry.delete(0, "end")
            self.password_entry.delete(0, "end")
            self.username_entry.focus()
