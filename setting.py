import customtkinter as ctk
import json
import os
from tkinter import messagebox
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class SettingsManager:
    """Manages loading and saving of application settings to a JSON file."""
    def __init__(self, settings_file="settings.json"):
        self.settings_file = settings_file
        self.settings = {
            "sender_email": "chuddcdo@gmail.com",
            "sender_password": "jfyb eoog ukxr hhiq",
            "username": "user",
            "password": "123"
        }
        self.load_settings()

    def load_settings(self):
        """Load settings from the JSON file or use defaults if file doesn't exist."""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    loaded_settings = json.load(f)
                    self.settings.update(loaded_settings)
                logging.info(f"Settings loaded from {self.settings_file}")
            else:
                logging.info("No settings file found, using default settings.")
        except Exception as e:
            logging.error(f"Error loading settings: {e}")
            messagebox.showerror("Error", f"Failed to load settings: {e}")

    def save_settings(self, new_settings):
        """Save settings to the JSON file."""
        try:
            self.settings.update(new_settings)
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=4)
            logging.info(f"Settings saved to {self.settings_file}")
            return True
        except Exception as e:
            logging.error(f"Error saving settings: {e}")
            messagebox.showerror("Error", f"Failed to save settings: {e}")
            return False

    def get_setting(self, key):
        """Get a specific setting value."""
        return self.settings.get(key, "")

class SettingsWindow:
    """GUI for editing application settings."""
    def __init__(self, parent, settings_manager):
        self.parent = parent
        self.settings_manager = settings_manager
        self.visibility_states = {}
        self.entries = {}
        self.create_window()

    def create_window(self):
        self.window = ctk.CTkToplevel(self.parent)
        self.window.title("Settings")
        self.window.geometry("400x400")
        self.window.resizable(False, False)
        self.window.transient(self.parent)
        self.window.grab_set()

        self.window.update_idletasks()
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - 400) // 2
        y = (screen_height - 400) // 2
        self.window.geometry(f"400x400+{x}+{y}")

        main_frame = ctk.CTkFrame(self.window, corner_radius=10, fg_color="#F5F5F5")
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        ctk.CTkLabel(
            main_frame,
            text="Application Settings",
            font=("Roboto", 16, "bold"),
            text_color="#333333"
        ).pack(pady=(10, 20))

        fields = [
            ("Sender Email:", "sender_email"),
            ("Sender Password:", "sender_password"),
            ("Username:", "username"),
            ("Password:", "password")
        ]

        for label_text, key in fields:
            field_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
            field_frame.pack(fill="x", padx=10, pady=5)

            ctk.CTkLabel(
                field_frame,
                text=label_text,
                font=("Roboto", 12),
                text_color="#333333",
                width=120,
                anchor="w"
            ).pack(side="left")

            entry_container = ctk.CTkFrame(field_frame, fg_color="transparent")
            entry_container.pack(side="left", fill="x", expand=True)

            is_password = "password" in key
            entry = ctk.CTkEntry(
                entry_container,
                font=("Roboto", 12),
                width=240,
                corner_radius=6,
                fg_color="#FAFAFA",
                text_color="#333333",
                border_color="#B0BEC5",
                show="*" if is_password else ""
            )
            entry.insert(0, self.settings_manager.get_setting(key))
            entry.pack(side="left", fill="x", expand=True, padx=(0, 32) if is_password else 0)

            self.entries[key] = entry

            if is_password:
                self.visibility_states[key] = True

                toggle_button = ctk.CTkButton(
                    entry_container,
                    text="👁",
                    width=30,
                    height=28,
                    corner_radius=6,
                    command=lambda e=entry, k=key: self.toggle_visibility(e, k),
                    fg_color="transparent",
                    hover_color="#E0E0E0",
                    text_color="#333333"
                )
                toggle_button.place(relx=1.0, x=-4, rely=0.5, anchor="e")

        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(pady=20)

        ctk.CTkButton(
            button_frame,
            text="Save",
            command=self.save_settings,
            font=("Roboto", 12),
            corner_radius=8,
            fg_color="#2196F3",
            hover_color="#1976D2",
            text_color="#FFFFFF",
            width=100
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            button_frame,
            text="Cancel",
            command=self.window.destroy,
            font=("Roboto", 12),
            corner_radius=8,
            fg_color="#757575",
            hover_color="#616161",
            text_color="#FFFFFF",
            width=100
        ).pack(side="left", padx=5)

    def toggle_visibility(self, entry, key):
        if self.visibility_states[key]:
            entry.configure(show="")
        else:
            entry.configure(show="*")
        self.visibility_states[key] = not self.visibility_states[key]

    def save_settings(self):
        new_settings = {key: entry.get().strip() for key, entry in self.entries.items()}
        if not new_settings["username"] or not new_settings["password"]:
            messagebox.showerror("Error", "Username and Password cannot be empty.")
            return
        if self.settings_manager.save_settings(new_settings):
            messagebox.showinfo("Success", "Settings saved successfully! New login credentials will apply on next login.")
            self.window.destroy()
        else:
            messagebox.showerror("Error", "Failed to save settings.")
