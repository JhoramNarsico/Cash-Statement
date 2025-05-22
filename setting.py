import customtkinter as ctk
import json
import os
import sys # Added import for sys.platform
from tkinter import messagebox
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_app_data_path(filename="settings.json"):
    """Gets the full path to a file in the user's app data directory."""
    app_data_dir = None
    your_app_name_folder = "HOACashFlowSettings" # IMPORTANT: Customize this to a unique name for your app

    # Try common environment variables for AppData
    if os.name == 'nt': # Windows
        app_data_dir = os.getenv('LOCALAPPDATA')
        if not app_data_dir: # Fallback if LOCALAPPDATA is not set
            app_data_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local")
    elif os.name == 'posix': # Linux, macOS
        # macOS typically uses ~/Library/Application Support/YourAppName
        # Linux typically uses ~/.config/YourAppName or ~/.local/share/YourAppName
        xdg_config_home = os.getenv('XDG_CONFIG_HOME')
        if xdg_config_home and os.path.isdir(xdg_config_home):
            app_data_dir = xdg_config_home
        else:
            app_data_dir = os.path.join(os.path.expanduser("~"), ".config") # Common Linux fallback

        # For macOS, a more specific path is often preferred
        if sys.platform == "darwin": # sys needs to be imported if you use this check
             mac_app_support = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
             if os.path.isdir(mac_app_support):
                 app_data_dir = mac_app_support


    if not app_data_dir or not os.path.isdir(app_data_dir):
        # Last resort: use current working directory (mainly for development or if standard paths fail)
        logging.warning(
            "Could not determine standard app data directory or it's not a directory. "
            f"Using current working directory for '{filename}'."
        )
        # In this fallback, don't create YourAppName subfolder to avoid clutter in dev environment
        return os.path.join(os.getcwd(), filename)

    # Create a subdirectory for your application within the determined app_data_dir
    app_specific_dir = os.path.join(app_data_dir, your_app_name_folder)

    try:
        if not os.path.exists(app_specific_dir):
            os.makedirs(app_specific_dir)
            logging.info(f"Created application settings directory: {app_specific_dir}")
    except OSError as e:
        logging.error(f"Failed to create application settings directory {app_specific_dir}: {e}")
        # If directory creation fails in AppData, fallback to current working directory
        logging.warning(f"Falling back to current working directory for '{filename}'.")
        return os.path.join(os.getcwd(), filename)

    return os.path.join(app_specific_dir, filename)


class SettingsManager:
    """Manages loading and saving of application settings to a JSON file."""
    def __init__(self):
        self.settings_file = get_app_data_path("settings.json")
        logging.info(f"SettingsManager initialized. Settings file path: {self.settings_file}")
        self.settings = {
            "sender_email": "chuddcdo@gmail.com",
            "sender_password": "oaki rktd kgqx cpwt",
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
                logging.info(f"No settings file found at {self.settings_file}. Using default settings and attempting to save them.")
                # Attempt to save default settings if file doesn't exist.
                # This also helps to create the file/directory structure early and test writability.
                self.save_settings({}) # Pass empty dict to just save current defaults.
        except json.JSONDecodeError as e:
            logging.error(f"Error decoding JSON from {self.settings_file}: {e}. Using default settings.")
            # If file is corrupt, use defaults. Optionally, you could try to backup/rename corrupt file.
        except Exception as e:
            logging.error(f"Error loading settings from {self.settings_file}: {e}. Using default settings.")

    def save_settings(self, new_settings):
        """Save settings to the JSON file."""
        try:
            self.settings.update(new_settings)
            # Ensure the directory exists before trying to write the file.
            # get_app_data_path attempts to create it, but this is a good safeguard.
            settings_dir = os.path.dirname(self.settings_file)
            if not os.path.exists(settings_dir):
                try:
                    os.makedirs(settings_dir)
                    logging.info(f"Created directory during save: {settings_dir}")
                except OSError as e:
                    logging.error(f"Failed to create directory {settings_dir} during save: {e}")
                    messagebox.showerror("Error", f"Failed to create settings directory:\n{str(e)}")
                    return False

            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=4)
            logging.info(f"Settings saved successfully to {self.settings_file}")
            return True
        except Exception as e:
            logging.error(f"Error saving settings to {self.settings_file}: {e}")
            messagebox.showerror(
                "Error",
                f"Failed to save settings to '{os.path.basename(self.settings_file)}':\n{str(e)}\n\n"
                f"Please check permissions for the application's settings folder:\n{os.path.dirname(self.settings_file)}"
            )
            return False

    def get_setting(self, key):
        """Get a specific setting value."""
        return self.settings.get(key, "")

class SettingsWindow:
    """GUI for editing application settings."""
    def __init__(self, parent, settings_manager):
        self.parent = parent
        self.settings_manager = settings_manager
        self.visibility_states = {} # Stores whether a password field is currently masked (True) or shown (False)
        self.entries = {}
        self.create_window()

    def create_window(self):
        self.window = ctk.CTkToplevel(self.parent)
        self.window.title("Settings")
        self.window.geometry("400x400") # Initial size
        self.window.resizable(False, False)
        self.window.transient(self.parent) # Keep on top of parent
        self.window.grab_set() # Modal behavior

        # Center the window (optional, but good UX)
        self.window.update_idletasks() # Ensure dimensions are updated
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        win_width = self.window.winfo_width()
        win_height = self.window.winfo_height()
        x = parent_x + (parent_width // 2) - (win_width // 2)
        y = parent_y + (parent_height // 2) - (win_height // 2)
        self.window.geometry(f"{win_width}x{win_height}+{x}+{y}")


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
                width=120, # Give label a fixed width to align entries
                anchor="w"
            ).pack(side="left")

            entry_container = ctk.CTkFrame(field_frame, fg_color="transparent")
            entry_container.pack(side="left", fill="x", expand=True)

            is_password = "password" in key.lower() # Make check case-insensitive
            entry_show_char = "*" if is_password else ""

            entry = ctk.CTkEntry(
                entry_container,
                font=("Roboto", 12),
                # width=240, # Let it expand, or set a specific width
                corner_radius=6,
                fg_color="#FAFAFA",
                text_color="#333333",
                border_color="#B0BEC5",
                show=entry_show_char
            )
            entry.insert(0, self.settings_manager.get_setting(key))
            entry.pack(side="left", fill="x", expand=True, padx=(0, 5 if is_password else 0)) # Add some padding if toggle exists

            self.entries[key] = entry
            self.visibility_states[key] = is_password # True if it's a password field (initially masked)

            if is_password:
                toggle_button = ctk.CTkButton(
                    entry_container,
                    text="👁",
                    width=30,
                    height=entry.winfo_reqheight(), # Match entry height
                    corner_radius=6,
                    command=lambda e=entry, k=key: self.toggle_visibility(e, k),
                    fg_color="#E0E0E0", # Light gray button
                    hover_color="#BDBDBD",
                    text_color="#333333"
                )
                toggle_button.pack(side="left", padx=(0,5)) # Pack next to entry

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
            fg_color="#757575", # A slightly darker gray for cancel
            hover_color="#616161",
            text_color="#FFFFFF",
            width=100
        ).pack(side="left", padx=5)

    def toggle_visibility(self, entry, key):
        if self.visibility_states[key]:  # If currently masked (True)
            entry.configure(show="")     # Unmask
            self.visibility_states[key] = False
        else:                            # If currently shown (False)
            entry.configure(show="*")    # Mask
            self.visibility_states[key] = True

    def save_settings(self):
        new_settings = {key: entry.get().strip() for key, entry in self.entries.items()}
        if not new_settings.get("username", "").strip() or not new_settings.get("password", "").strip():
            messagebox.showerror("Error", "Username and Password cannot be empty.")
            return

        if self.settings_manager.save_settings(new_settings):
            messagebox.showinfo("Success", "Settings saved successfully!\nNew login credentials will apply on the next application start.")
            self.window.destroy()
        # Else, the error message is shown by SettingsManager.save_settings()
