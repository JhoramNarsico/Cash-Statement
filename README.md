# HOA Cash Flow Statement Generator

This application provides a graphical user interface (GUI) for generating Homeowners Association (HOA) Cash Flow Statements. It allows users to input financial data, calculates totals, loads data from existing DOCX/PDF files, exports reports to PDF and DOCX formats, and sends these reports via email.

## Features

*   Input fields for beginning balances, cash inflows, and cash outflows.
*   Automatic calculation of total receipts, total outflows, and ending balances.
*   Load data from previously generated DOCX or PDF statements.
*   Export the statement to a formatted PDF file (Folio size).
*   Save the statement to a formatted DOCX file (Folio size).
*   Include HOA logo and address in the header of exported documents.
*   Configure signatory names (Prepared by, Checked by, Noted by).
*   Send the generated PDF and DOCX files as email attachments via Gmail.
*   Secure login system.
*   Configurable settings (email credentials, login details) via a settings window and `settings.json`.

## Prerequisites

Before you begin, ensure you have the following installed:

1.  **Python:** Version 3.8 or higher is recommended. You can download it from [python.org](https://www.python.org/downloads/). Make sure to check the option "Add Python to PATH" during installation on Windows.
2.  **pip:** The Python package installer. It usually comes bundled with Python installs. You can verify by running `pip --version` in your terminal or command prompt.

## Installation

Follow these steps to set up the project on your local machine:

1.  **Clone the repository (or download the source code):**
    ```bash
    git clone <your-repository-url> # Replace with your actual repo URL if applicable
    cd <repository-folder-name>      # Navigate into the project directory
    ```
    Alternatively, download the project files as a ZIP and extract them.

2.  **Install required Python libraries:**
    Create a file named `requirements.txt` in the project's root directory with the following content:

    ```txt
    customtkinter
    tkcalendar
    reportlab
    python-docx
    pdfplumber
    Pillow
    ```

    Then, open your terminal or command prompt, navigate to the project directory, and run:
    ```bash
    pip install -r requirements.txt
    ```
    This command will install all the necessary dependencies listed in the file.

## Running the Application

Once the installation is complete, you can run the application directly from the source code:

1.  **Navigate to the project directory** in your terminal or command prompt.
2.  **Run the main script:**
    ```bash
    python main.py
    ```
3.  The login window should appear. Use the credentials configured in `settings.json` (default: `user` / `123`) or the ones you set via the Settings window on a previous run.
4.  After successful login, the main application window will open.

## Creating an Installer (using PyInstaller)

PyInstaller bundles a Python application and all its dependencies into a single package. Users can run the packaged app without installing a Python interpreter or any modules.

1.  **Install PyInstaller:**
    If you don't have PyInstaller installed, open your terminal or command prompt and run:
    ```bash
    pip install pyinstaller
    ```

2.  **Bundle the application:**
    Navigate to the project's root directory in your terminal. Run the following PyInstaller command:

    ```bash
    pyinstaller --name "HOACashFlow" \
                --onefile \
                --windowed \
                --icon="path/to/your/icon.ico" \
                --add-data "Cash-Statement/chud logo.png:Cash-Statement" \
                --add-data "Cash-Statement/xu logo.png:Cash-Statement" \
                --add-data "settings.json:." \
                main.py
    ```

    **Explanation of Flags:**
    *   `--name "HOACashFlow"`: Sets the name of the executable and the build folders.
    *   `--onefile`: Creates a single executable file (can be slower to start). Remove this if you prefer a folder with multiple files (faster startup).
    *   `--windowed` or `-w`: Prevents a console window from appearing when the GUI application runs. Essential for GUI apps.
    *   `--icon="path/to/your/icon.ico"`: (Optional) Specifies an icon file (`.ico` on Windows) for the executable. Replace the path with your actual icon file.
    *   `--add-data "source:destination"`: This is crucial for including non-code files like images and configuration files.
        *   The format is `source_path;destination_folder` (Windows) or `source_path:destination_folder` (Mac/Linux).
        *   `source_path`: The path to the file or folder in your project structure.
        *   `destination_folder`: The path where the file/folder should be placed *relative to the executable* at runtime. `.` means the same directory as the executable.
        *   **Important:** Adjust the source paths (`Cash-Statement/chud logo.png`, `Cash-Statement/xu logo.png`, `settings.json`) if your project structure is different. Ensure the paths are correct relative to where you run the `pyinstaller` command.
    *   `main.py`: The entry point script of your application.

3.  **Find the Installer:**
    PyInstaller will create a `build` folder and a `dist` folder in your project directory. The final executable (e.g., `HOACashFlow.exe` on Windows) will be inside the `dist` folder.

4.  **Distribution:**
    You can now distribute the contents of the `dist` folder (either the single `.exe` file if using `--onefile`, or the entire folder if not). Users can run the executable directly without needing Python installed. **Crucially, the data files (images, `settings.json`) must be present relative to the executable as specified in the `--add-data` flags.** If you used `--onefile`, PyInstaller unpacks these temporarily at runtime.

## Configuration

*   **`settings.json`:** This file stores application settings:
    *   `sender_email`: The Gmail address used for sending emails.
    *   `sender_password`: **Important:** This should be a Gmail [App Password](https://support.google.com/accounts/answer/185833), not your regular Gmail password, especially if 2-Factor Authentication is enabled.
    *   `username`: The username for logging into the application.
    *   `password`: The password for logging into the application.
*   **Settings Window:** You can modify these settings (except the password visibility) through the "Manage Settings" button within the application after logging in.

## Troubleshooting

*   **Login Failed:** Ensure the username and password match those in `settings.json`. Check for typos or extra spaces.
*   **Email Sending Failed:**
    *   Verify the `sender_email` and `sender_password` in `settings.json`.
    *   Ensure you are using a Gmail **App Password**.
    *   Check your internet connection.
*   **Installer Issues (File Not Found Errors):**
    *   Double-check the paths used in the `--add-data` flags in the PyInstaller command. They must be correct relative to the location where you run the command.
    *   Make sure the destination paths in `--add-data` match how the application expects to find the files at runtime (e.g., if the code looks for `Cash-Statement/logo.png`, the destination should likely be `Cash-Statement`).
*   **Module Not Found Errors:** Make sure all dependencies were installed correctly using `pip install -r requirements.txt`. If running the bundled executable, this usually indicates PyInstaller missed a dependency (less common with `--add-data` used correctly).

---
