# HOA Cash Flow Statement Generator

## Overview

The HOA Cash Flow Statement Generator is a desktop application designed to help Homeowners Associations (HOAs) easily create, manage, and distribute their monthly cash flow statements. It simplifies financial record-keeping by providing a user-friendly interface for inputting income and expenses, automatically calculating totals, and generating professional-looking documents in PDF and Word formats. The application also includes functionality to email these statements directly to specified recipients.

## Features

*   **Intuitive Data Entry:**
    *   Dedicated fields for all common HOA income sources (monthly dues, certifications, membership fees, vehicle stickers, rentals, solicitations, interest income, livelihood fees, other inflows).
    *   Comprehensive expense tracking (snacks/meals, transportation, office supplies, printing, labor, billboard, cleaning, miscellaneous, federation fees, uniforms, BOD meetings, general assembly, cash deposits, withholding tax, refunds, other outflows).
    *   Separate input for beginning cash balances (cash in bank, cash on hand).
*   **Automatic Calculations:**
    *   Real-time calculation of total receipts, total cash outflows, and ending cash balance.
    *   Automatic calculation of ending cash in bank and cash on hand.
*   **Professional Document Generation:**
    *   Export cash flow statements to **PDF** format, suitable for official records and sharing.
    *   Save statements as **Word (.docx)** documents for easy editing or archiving.
    *   Customizable document header including HOA address and logo.
    *   Standardized footer with signatory lines for Treasurer, Auditor, HOA President, and CHUDD HCD-CORDS representative.
*   **Email Functionality:**
    *   Send generated PDF and Word documents directly via email to multiple recipients.
    *   Uses Gmail SMTP (requires sender's Gmail credentials and an "App Password" if 2-Step Verification is enabled).
*   **Data Management:**
    *   Load data from previously saved Word (.docx) or PDF documents (PDF loading may have limitations with complex, externally generated files).
    *   Clear all financial data fields to start a new statement.
*   **User-Friendly Interface:**
    *   Modern look and feel using CustomTkinter.
    *   Tooltips for guidance on various input fields and buttons.
    *   Interactive calendar for easy date selection.
    *   Responsive layout elements.
*   **Secure Login:**
    *   Password-protected access to the application.
    *   Default credentials provided, configurable via a settings panel.
*   **Customizable Settings:**
    *   Configure sender email credentials (email and app password).
    *   Change application login username and password.
    *   Select a custom logo for document headers.
    *   Set the header address for documents.
    *   Specify default footer images.
*   **Keyboard Shortcuts:**
    *   `Ctrl+L`: Load Document
    *   `Ctrl+E`: Export to PDF
    *   `Ctrl+W`: Save to Word
    *   `Ctrl+G`: Send Email
    *   `Ctrl+Q`: Quit Application

## System Requirements

*   **Operating System:** Windows 7/8/10/11 (Installer is for Windows)
*   **Internet Connection:** Required for sending emails.

## Installation

1.  **Download:** Obtain the `HOA_Cash_Flow_Setup.exe` installer file.
2.  **Run Installer:** Double-click `HOA_Cash_Flow_Setup.exe` to start the installation process.
3.  **Administrator Privileges:** The installer will likely request administrator privileges to install the application in the Program Files directory. Please allow this.
4.  **Follow Prompts:** Follow the on-screen instructions. You can typically accept the default settings.
    *   Choose an installation location (default is recommended).
    *   Decide whether to create Desktop and Start Menu icons.
5.  **Complete Installation:** Once the installation is complete, you can launch the application from the Start Menu or Desktop icon if created.

## First Time Use & Configuration

1.  **Login:**
    *   The default login credentials are:
        *   Username: `user`
        *   Password: `123`
2.  **Settings:** It is highly recommended to configure the settings immediately after the first login:
    *   Click the "Manage Settings" button.
    *   **Sender Email & Password:** Enter your Gmail address and an "App Password" (if you have 2-Step Verification enabled on your Gmail account).
        *   To create an App Password: Go to your Google Account > Security > 2-Step Verification > App passwords.
    *   **Application Login:** Change the default username and password for better security.
    *   Click "Save".
3.  **Header Configuration:**
    *   **Header Address:** Enter your HOA's official address in the "Header Address" field on the main screen.
    *   **Header Logo:** Click "Select Logo Image" to choose your HOA's logo. This will be used in generated PDF and Word documents.

## Usage

1.  **Enter Beginning Balances:** Fill in "Cash in Bank-beg" and "Cash on Hand-beg".
2.  **Input Inflows:** Record all income in the "Cash Inflows" section.
3.  **Input Outflows:** Record all expenses in the "Cash Outflows" section.
4.  **Verify Totals:** The "Totals (Calculated)" and "Ending Cash Balances (Calculated)" sections will update automatically.
5.  **Set Report Date:** Click the date button (e.g., "Jan 01, 2024") to select the correct month/period for the statement.
6.  **Enter Signatory Names:** Fill in the names for "Prepared by", "Noted by", and "Checked by" at the bottom of the screen.
7.  **Export/Save:**
    *   Use "Export PDF (Ctrl+E)" or "Save Word (Ctrl+W)" to generate the statement.
8.  **Email:**
    *   Enter recipient email addresses (comma-separated) in the "Recipients" field.
    *   Click "Email (Ctrl+G)" to send the generated PDF and Word files.

## Uninstallation

1.  Go to **Windows Settings > Apps > Apps & features**.
2.  Find "HOA Cash Flow Statement" in the list.
3.  Click on it and select "Uninstall".
4.  Follow the prompts to remove the application.
    *   User settings (like `settings.json`) stored in `%LOCALAPPDATA%\HOACashFlowSettings` may need to be manually deleted if you wish to remove all traces.

## For Developers: Build Process

This application is built using Python and packaged for distribution.

### 1. PyInstaller (Creating the Executable Bundle)

PyInstaller is used to bundle the Python application and its dependencies into a standalone executable and supporting files.

*   **Environment:** A dedicated Python virtual environment is used to manage dependencies.
*   **Dependencies:** Key libraries include `customtkinter`, `Pillow`, `tkcalendar`, `reportlab`, `python-docx`, `pdfplumber`, `camelot-py`, `PyMuPDF`.
*   **`.spec` File:** A `main.spec` file is used to configure the PyInstaller build. This file specifies:
    *   The main script (`main.py`).
    *   Data files to be included (e.g., `chud logo.png`, `xu logo.png`, CustomTkinter assets).
    *   Hidden imports to ensure all necessary modules are bundled.
    *   Executable name and other build options.
*   **Build Command:**
    ```bash
    # Ensure virtual environment is active
    pyinstaller main.spec
    ```
*   **Output:** This produces a `dist/HOA_Cash_Flow/` directory containing the application executable and all its bundled dependencies. This directory is then used as the input for Inno Setup.

### 2. Inno Setup (Creating the Windows Installer)

Inno Setup is used to create a user-friendly Windows installer (`.exe`) from the files generated by PyInstaller.

*   **Script File (`.iss`):** An Inno Setup script (e.g., `HOA_Cash_Flow_Installer.iss`) defines the installer's behavior. Key sections include:
    *   `[Setup]`: Defines application metadata (name, version, publisher, AppId), installation directory, privileges, output filename, compression, etc.
    *   `[Languages]`: Specifies supported languages.
    *   `[Tasks]`: Defines optional installation tasks like creating desktop icons.
    *   `[Files]`: Specifies which files and folders (from the PyInstaller `dist/HOA_Cash_Flow/` output) to include in the installer and where they should be placed on the user's system (typically ` {app}` which maps to the installation directory).
    *   `[Icons]`: Defines Start Menu and Desktop shortcuts.
    *   `[Run]`: Specifies commands to run after installation (e.g., launching the application).
    *   `[UninstallDelete]`: Defines files and directories to be removed during uninstallation.
*   **Compilation:** The `.iss` script is compiled using the Inno Setup Compiler to produce the final `setup.exe`.

## Troubleshooting / Known Issues

*   **Permission Denied for `settings.json`:** This should be resolved by the application now saving settings to the user's AppData folder. If it occurs during development, ensure you have write permissions in your project directory.
*   **"Load from PDF" Inaccuracy:** Loading data from PDFs that were not generated by this application, or that have very complex/non-standard layouts, may not be perfectly accurate. The `camelot-py` library (which benefits from Ghostscript for some PDFs) may misinterpret some tables. If you encounter issues with specific PDFs, ensuring Ghostscript is installed and in your system PATH *might* improve results.
*   **Email Sending Failures:**
    *   Verify sender email and app password are correct in settings.
    *   Ensure an active internet connection.
    *   Check if your antivirus or firewall is blocking the application's SMTP connection (port 587 for Gmail).
    *   For Gmail, ensure "Less secure app access" is NOT required if using an App Password (App Passwords bypass this).

## License

(Specify your license here, e.g., MIT, GPL, Proprietary, etc. If you don't have one, you might consider adding one like MIT for open source.)

Example:
This project is licensed under the MIT License - see the LICENSE.md file for details (if you create one).

---
