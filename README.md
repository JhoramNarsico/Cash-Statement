# HOA Cash Flow Statement Generator

## Overview

The HOA Cash Flow Statement Generator is a desktop application designed to help Homeowners Associations (HOAs) easily create, manage, and distribute their cash flow statements. It provides a user-friendly interface for data entry, automatically calculates totals, and allows for exporting statements to PDF and DOCX formats. The application also features functionality to load data from existing documents and manage email attachments for distribution.

## Features

*   **Intuitive Data Entry**: Simple forms for inputting beginning balances, cash inflows, and cash outflows.
*   **Automatic Calculations**: Automatically calculates total receipts, total outflows, and ending cash balances.
*   **Document Export**: Export cash flow statements to professionally formatted PDF and DOCX files.
*   **Data Loading**: Load financial data from existing PDF or DOCX statements, including header information like logo and address.
*   **Email Attachments Management**: A dedicated window to add multiple files (reports, notices, etc.) and send them via email.
*   **Secure Login**: Protects access to the application with a username and password.
*   **Configurable Settings**:
    *   Manage login credentials.
    *   Manage email sender credentials (specifically designed for Gmail with App Passwords).
    *   Settings are saved persistently.
*   **Customizable Header**: Include HOA logo and address in the generated statements.
*   **Date Selection**: Easy date selection for reports using an interactive calendar.
*   **Responsive Design**: Adapts to different screen sizes.
*   **Loading Indicators**: Provides feedback for long-running operations.

## Prerequisites

*   Python 3.8 or higher
*   `pip` (Python package installer)

## Installation

1.  **Clone the repository (or download the source code):**
    ```bash
    git clone https://github.com/your-username/hoa-cash-flow-generator.git
    cd hoa-cash-flow-generator
    ```

2.  **Create and activate a virtual environment (recommended):**
    *   On Windows:
        ```bash
        python -m venv venv
        .\venv\Scripts\activate
        ```
    *   On macOS/Linux:
        ```bash
        python3 -m venv venv
        source venv/bin/activate
        ```

3.  **Install the required dependencies:**
    Create a `requirements.txt` file with the following content:
    ```
    customtkinter
    tkcalendar
    Pillow
    reportlab
    pdfplumber
    camelot-py[cv]
    python-docx
    PyMuPDF
    ```
    Then run:
    ```bash
    pip install -r requirements.txt
    ```
    *Note: `camelot-py[cv]` requires OpenCV. If you encounter issues, you might need to install `opencv-python` separately or ensure its dependencies are met for your OS.*

## Usage

1.  **Run the application:**
    ```bash
    python main.py
    ```

2.  **Login:**
    *   The application will first present a login screen.
    *   Default credentials are:
        *   Username: `user`
        *   Password: `123`
    *   These can be changed via the "Account Settings" menu within the application after logging in.

3.  **Main Application Interface:**
    *   **Data Entry**: Fill in the fields for beginning balances, cash inflows, and cash outflows. Totals will update automatically.
    *   **Header Configuration**:
        *   **Header Address**: Enter the address to appear on the statement.
        *   **Header Logo**: Click "Select Logo" to choose an image file. "Remove Logo" clears the selection.
    *   **Report Date**: Click the date button to open a calendar and select the report month/year.
    *   **Action Buttons**:
        *   **Load (Ctrl+L)**: Load data from an existing DOCX or PDF statement.
        *   **Clear Fields**: Clears all cash flow data fields, recipient list, and additional attachments.
        *   **Export PDF (Ctrl+E)**: Save the current statement as a PDF file.
        *   **Save Word (Ctrl+W)**: Save the current statement as a DOCX file.
    *   **Account Settings / Manage Attachments**:
        *   **Account Settings**: Opens a window to change login credentials and email sender credentials.
        *   **Manage Attachments & Send**: Opens a window to:
            *   Enter recipient email addresses (comma-separated).
            *   Add or remove additional files to be sent along with an email.
            *   Send the email with all listed attachments.

## Configuration

### Login Credentials

*   Default: Username `user`, Password `123`.
*   To change: After logging in, click the "Account Settings" button. Enter new credentials and click "Save". Changes apply on the next application start.

### Email Sender Credentials (for "Manage Attachments & Send")

*   The application is configured to use Gmail for sending emails.
*   **Default**: `chuddcdo@gmail.com` 
*   **To change**:
    1.  Go to "Account Settings".
    2.  Enter your Gmail address in "Sender Email".
    3.  **Crucially, you must use a Gmail App Password, not your regular Gmail password.**
        *   **How to get a Gmail App Password:**
            1.  Go to your Google Account (myaccount.google.com).
            2.  Ensure 2-Step Verification is turned ON for your Google Account.
            3.  Navigate to Security > 2-Step Verification > App passwords (you might need to sign in again).
            4.  Select "Mail" for the app and "Other (Custom name)" for the device. Give it a name (e.g., "HOA Cash Flow App").
            5.  Google will generate a 16-character App Password. Copy this password.
            6.  Enter this 16-character App Password into the "Sender Password" field in the application's settings.
    4.  Click "Save".

### Settings Storage

*   Settings (login and email credentials) are stored in a `settings.json` file.
*   This file is located in a user-specific application data directory:
    *   **Windows**: `C:\Users\<YourUsername>\AppData\Local\HOACashFlowSettings\settings.json` (or similar, `LOCALAPPDATA`)
    *   **macOS**: `~/Library/Application Support/HOACashFlowSettings/settings.json`
    *   **Linux**: `~/.config/HOACashFlowSettings/settings.json` (or `XDG_CONFIG_HOME`)
