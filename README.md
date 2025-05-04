# HOA Cash Flow Statement Generator

This is a desktop application built with Python and `customtkinter` designed to help Homeowners Associations (HOAs) generate, manage, load, save, and email cash flow statements.

## Features

*   **Graphical User Interface (GUI):** User-friendly interface built with `customtkinter`.
*   **Data Entry:** Fields for entering beginning balances, cash inflows (dues, fees, rentals, etc.), and cash outflows (expenses, deposits, etc.).
*   **Automatic Calculations:** Automatically calculates total receipts, total outflows, and ending cash balances (overall, bank, hand).
*   **Date Selection:** Custom calendar widget (`HoverCalendar`) for easy date selection.
*   **Header Configuration:** Ability to set a custom header address and select a logo image for reports.
*   **File Operations:**
    *   **Load:** Load existing cash flow data from `.docx` (Microsoft Word) or `.pdf` files.
    *   **Export:** Export the current statement to a `.pdf` file (formatted for Folio 8.5" x 13" paper).
    *   **Save:** Save the current statement to a `.docx` file (formatted for Folio 8.5" x 13" paper).
*   **Email Integration:** Send the generated statement (both PDF and DOCX attached) directly via email using Gmail SMTP.
*   **Signatory Management:** Fields to specify recipients and signatory names (Prepared by, Checked by, Noted by x2) which appear in the exported documents.
*   **Settings Management:** A separate settings window to configure:
    *   Sender email address and password (specifically Gmail App Password).
    *   Application login username and password.
*   **Login System:** Simple username/password login for basic access control.
*   **Footer:** Includes predefined footer images (`chud logo.png`, `xu logo.png`).

## Screenshots

*(Optional: You can add screenshots here to showcase the UI)*

*   Login Screen
*   Main Application Window
*   Settings Window
*   Calendar Popup

## Requirements

*   Python 3.7+
*   pip (Python package installer)

## Dependencies

The application relies on the following Python libraries:

*   `customtkinter`: For the modern GUI elements.
*   `tkcalendar`: Base for the custom calendar widget.
*   `reportlab`: For generating PDF documents.
*   `python-docx`: For creating and parsing `.docx` files.
*   `pdfplumber`: For extracting text and tables from `.pdf` files.
*   `Pillow`: For handling images (logo loading, footer images).

## Installation & Setup

1.  **Clone or Download:** Get the project files onto your local machine.
    ```bash
    git clone <repository_url> # If using Git
    cd <repository_directory>
    ```
    Or download the ZIP file and extract it.

2.  **Create a Virtual Environment (Recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

3.  **Install Dependencies:**
    ```bash
    pip install customtkinter tkcalendar reportlab python-docx pdfplumber Pillow
    ```

4.  **Place Footer Images:** Ensure the image files `chud logo.png` and `xu logo.png` are placed inside a folder named `itcc42` within the main project directory. The application looks for them at `itcc42/chud logo.png` and `itcc42/xu logo.png`.

5.  **Configure Gmail for Sending (Important):**
    *   The application uses Gmail's SMTP server (`smtp.gmail.com`).
    *   You **must** use a **Gmail App Password**, not your regular Google account password. Standard password login is often blocked by Google for security reasons when accessed from apps.
    *   **How to get an App Password:**
        *   Go to your Google Account settings.
        *   Enable 2-Step Verification if you haven't already.
        *   Navigate to Security -> 2-Step Verification -> App passwords.
        *   Generate a new App Password (select "Mail" for the app and "Windows Computer" or similar for the device).
        *   Copy the generated 16-character password.
    *   You will enter this App Password in the application's Settings window.

6.  **Settings File (`settings.json`):**
    *   The application uses a `settings.json` file to store email credentials and login details.
    *   If the file doesn't exist when you first run the app or open settings, it will use the hardcoded defaults from `setting.py`.
    *   Saving settings via the Settings window will create or update `settings.json`.

## Usage

1.  **Run the Application:**
    ```bash
    python main.py
    ```

2.  **Login:**
    *   The login window will appear.
    *   Enter the username and password.
    *   By default (if `settings.json` doesn't exist or hasn't been changed), the credentials are:
        *   Username: `user`
        *   Password: `123`
    *   These can be changed in the Settings window after logging in.

3.  **Main Window:**
    *   **Header Config:**
        *   Enter the desired **Header Address**.
        *   Click **Select Logo Image** to choose a logo file for the reports.
        *   Click **Manage Settings** to open the settings window (see Configuration below).
        *   Click the **Date Button** (shows current/selected date) to open the calendar and select the report date.
    *   **Data Entry:** Fill in the amounts for beginning balances, inflows, and outflows in the respective sections. Fields are automatically formatted as currency. Calculated fields (Totals, Ending Balances) are disabled and update automatically.
    *   **Signatories/Recipients:** Fill in the recipient emails (comma-separated) and the names for Prepared by, Noted by (x2), and Checked by.
    *   **Action Buttons:**
        *   **Load (Ctrl+L):** Open a file dialog to select a `.docx` or `.pdf` file to load data from.
        *   **Clear Fields:** Clears all cash flow data entry fields (keeps logo, address, date, signatories).
        *   **Export PDF (Ctrl+E):** Save the current statement as a PDF file.
        *   **Save Word (Ctrl+W):** Save the current statement as a DOCX file.
        *   **Email (Ctrl+G):** Send the generated PDF and DOCX files via email to the specified recipients using the configured sender credentials.
    *   **Footer:** The footer section displays the CHUDD and XU logos.

4.  **Keyboard Shortcuts:**
    *   `Ctrl+L`: Load Document
    *   `Ctrl+E`: Export to PDF
    *   `Ctrl+W`: Save to Word
    *   `Ctrl+G`: Send Email
    *   `Ctrl+Q`: Quit Application

## Configuration (Settings Window)

*   Access the Settings window via the **Manage Settings** button on the main screen.
*   Here you can view and edit:
    *   **Sender Email:** The Gmail address used to send emails.
    *   **Sender Password:** The **Gmail App Password** (16 characters).
    *   **Username:** The username required to log into this application.
    *   **Password:** The password required to log into this application.
*   Click **Save** to apply changes (updates `settings.json`). Saved login credentials take effect on the next login.
*   Click **Cancel** to close without saving.

