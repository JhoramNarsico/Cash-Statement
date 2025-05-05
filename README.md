# HOA Cash Flow Statement Generator

This application helps Homeowners Associations (HOAs) generate, manage, load, save, and email cash flow statements using a Python GUI built with `customtkinter`.

## Requirements

*   Python 3.7 or later
*   `pip` (Python package installer)

## Installation

Follow these steps to set up the application environment:

1.  **Clone or Download the Repository:**
    ```bash
    git clone <repository_url> # Or download and extract the ZIP file
    cd <repository_directory>
    ```

2.  **Create and Activate a Virtual Environment (Recommended):**
    *   **On macOS/Linux:**
        ```bash
        python3 -m venv venv
        source venv/bin/activate
        ```
    *   **On Windows:**
        ```bash
        python -m venv venv
        .\venv\Scripts\activate
        ```

3.  **Install Required Libraries:**
    Install all necessary Python packages using pip:
    ```bash
    pip install customtkinter tkcalendar reportlab python-docx pdfplumber Pillow
    ```

4.  **Place Footer Images:**
    *   Ensure you have a folder named `itcc42` in the main project directory.
    *   Place the required footer images (`chud logo.png` and `xu logo.png`) inside this `itcc42` folder. The application expects them at these paths:
        *   `itcc42/chud logo.png`
        *   `itcc42/xu logo.png`

## Running the Application (from Source)

Once the installation is complete:

1.  **Ensure your virtual environment is activated.** (See step 2 in Installation).
2.  **Navigate to the project directory** in your terminal.
3.  **Run the main script:**
    ```bash
    python main.py
    ```
4.  **Login:** The application will start with a login window. Use the credentials configured in `settings.json` or the defaults if the file doesn't exist:
    *   Default Username: `user`
    *   Default Password: `123`
    (These can be changed via the "Manage Settings" button after logging in).

## Creating an Executable with PyInstaller

PyInstaller allows you to package the Python application and all its dependencies into a single executable file that can be run on other computers without requiring Python or the libraries to be installed.

1.  **Install PyInstaller:**
    If you don't have PyInstaller installed in your virtual environment, install it:
    ```bash
    pip install pyinstaller
    ```

2.  **Prepare Assets:** Make sure the `itcc42` folder with the footer images exists in the same directory where you run the PyInstaller command. PyInstaller needs to find these files during the build process to bundle them. The `settings.json` file will also be bundled if it exists.

3.  **Run PyInstaller Command:**
    Open your terminal, ensure your virtual environment is active, navigate to the project's root directory (where `main.py` is located), and run the following command:

    *   **On macOS/Linux:**
        ```bash
        pyinstaller --onefile --windowed --add-data="itcc42:itcc42" --add-data="settings.json:." main.py
        ```

    *   **On Windows:**
        ```bash
        pyinstaller --onefile --windowed --add-data="itcc42;itcc42" --add-data="settings.json;." main.py
        ```

    **Explanation of Options:**
    *   `--onefile`: Creates a single executable file. Easier to distribute, but might start slightly slower.
    *   `--windowed`: Prevents a console window from appearing when the application runs (essential for GUI apps).
    *   `--add-data="source:destination"` (macOS/Linux uses `:` separator)
    *   `--add-data="source;destination"` (Windows uses `;` separator)
        *   `itcc42:itcc42` or `itcc42;itcc42`: Bundles the *entire* `itcc42` folder and its contents into a folder named `itcc42` inside the executable's temporary runtime directory. This ensures the app can find `itcc42/chud logo.png` and `itcc42/xu logo.png` at runtime.
        *   `settings.json:.` or `settings.json;.`: Bundles the `settings.json` file (if it exists) into the *root* of the executable's temporary runtime directory. The `.` represents the root.
    *   `main.py`: The main entry point script of your application.

4.  **Find the Executable:**
    PyInstaller will create a few folders (`build`, `dist`) and a `.spec` file. Your executable file will be located inside the `dist` folder (e.g., `dist/main` or `dist/main.exe`).

5.  **Run the Executable:**
    Navigate to the `dist` folder and double-click the executable file (`main` or `main.exe`) to run the application.

**Optional PyInstaller Flags:**

*   `--name="YourAppName"`: Sets the name of the executable file.
*   `--icon="path/to/your/icon.ico"` (Windows) or `--icon="path/to/your/icon.icns"` (macOS): Adds a custom icon to the executable.
*   `--hidden-import=module_name`: If PyInstaller fails to detect certain libraries used indirectly (sometimes happens with libraries like `customtkinter` or its dependencies), you might need to explicitly tell it to include them using this flag. Add one flag for each missing module.

## Important Notes

*   **Gmail App Password:** For the email functionality to work, you **must** configure a Gmail "App Password" in the application's settings, not your regular Google account password. See Google's documentation on how to create App Passwords (requires 2-Step Verification).
*   **Footer Images:** The executable relies on the `itcc42` folder being correctly bundled using the `--add-data` flag during the PyInstaller build process.
