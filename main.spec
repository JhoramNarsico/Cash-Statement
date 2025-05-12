# -*- mode: python ; coding: utf-8 -*-

# Import necessary PyInstaller modules
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Define the path to your application's assets (logos, etc.)
# This makes it easier to manage if you have many assets
assets_path = './' # Assuming logos are in the same directory as main.py
                  # Or specify a subfolder like './assets/'

block_cipher = None

a = Analysis(
    ['main.py'],                               # Your main script(s)
    pathex=[],                                 # Additional search paths for imports (usually not needed if your project is structured well)
    binaries=[],                               # List of non-Python libraries (e.g., .dll, .so) if PyInstaller misses them
    datas=[
        (assets_path + 'chud logo.png', '.'),  # (source, destination_in_bundle)
        (assets_path + 'xu logo.png', '.'),
        # Add any other data files your application needs (e.g., default settings if you bundled one)
        # Example: ('path/to/your/app_icon.ico', '.') # To bundle an icon
        # If you have an 'images' folder:
        # ('images/another_logo.png', 'images') # This would put it in an 'images' subfolder in the bundle
    ],
    hiddenimports=[
        'babel.numbers',                      # tkcalendar dependency (often needed as hidden)
        'PIL._tkinter_finder',                # Pillow/Tkinter integration
        'customtkinter',                      # Explicitly add if any issues
        'reportlab', 'reportlab.graphics',    # Ensure all major parts of reportlab are found
        'docx',
        'pdfplumber',
        'camelot',
        'fitz', 'fitz_old',                   # PyMuPDF (fitz)
        'tkinter', 'tkcalendar',
        # Add more here if PyInstaller reports "ModuleNotFound" for specific modules during runtime
        # For example, sometimes specific backends for libraries are missed:
        # 'pandas._libs.tslibs.base', # Example if you were using pandas and it missed something
        # 'scipy.special._cdflib'     # Example
    ],
    hookspath=[],                              # Paths to custom PyInstaller hooks (advanced)
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],                               # Modules to explicitly exclude (e.g., to reduce size, if you know they are not needed)
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Collect data files from libraries if needed (example for tkcalendar, might be needed for others)
# This automatically finds data files that a library itself might need.
# a.datas += collect_data_files('tkcalendar')
# a.datas += collect_data_files('customtkinter') # If CTk has assets it needs bundled

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],                                        # Files to be bundled with the executable but not part of PYZ
    exclude_binaries=True,                     # Keep this True for --onedir
    name='CashFlowApp',                        # Name of the executable
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,                                  # Use UPX for compression (if installed and desired, can reduce size)
    upx_exclude=[],
    runtime_tmpdir=None,                       # In --onedir, this is not used as extensively as in --onefile
    console=False,                             # For GUI applications (same as --windowed)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico'    # Specify your application icon (Windows .ico, macOS .icns)
                                               # Ensure 'your_app_icon.ico' is in your assets_path
)

# This part is crucial for --onedir mode. It bundles everything into a directory.
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CashFlowApp',                        # Name of the output directory in 'dist'
)