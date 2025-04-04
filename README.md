python libraries:
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client reportlab python-docx


installer:

pyinstaller --add-data "credentials.json;." --add-data "logo.png;." --noconfirm --onefile --windowed apicash_statement.py
