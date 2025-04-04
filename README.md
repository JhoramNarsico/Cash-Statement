install first ang tkinter (python GUI)
link: https://www.youtube.com/watch?v=chrXKX4FV-c


para mu run sa vs code, i install sani: 

*pip install tkcalendar

*pip install reportlab

*pip install reportlab python-docx pyinstaller

export to exe. app:

pyinstaller --add-data "credentials.json;." --add-data "logo.png;." --noconfirm --onefile --windowed apicash_statement.py
