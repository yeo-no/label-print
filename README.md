# label-print
python

# build 명령어
pyinstaller --onefile --noconsole --icon pepeIcon.ico --hidden-import xlrd --hidden-import openpyxl --hidden-import pandas --hidden-import fpdf --hidden-import qrcode --collect-submodules pandas smartstore_label_print.py

