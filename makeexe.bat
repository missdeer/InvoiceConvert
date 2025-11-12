d:\Python313\Scripts\pyinstaller.exe --onefile --clean main.py --icon=%~dp0\invoice.ico --exclude-module=tkinter --exclude-module=test --exclude-module=unittest --exclude-module=setuptools --exclude-module=pydoc
ren main.exe InvoiceConvert.exe
