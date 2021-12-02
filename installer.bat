::pyinstaller --distpath installer --paths=venv/Lib/site-packages/openpyxl/ --hidden-import openpyxl --clean main.py
::pyinstaller --distpath installer --clean main.py
pyinstaller --distpath installer --onefile --clean main.py




















