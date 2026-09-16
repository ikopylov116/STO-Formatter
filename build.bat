@echo off
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --onefile --windowed --name STO-Formatter sto_formatter.py
pause
