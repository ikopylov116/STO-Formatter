@echo off
setlocal
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
for /f "delims=" %%i in ('python -c "import sys; print(sys.base_prefix)"') do set "PYTHON_BASE=%%i"
set "TCL_LIBRARY=%PYTHON_BASE%\tcl\tcl8.6"
set "TK_LIBRARY=%PYTHON_BASE%\tcl\tk8.6"
if not exist "%TCL_LIBRARY%" (
  echo ERROR: Tcl directory not found: %TCL_LIBRARY%
  pause
  exit /b 1
)
if not exist "%TK_LIBRARY%" (
  echo ERROR: Tk directory not found: %TK_LIBRARY%
  pause
  exit /b 1
)
python -c "import tkinter; r=tkinter.Tk(); r.withdraw(); r.destroy(); print('Tkinter OK')"
if errorlevel 1 (
  echo ERROR: Tkinter test failed.
  pause
  exit /b 1
)
python -m PyInstaller --noconfirm --clean --onefile --windowed --name STO-Formatter sto_formatter.py
if errorlevel 1 (
  echo ERROR: PyInstaller build failed.
  pause
  exit /b 1
)
echo.
echo Build complete: dist\STO-Formatter.exe
pause
