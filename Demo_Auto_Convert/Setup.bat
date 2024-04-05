@echo off

REM Define the URL to download the Python installer
set "PYTHON_URL=https://www.python.org/ftp/python/3.10.0/python-3.10.0-amd64.exe"

REM Define the installation parameters
set "INSTALL_PARAMS=/quiet InstallAllUsers=1 PrependPath=1"

REM Define the installation directory
set "INSTALL_DIR=C:\Python310"

REM Download Python installer
echo Downloading Python installer...
curl -o python_installer.exe %PYTHON_URL%

REM Install Python silently
echo Installing Python silently...
python_installer.exe %INSTALL_PARAMS% TargetDir=%INSTALL_DIR%

REM Clean up: Delete the installer file
del python_installer.exe

py -m pip install pdfplumber pandas openpyxl tqdm tkinter 
echo Installation pack complete.

echo Python installation completed.
