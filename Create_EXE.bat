@echo off
REM Compila lo script Python in un eseguibile con PyInstaller
SET PYINSTALLER_PATH="C:\Users\e0840680\AppData\Roaming\Python\Python310\Scripts\pyinstaller.exe"
SET SCRIPT_NAME=QT_ITA-Offerta_Xmoem_v5.py
SET IMAGE_FILE=Eaton-Logo.jpg

%PYINSTALLER_PATH% --onefile --add-data "%IMAGE_FILE%;." %SCRIPT_NAME%
pause
