@echo off
REM Clean previous builds
rmdir /s /q build dist *.spec

REM Build the new executable
pyinstaller --onefile --noconsole --icon=assets\JTC_logo.ico data_entry_toolset.py

pause
