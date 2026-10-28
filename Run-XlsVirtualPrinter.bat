@echo off
REM Launches XLS Virtual Printer in the background without keeping a console window open.
cd /d "%~dp0"
start "" "%~dp0bin\XlsVirtualPrinter.exe" %*
exit

