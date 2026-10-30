@echo off
setlocal
echo ============================================================
echo Building XLS Virtual Printer (32-bit x86 Win32 Compatible)
echo Target OS: Windows XP / Server 2008 / Windows 7 / 10 / 11
echo ============================================================

if not exist "bin" mkdir "bin"

set CSC40=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
set CSC35=C:\Windows\Microsoft.NET\Framework\v3.5\csc.exe
set CSC20=C:\Windows\Microsoft.NET\Framework\v2.0.50727\csc.exe

if exist "%CSC40%" (
    set CSC=%CSC40%
) else if exist "%CSC35%" (
    set CSC=%CSC35%
) else (
    set CSC=%CSC20%
)

echo Using Compiler: %CSC%

"%CSC%" /nologo /codepage:65001 /platform:x86 /target:winexe /optimize+ /out:"bin\XlsVirtualPrinter.exe" ^
    /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Data.dll /r:System.Xml.dll ^
    "src\Core\AppLogger.cs" ^
    "src\Core\PrintModels.cs" ^
    "src\Core\ArabicTextHelper.cs" ^
    "src\Core\EmfSpoolParser.cs" ^
    "src\Core\TableLayoutEngine.cs" ^
    "src\Core\Biff8XlsWriter.cs" ^
    "src\Core\XlsxAndCsvWriter.cs" ^
    "src\Core\SamplePrintGenerator.cs" ^
    "src\Spooler\VirtualPrinterManager.cs" ^
    "src\UI\MainForm.cs" ^
    "src\Program.cs"

if errorlevel 1 (
    echo [ERROR] Build failed!
    exit /b 1
)

copy /Y "src\App.config" "bin\XlsVirtualPrinter.exe.config" >nul
echo [SUCCESS] Built bin\XlsVirtualPrinter.exe (32-bit x86)

if exist "\\mailserver\Shared\akram" (
    copy /Y "bin\XlsVirtualPrinter.exe" "\\mailserver\Shared\akram\XlsVirtualPrinter.exe" >nul
    copy /Y "bin\XlsVirtualPrinter.exe.config" "\\mailserver\Shared\akram\XlsVirtualPrinter.exe.config" >nul
    echo [SYNC] Updated \\mailserver\Shared\akram\XlsVirtualPrinter.exe
)
