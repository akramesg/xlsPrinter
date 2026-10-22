@echo off
setlocal
echo ============================================================
echo Building XLS Virtual Printer (32-bit x86 Win32 Compatible)
echo Target OS: Windows XP / Server 2008 / Windows 7 / 10 / 11
echo ============================================================

if not exist "bin" mkdir "bin"

set CSC35=C:\Windows\Microsoft.NET\Framework\v3.5\csc.exe
set CSC20=C:\Windows\Microsoft.NET\Framework\v2.0.50727\csc.exe
set CSC40=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe

if exist "%CSC35%" (
    set CSC=%CSC35%
) else if exist "%CSC20%" (
    set CSC=%CSC20%
) else (
    set CSC=%CSC40%
)

echo Using Compiler: %CSC%

"%CSC%" /nologo /platform:x86 /target:exe /optimize+ /out:"bin\XlsVirtualPrinter.exe" ^
    /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Data.dll /r:System.Xml.dll ^
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
exit /b 0
