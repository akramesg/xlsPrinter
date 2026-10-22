# XLS Virtual Printer (Win32 GDI Print-to-Excel Converter)

A 32-bit (`x86`) Windows Virtual Printer and Spooler-to-Excel converter designed for **Windows XP, Windows Server 2008, Windows 7, and Windows 10/11**. It intercepts standard Win32 GDI print jobs (`.SPL` / `.EMF` / `.PRN`) from classic accounting, ERP, or desktop systems, reconstructs rows and columns, properly normalizes **Arabic characters**, and lets you preview, filter/select rows, columns, or page regions before exporting to **Excel 97–2003 (`.xls` BIFF8)**, **Excel OpenXML (`.xlsx`)**, or **UTF-8 `.csv`**.

---

## Key Features

1. **Zero Kernel Driver / Zero WDK Required (`winspool.drv` + `NT EMF 1.003`)**
   - Uses Windows' built-in `winprint` print processor and EMF spooling (`PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS`) to capture 100% of GDI text coordinates (`EMR_EXTTEXTOUTW`, `EMR_EXTTEXTOUTA`, `EMR_SMALLTEXTOUT`) directly from the Windows Spooler queue.
   - Works on 32-bit Windows XP SP2+, Server 2008, Windows 7 x86, and 64-bit Windows without requiring kernel driver signing.
2. **Full Arabic Character & Right-to-Left (RTL) Support**
   - **UTF-16LE Unicode** (`EMR_EXTTEXTOUTW`) and **Windows-1256 Arabic ANSI** (`EMR_EXTTEXTOUTA` with `ARABIC_CHARSET = 178`) decoding.
   - **Arabic Presentation Forms Normalization**: Automatically maps shaped glyph forms (`U+FE70`–`U+FEFC` and `U+FB50`–`U+FDFF`, including `لا` Lam-Alef ligatures) back to standard searchable/sortable Unicode Arabic (`U+0600`–`U+06FF`).
   - **Visual-to-Logical Arabic Reversal**: Optional toggle for legacy DOS/Win32 apps that emit Arabic characters in visual left-to-right order while preserving embedded numbers and dates.
   - **RTL Excel Sheet View**: Sets the BIFF8 `WINDOW2` (`0x0040` `fRightToLeft`) and `XF` reading-order flags so Excel opens the sheet in Right-to-Left Arabic layout.
3. **Interactive Visual Page Preview & Data Selection UI**
   - **Rubberband Area Selection**: Drag a rectangle directly on the visual page preview to isolate a specific table or region on the printed page (e.g., excluding page headers/footers).
   - **Row & Column Selection**: Check/uncheck individual rows or columns, or edit cell values in the grid before saving.
4. **Native `.XLS` (BIFF8), `.XLSX`, and `.CSV` Export**
   - Self-contained binary **Excel 97–2003 `.xls` (BIFF8 inside OLE2 Compound Document)** writer — does **not** require Microsoft Office or Excel to be installed on the machine.

---

## Building from Source

Run [`build.bat`](file:///c:/Dev/xlsPrinter/build.bat) from any command prompt (uses the built-in .NET Framework 2.0/3.5/4.0 compiler in `C:\Windows\Microsoft.NET\Framework`):

```bat
build.bat
```

This produces the standalone 32-bit executable:
- `bin\XlsVirtualPrinter.exe`

---

## Usage

### 1. Interactive GUI Mode (Watcher + Preview + XLS Export)
Run `bin\XlsVirtualPrinter.exe`:
- Click **Install Virtual Printer** (run as Administrator once) to register **"XLS Virtual Printer"** in Windows Printers & Faxes.
- Print from your classic Win32 application to **XLS Virtual Printer**.
- The preview window automatically pops up with your captured print job:
  - Drag a box on the **left page preview** to filter a rectangular table area.
  - Check/uncheck columns or rows on the **right grid**.
  - Click **Export Selected to .XLS (97-2003)**, **Export to .XLSX**, or **Export to .CSV**.

### 2. Command-Line Mode
```bat
:: Install the virtual printer (run as Administrator)
bin\XlsVirtualPrinter.exe --install-printer

:: Uninstall the virtual printer
bin\XlsVirtualPrinter.exe --uninstall-printer

:: Convert an existing .SPL / .EMF / .PRN file directly from CLI
bin\XlsVirtualPrinter.exe --convert "input.spl" "output.xls"

:: Run the automated Arabic & BIFF8 self-test suite
bin\XlsVirtualPrinter.exe --self-test
```
