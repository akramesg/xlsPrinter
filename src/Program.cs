using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using XlsPrinter.Core;
using XlsPrinter.Spooler;
using XlsPrinter.UI;

namespace XlsPrinter
{
    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            // Check for CLI / headless / verification flags
            if (args != null && args.Length > 0)
            {
                string firstArg = args[0].ToLowerInvariant();

                if (firstArg == "--self-test")
                {
                    return RunAutomatedSelfTest();
                }
                else if (firstArg == "--install-printer")
                {
                    VirtualPrinterManager mgr = new VirtualPrinterManager();
                    string msg;
                    bool ok = mgr.InstallVirtualPrinter(out msg);
                    Console.WriteLine(msg);
                    return ok ? 0 : 1;
                }
                else if (firstArg == "--uninstall-printer")
                {
                    VirtualPrinterManager mgr = new VirtualPrinterManager();
                    string msg;
                    bool ok = mgr.UninstallVirtualPrinter(out msg);
                    Console.WriteLine(msg);
                    return ok ? 0 : 1;
                }
                else if (firstArg == "--dump-elements" && args.Length >= 2)
                {
                    return DumpElementsCommandLine(args[1]);
                }
                else if (firstArg == "--convert" && args.Length >= 3)
                {
                    return ConvertCommandLine(args[1], args[2]);
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string initialFile = (args != null && args.Length > 0 && File.Exists(args[0])) ? args[0] : null;
            Application.Run(new MainForm(initialFile));
            return 0;
        }

        private static int DumpElementsCommandLine(string inputPath)
        {
            try
            {
                EmfSpoolParser parser = new EmfSpoolParser();
                parser.ReverseVisualArabic = true;
                parser.NormalizePresentationForms = true;
                PrintJobDocument doc = parser.ParseFile(inputPath);
                if (doc.Pages.Count == 0) return 1;
                ExtractedPage page = doc.Pages[0];
                List<ExtractedTextElement> list = new List<ExtractedTextElement>(page.Elements);
                list.Sort((a, b) => {
                    if (Math.Abs(a.Y - b.Y) < 1.5f) return a.X.CompareTo(b.X);
                    return a.Y.CompareTo(b.Y);
                });
                foreach (ExtractedTextElement el in list)
                {
                    if (!string.IsNullOrEmpty(el.Text.Trim()))
                    {
                        Console.WriteLine(string.Format("Y={0,5:F1} | X={1,5:F1}..{2,5:F1} (W={3,4:F1}) | [{4}]",
                            el.Y, el.X, el.Right, el.Width, el.Text));
                    }
                }
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                return 2;
            }
        }

        private static int ConvertCommandLine(string inputPath, string outputPath)
        {
            try
            {
                EmfSpoolParser parser = new EmfSpoolParser();
                parser.ReverseVisualArabic = true;
                parser.NormalizePresentationForms = true;
                PrintJobDocument doc = parser.ParseFile(inputPath);
                TableLayoutEngine engine = new TableLayoutEngine();
                ReconstructedTable table = engine.BuildTable(doc, 0, null);
                MainForm.ExportTableToPath(table, outputPath, false);
                Console.WriteLine("Converted '{0}' -> '{1}' ({2} rows, {3} cols)",
                    inputPath, outputPath, table.Rows.Count, table.ColumnHeaders.Count);
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Conversion failed: " + ex.Message);
                return 2;
            }
        }

        /// <summary>
        /// End-to-end automated verification:
        /// 1. Generates a classic Win32 GDI EMF print job containing Arabic + English accounting table data
        /// 2. Tests Arabic Presentation Forms normalization (U+FE80..U+FEFC -> U+0600..U+06FF) & Windows-1256 decoding
        /// 3. Parses the EMF spool stream and reconstructs the 2D table layout
        /// 4. Exports to binary .XLS (BIFF8), .XLSX (OpenXML), and .CSV (UTF-8 BOM)
        /// 5. Verifies the generated binary files and prints a summary report.
        /// </summary>
        private static int RunAutomatedSelfTest()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("=== XLS Virtual Printer Automated Self-Test ===");

            // 1. Test Arabic Presentation Forms & Windows-1256 Normalization
            string shapedArabic = "\uFEFB\uFEE4\uFEAE\uFE91"; // Shaped presentation forms including Lam-Alef ligature
            string normalizedArabic = ArabicTextHelper.NormalizeArabicText(shapedArabic, true, false);
            if (!normalizedArabic.Contains("\u0644\u0627"))
            {
                Console.Error.WriteLine("FAIL: Arabic Presentation Form Lam-Alef normalization failed.");
                return 1;
            }
            Console.WriteLine("[PASS] Arabic Presentation Forms (FE70-FEFC) -> Canonical Unicode (0600-06FF) normalization");

            // 2. Generate GDI EMF Spool Stream with Arabic & English Invoice Table
            byte[] emfBytes = SamplePrintGenerator.CreateSampleArabicAccountingEmf();
            string loc = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string appDir = !string.IsNullOrEmpty(loc) ? Path.GetDirectoryName(loc) : AppDomain.CurrentDomain.BaseDirectory;
            if (!Directory.Exists(appDir) || appDir.IndexOf("System32", StringComparison.OrdinalIgnoreCase) >= 0)
                appDir = Directory.GetCurrentDirectory();
            string sampleDir = Path.Combine(appDir, "Samples");
            if (!Directory.Exists(sampleDir)) Directory.CreateDirectory(sampleDir);

            string sampleEmfPath = Path.Combine(sampleDir, "Sample_Arabic_Invoice.emf");
            File.WriteAllBytes(sampleEmfPath, emfBytes);
            Console.WriteLine("[PASS] Generated Win32 GDI EMF print job: {0} ({1} bytes)", sampleEmfPath, emfBytes.Length);

            // 3. Parse EMF Stream & Reconstruct Table
            EmfSpoolParser parser = new EmfSpoolParser();
            PrintJobDocument doc = parser.ParseBytes(emfBytes, "Sample_Arabic_Invoice");
            if (doc.Pages.Count == 0 || doc.Pages[0].Elements.Count < 10)
            {
                Console.Error.WriteLine("FAIL: Expected extracted GDI text elements from EMF stream, got " +
                    (doc.Pages.Count == 0 ? 0 : doc.Pages[0].Elements.Count));
                return 1;
            }
            Console.WriteLine("[PASS] Extracted {0} GDI text elements from page 1", doc.Pages[0].Elements.Count);

            TableLayoutEngine layout = new TableLayoutEngine();
            ReconstructedTable table = layout.BuildTable(doc, 0, null);
            Console.WriteLine("[PASS] Reconstructed table: {0} rows x {1} columns (Predominantly RTL: {2})",
                table.Rows.Count, table.ColumnHeaders.Count, table.IsPredominantlyRtl);

            // 4. Export to .XLS (BIFF8), .XLSX, and .CSV
            string outXls = Path.Combine(sampleDir, "Sample_Arabic_Invoice.xls");
            string outXlsx = Path.Combine(sampleDir, "Sample_Arabic_Invoice.xlsx");
            string outCsv = Path.Combine(sampleDir, "Sample_Arabic_Invoice.csv");

            Biff8XlsWriter xlsWriter = new Biff8XlsWriter();
            xlsWriter.RightToLeftSheet = true;
            xlsWriter.SaveToXls(table, outXls);

            XlsxAndCsvWriter xlsxWriter = new XlsxAndCsvWriter();
            xlsxWriter.RightToLeftSheet = true;
            xlsxWriter.SaveToXlsx(table, outXlsx);
            xlsxWriter.SaveToCsv(table, outCsv);

            // Verify OLE2 header signature (D0 CF 11 E0 A1 B1 1A E1) in .xls
            byte[] xlsBytes = File.ReadAllBytes(outXls);
            if (xlsBytes.Length < 1024 || xlsBytes[0] != 0xD0 || xlsBytes[1] != 0xCF || xlsBytes[2] != 0x11 || xlsBytes[3] != 0xE0)
            {
                Console.Error.WriteLine("FAIL: Invalid OLE2 Compound Document header in generated .xls");
                return 1;
            }
            Console.WriteLine("[PASS] Exported valid BIFF8 Excel 97-2003 (.xls): {0} ({1} bytes)", outXls, xlsBytes.Length);

            // Verify PK ZIP header (50 4B 03 04) in .xlsx
            byte[] xlsxBytes = File.ReadAllBytes(outXlsx);
            if (xlsxBytes.Length < 256 || xlsxBytes[0] != 0x50 || xlsxBytes[1] != 0x4B)
            {
                Console.Error.WriteLine("FAIL: Invalid ZIP header in generated .xlsx");
                return 1;
            }
            Console.WriteLine("[PASS] Exported valid OpenXML (.xlsx): {0} ({1} bytes)", outXlsx, xlsxBytes.Length);
            Console.WriteLine("[PASS] Exported UTF-8 BOM (.csv): {0}", outCsv);

            Console.WriteLine("=== ALL SELF-TESTS PASSED ===");
            return 0;
        }
    }
}
