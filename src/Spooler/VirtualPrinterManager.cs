using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32;

namespace XlsPrinter.Spooler
{
    public class PrintJobCapturedEventArgs : EventArgs
    {
        public int JobId { get; set; }
        public string DocumentTitle { get; set; }
        public string SpoolFilePath { get; set; }
        public byte[] SpoolBytes { get; set; }
    }

    /// <summary>
    /// Installs, configures, and monitors the "XLS Virtual Printer" on Windows XP, Server 2008, Win7, and Win10/11 (x86/x64).
    /// Uses native Win32 winspool.drv APIs and configures the printer to:
    /// 1. Spool in NT EMF format (so all GDI text coordinates & Unicode/Arabic glyphs are preserved)
    /// 2. Keep Printed Jobs (PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS) so the .SPL file remains safely on disk until captured.
    /// </summary>
    public class VirtualPrinterManager : IDisposable
    {
        public const string DefaultPrinterName = "XLS Virtual Printer";
        public const string DefaultPortName = "NUL:";

        // Win32 Spooler Constants
        private const int PRINTER_ACCESS_ADMINISTER = 0x00000004;
        private const int PRINTER_ACCESS_USE = 0x00000008;
        private const int STANDARD_RIGHTS_REQUIRED = 0x000F0000;
        private const int PRINTER_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED | PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE;

        private const uint PRINTER_ATTRIBUTE_LOCAL = 0x00000040;
        private const uint PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST = 0x00000200;
        private const uint PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS = 0x00000100;

        private const int JOB_CONTROL_DELETE = 5;

        private const uint JOB_STATUS_SPOOLING = 0x00000008;
        private const uint JOB_STATUS_PRINTED = 0x00000080;
        private const uint JOB_STATUS_COMPLETE = 0x00001000;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct PRINTER_DEFAULTS
        {
            public string pDatatype;
            public IntPtr pDevMode;
            public int DesiredAccess;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct PRINTER_INFO_2
        {
            public string pServerName;
            public string pPrinterName;
            public string pShareName;
            public string pPortName;
            public string pDriverName;
            public string pComment;
            public string pLocation;
            public IntPtr pDevMode;
            public string pSepFile;
            public string pPrintProcessor;
            public string pDatatype;
            public string pParameters;
            public IntPtr pSecurityDescriptor;
            public uint Attributes;
            public uint Priority;
            public uint DefaultPriority;
            public uint StartTime;
            public uint UntilTime;
            public uint Status;
            public uint cJobs;
            public uint AveragePPM;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct JOB_INFO_1
        {
            public uint JobId;
            public string pPrinterName;
            public string pMachineName;
            public string pUserName;
            public string pDocument;
            public string pDatatype;
            public string pStatus;
            public uint Status;
            public uint Priority;
            public uint Position;
            public uint TotalPages;
            public uint PagesPrinted;
            public SYSTEMTIME Submitted;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SYSTEMTIME
        {
            public ushort wYear;
            public ushort wMonth;
            public ushort wDayOfWeek;
            public ushort wDay;
            public ushort wHour;
            public ushort wMinute;
            public ushort wSecond;
            public ushort wMilliseconds;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct DRIVER_INFO_1
        {
            public string pName;
        }

        [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, ref PRINTER_DEFAULTS pDefault);

        [DllImport("winspool.drv", SetLastError = true)]
        private static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr AddPrinter(string pName, uint Level, ref PRINTER_INFO_2 pPrinter);

        [DllImport("winspool.drv", SetLastError = true)]
        private static extern bool DeletePrinter(IntPtr hPrinter);

        [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool EnumJobs(IntPtr hPrinter, uint FirstJob, uint NoJobs, uint Level, IntPtr pJob, uint cbBuf, out uint pcbNeeded, out uint pcReturned);

        [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetJob(IntPtr hPrinter, uint JobId, uint Level, IntPtr pJob, uint Command);

        [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool EnumPrinterDrivers(string pName, string pEnvironment, uint Level, IntPtr pDriverInfo, uint cbBuf, out uint pcbNeeded, out uint pcReturned);

        [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool EnumPorts(string pName, uint Level, IntPtr pPorts, uint cbBuf, out uint pcbNeeded, out uint pcReturned);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct PORT_INFO_1
        {
            public string pName;
        }

        public string PrinterName { get; set; }
        public event EventHandler<PrintJobCapturedEventArgs> JobCaptured;

        private Thread _watcherThread;
        private volatile bool _isRunning;
        private readonly Dictionary<int, DateTime> _processedJobIds = new Dictionary<int, DateTime>();
        private readonly Dictionary<string, DateTime> _processedFiles = new Dictionary<string, DateTime>();

        public VirtualPrinterManager()
        {
            PrinterName = DefaultPrinterName;
        }

        /// <summary>
        /// Checks if the virtual printer is currently installed in Windows.
        /// </summary>
        public bool IsPrinterInstalled()
        {
            PRINTER_DEFAULTS defs = new PRINTER_DEFAULTS();
            defs.DesiredAccess = PRINTER_ACCESS_USE;
            IntPtr hPrinter;
            if (OpenPrinter(PrinterName, out hPrinter, ref defs))
            {
                ClosePrinter(hPrinter);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Installs "XLS Virtual Printer" using an existing inbox GDI driver (such as Microsoft Shared Fax Driver,
        /// Generic / Text Only, HP LaserJet, or Microsoft Print to PDF/XPS) configured with winprint / NT EMF 1.003
        /// and PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS.
        /// </summary>
        public bool InstallVirtualPrinter(out string statusMessage)
        {
            if (IsPrinterInstalled())
            {
                EnsureEmfSpoolSettingsInRegistry(PrinterName);
                statusMessage = "Virtual printer '" + PrinterName + "' is already installed and configured for EMF capture.";
                return true;
            }

            string driverName = FindBestInstalledGdiDriver();
            if (string.IsNullOrEmpty(driverName))
            {
                // On Windows XP with zero printer drivers installed, attempt to install the inbox "Generic / Text Only" driver
                bool installedGeneric = TryInstallInboxGenericDriver();
                if (installedGeneric)
                {
                    driverName = FindBestInstalledGdiDriver();
                }
            }

            if (string.IsNullOrEmpty(driverName))
            {
                statusMessage = "No printer drivers are installed on this Windows XP machine.\n\n" +
                                "To fix this:\n" +
                                "1. Open Windows 'Printers and Faxes' -> 'Add a printer'\n" +
                                "2. Choose 'Local printer attached to this computer' (uncheck automatically detect)\n" +
                                "3. Port: FILE: or LPT1:\n" +
                                "4. Manufacturer: 'Generic' -> Printer: 'Generic / Text Only'\n" +
                                "5. Finish, then click 'Install Virtual Printer' again.";
                return false;
            }

            // Find a valid, existing local port (NUL:, FILE:, or LPT1:)
            string portName = FindBestAvailablePort();

            PRINTER_INFO_2 pi2 = new PRINTER_INFO_2();
            pi2.pServerName = null;
            pi2.pPrinterName = PrinterName;
            pi2.pPortName = portName;
            pi2.pDriverName = driverName;
            pi2.pComment = "Virtual Print-to-XLS Driver (Arabic & Win32 GDI EMF)";
            pi2.pLocation = "Local Virtual XLS Port";
            pi2.pPrintProcessor = "winprint";
            pi2.pDatatype = "NT EMF 1.003";
            pi2.Attributes = PRINTER_ATTRIBUTE_LOCAL | PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS | PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST;
            pi2.Priority = 1;
            pi2.DefaultPriority = 1;

            IntPtr hPrinter = AddPrinter(null, 2, ref pi2);
            if (hPrinter == IntPtr.Zero)
            {
                int err = Marshal.GetLastWin32Error();

                // If NT EMF 1.003 is rejected during AddPrinter by some XP drivers, retry with RAW
                pi2.pDatatype = "RAW";
                hPrinter = AddPrinter(null, 2, ref pi2);

                if (hPrinter == IntPtr.Zero)
                {
                    // If still failed, try fallback port "FILE:"
                    if (portName != "FILE:")
                    {
                        pi2.pPortName = "FILE:";
                        hPrinter = AddPrinter(null, 2, ref pi2);
                    }
                }

                if (hPrinter == IntPtr.Zero)
                {
                    string errExplanation = GetFriendlyErrorMessage(err, driverName, portName);
                    statusMessage = "Failed to add printer '" + PrinterName + "' (Win32 Error " + err + "):\n\n" + errExplanation;
                    return false;
                }
            }

            ClosePrinter(hPrinter);
            EnsureEmfSpoolSettingsInRegistry(PrinterName);

            statusMessage = "Installed '" + PrinterName + "' successfully using driver '" + driverName + "' on port '" + pi2.pPortName + "'.";
            return true;
        }

        private static string GetFriendlyErrorMessage(int err, string driverName, string portName)
        {
            if (err == 5) // ERROR_ACCESS_DENIED
            {
                return "Access Denied. Windows requires Administrator privileges to register printer drivers and ports.";
            }
            if (err == 1797) // ERROR_UNKNOWN_PRINTER_DRIVER
            {
                return "The printer driver '" + driverName + "' was not accepted by the Windows Spooler on this system.\n" +
                       "Please add 'Generic / Text Only' from Windows Printers & Faxes first.";
            }
            if (err == 1796) // ERROR_UNKNOWN_PORT
            {
                return "The port '" + portName + "' is not recognized by the Windows Spooler.\n" +
                       "Make sure 'FILE:' or 'LPT1:' is available in your Ports list.";
            }
            if (err == 1802) // ERROR_PRINTER_ALREADY_EXISTS
            {
                return "A printer named '" + DefaultPrinterName + "' already exists in Windows.";
            }
            try
            {
                return new Win32Exception(err).Message;
            }
            catch
            {
                return "Spooler returned error code " + err;
            }
        }

        private static bool TryInstallInboxGenericDriver()
        {
            try
            {
                string windir = Environment.GetEnvironmentVariable("windir") ?? @"C:\WINDOWS";
                string infPath = Path.Combine(windir, @"inf\ntprint.inf");
                if (!File.Exists(infPath))
                    return false;

                // Uses rundll32 printui.dll,PrintUIEntry to install the built-in Windows Generic / Text Only driver from ntprint.inf
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = "rundll32.exe";
                psi.Arguments = "printui.dll,PrintUIEntry /ia /m \"Generic / Text Only\" /f \"" + infPath + "\" /q";
                psi.CreateNoWindow = true;
                psi.UseShellExecute = false;

                using (Process proc = Process.Start(psi))
                {
                    proc.WaitForExit(10000);
                    return proc.ExitCode == 0;
                }
            }
            catch
            {
                return false;
            }
        }

        public bool UninstallVirtualPrinter(out string statusMessage)
        {
            PRINTER_DEFAULTS defs = new PRINTER_DEFAULTS();
            defs.DesiredAccess = PRINTER_ALL_ACCESS;
            IntPtr hPrinter;
            if (!OpenPrinter(PrinterName, out hPrinter, ref defs))
            {
                statusMessage = "Printer '" + PrinterName + "' is not installed or requires Administrator privileges.";
                return false;
            }

            bool ok = DeletePrinter(hPrinter);
            int err = Marshal.GetLastWin32Error();
            ClosePrinter(hPrinter);

            if (!ok)
            {
                statusMessage = "Failed to remove printer (Win32 Error " + err + "). Run as Administrator.";
                return false;
            }

            statusMessage = "Uninstalled '" + PrinterName + "' successfully.";
            return true;
        }

        /// <summary>
        /// Enumerates already-installed printer drivers on the machine and picks the best GDI driver
        /// that supports EMF spooling (standard on Windows XP, 7, Server 2008, 10, 11).
        /// </summary>
        public string FindBestInstalledGdiDriver()
        {
            List<string> drivers = GetInstalledDrivers();

            // Preferred GDI drivers that spool clean NT EMF 1.00x records without popping up dialogs
            // NOTE: Do NOT use "Fax" or "Microsoft Shared Fax Driver" as it intercepts print jobs and opens Windows Fax Wizard!
            string[] preferredPatterns = new string[]
            {
                "Generic / Text Only",
                "HP LaserJet",
                "HP Color LaserJet",
                "Epson",
                "Canon",
                "Generic IBM Graphics",
                "Microsoft Print to PDF",
                "Microsoft XPS Document Writer",
                "UniDrv"
            };

            for (int p = 0; p < preferredPatterns.Length; p++)
            {
                string pat = preferredPatterns[p];
                for (int i = 0; i < drivers.Count; i++)
                {
                    if (drivers[i].IndexOf(pat, StringComparison.OrdinalIgnoreCase) >= 0 &&
                        drivers[i].IndexOf("Fax", StringComparison.OrdinalIgnoreCase) < 0)
                        return drivers[i];
                }
            }

            for (int i = 0; i < drivers.Count; i++)
            {
                if (drivers[i].IndexOf("Fax", StringComparison.OrdinalIgnoreCase) < 0)
                    return drivers[i];
            }

            if (drivers.Count > 0)
                return drivers[0];

            return null;
        }

        public List<string> GetInstalledDrivers()
        {
            List<string> result = new List<string>();
            uint cbNeeded = 0, cReturned = 0;
            EnumPrinterDrivers(null, null, 1, IntPtr.Zero, 0, out cbNeeded, out cReturned);
            if (cbNeeded == 0) return result;

            IntPtr pAddr = Marshal.AllocHGlobal((int)cbNeeded);
            try
            {
                if (EnumPrinterDrivers(null, null, 1, pAddr, cbNeeded, out cbNeeded, out cReturned))
                {
                    IntPtr current = pAddr;
                    int structSize = Marshal.SizeOf(typeof(DRIVER_INFO_1));
                    for (int i = 0; i < cReturned; i++)
                    {
                        DRIVER_INFO_1 di = (DRIVER_INFO_1)Marshal.PtrToStructure(current, typeof(DRIVER_INFO_1));
                        if (!string.IsNullOrEmpty(di.pName))
                            result.Add(di.pName);
                        current = new IntPtr(current.ToInt64() + structSize);
                    }
                }
            }
            finally
            {
                Marshal.FreeHGlobal(pAddr);
            }
            return result;
        }

        public List<string> GetInstalledPorts()
        {
            List<string> result = new List<string>();
            uint cbNeeded = 0, cReturned = 0;
            EnumPorts(null, 1, IntPtr.Zero, 0, out cbNeeded, out cReturned);
            if (cbNeeded == 0) return result;

            IntPtr pAddr = Marshal.AllocHGlobal((int)cbNeeded);
            try
            {
                if (EnumPorts(null, 1, pAddr, cbNeeded, out cbNeeded, out cReturned))
                {
                    IntPtr current = pAddr;
                    int structSize = Marshal.SizeOf(typeof(PORT_INFO_1));
                    for (int i = 0; i < cReturned; i++)
                    {
                        PORT_INFO_1 pi = (PORT_INFO_1)Marshal.PtrToStructure(current, typeof(PORT_INFO_1));
                        if (!string.IsNullOrEmpty(pi.pName))
                            result.Add(pi.pName);
                        current = new IntPtr(current.ToInt64() + structSize);
                    }
                }
            }
            finally
            {
                Marshal.FreeHGlobal(pAddr);
            }
            return result;
        }

        private string FindBestAvailablePort()
        {
            List<string> ports = GetInstalledPorts();

            // Check if NUL: or nul: exists
            for (int i = 0; i < ports.Count; i++)
            {
                if (ports[i].Equals("NUL:", StringComparison.OrdinalIgnoreCase))
                    return ports[i];
            }

            // Check if FILE: exists
            for (int i = 0; i < ports.Count; i++)
            {
                if (ports[i].Equals("FILE:", StringComparison.OrdinalIgnoreCase))
                    return ports[i];
            }

            // Check if LPT1: exists
            for (int i = 0; i < ports.Count; i++)
            {
                if (ports[i].Equals("LPT1:", StringComparison.OrdinalIgnoreCase))
                    return ports[i];
            }

            if (ports.Count > 0)
                return ports[0];

            return "FILE:";
        }

        private void EnsureEmfSpoolSettingsInRegistry(string printerName)
        {
            try
            {
                string regPath = @"SYSTEM\CurrentControlSet\Control\Print\Printers\" + printerName;
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(regPath, true))
                {
                    if (key != null)
                    {
                        key.SetValue("Print Processor", "winprint", RegistryValueKind.String);
                        key.SetValue("Datatype", "NT EMF 1.003", RegistryValueKind.String);
                        // Ensure KEEPPRINTEDJOBS (0x100) is set on Attributes
                        object attrObj = key.GetValue("Attributes");
                        if (attrObj is int)
                        {
                            int attr = (int)attrObj;
                            attr |= (int)PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS;
                            // Clear PRINTER_ATTRIBUTE_ENABLE_DEVQ or RAW-only flags if set
                            key.SetValue("Attributes", attr, RegistryValueKind.DWord);
                        }
                    }
                }
            }
            catch
            {
                // Ignore if running without Admin rights
            }
        }

        /// <summary>
        /// Starts background monitoring of the Windows Spooler queue and the %SystemRoot%\System32\spool\PRINTERS directory.
        /// </summary>
        public void StartWatching()
        {
            if (_isRunning) return;
            _isRunning = true;
            _watcherThread = new Thread(WatcherLoop);
            _watcherThread.IsBackground = true;
            _watcherThread.Name = "XlsSpoolWatcher";
            _watcherThread.Start();
        }

        public void StopWatching()
        {
            _isRunning = false;
        }

        private void WatcherLoop()
        {
            string spoolFolder = GetSpoolDirectory();

            while (_isRunning)
            {
                try
                {
                    PollPrinterJobs(spoolFolder);
                }
                catch
                {
                    // Keep watcher alive
                }
                Thread.Sleep(600);
            }
        }

        public static string GetSpoolDirectory()
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Print\Printers", false))
                {
                    if (key != null)
                    {
                        string dir = key.GetValue("DefaultSpoolDirectory") as string;
                        if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                            return dir;
                    }
                }
            }
            catch { }

            string sysRoot = Environment.GetEnvironmentVariable("SystemRoot") ?? @"C:\Windows";
            return Path.Combine(sysRoot, @"System32\spool\PRINTERS");
        }

        private void PollPrinterJobs(string spoolFolder)
        {
            PRINTER_DEFAULTS defs = new PRINTER_DEFAULTS();
            defs.DesiredAccess = PRINTER_ACCESS_USE | PRINTER_ACCESS_ADMINISTER;
            IntPtr hPrinter;
            if (!OpenPrinter(PrinterName, out hPrinter, ref defs))
            {
                defs.DesiredAccess = PRINTER_ACCESS_USE;
                if (!OpenPrinter(PrinterName, out hPrinter, ref defs))
                    return;
            }

            try
            {
                uint cbNeeded = 0, cReturned = 0;
                EnumJobs(hPrinter, 0, 64, 1, IntPtr.Zero, 0, out cbNeeded, out cReturned);
                if (cbNeeded == 0) return;

                IntPtr pJobBuf = Marshal.AllocHGlobal((int)cbNeeded);
                try
                {
                    if (EnumJobs(hPrinter, 0, 64, 1, pJobBuf, cbNeeded, out cbNeeded, out cReturned))
                    {
                        int structSize = Marshal.SizeOf(typeof(JOB_INFO_1));
                        IntPtr current = pJobBuf;
                        for (int i = 0; i < cReturned; i++)
                        {
                            JOB_INFO_1 job = (JOB_INFO_1)Marshal.PtrToStructure(current, typeof(JOB_INFO_1));
                            current = new IntPtr(current.ToInt64() + structSize);

                            int jobId = (int)job.JobId;
                            if (_processedJobIds.ContainsKey(jobId))
                                continue;

                            // Wait until job finishes spooling
                            if ((job.Status & JOB_STATUS_SPOOLING) != 0)
                                continue;

                            XlsPrinter.Core.AppLogger.Log(string.Format("SpoolWatcher: Detected new print job #{0}: '{1}'", jobId, job.pDocument));

                            string splPath = FindSpoolFileForJob(spoolFolder, jobId);
                            if (!string.IsNullOrEmpty(splPath) && File.Exists(splPath))
                            {
                                byte[] bytes = TryReadSharedFile(splPath);
                                if (bytes != null && bytes.Length > 0)
                                {
                                    _processedJobIds[jobId] = DateTime.Now;
                                    XlsPrinter.Core.AppLogger.Log(string.Format("SpoolWatcher: Read {0} bytes from {1}", bytes.Length, splPath));

                                    // Delete the completed job from the queue so it doesn't clutter the spooler
                                    SetJob(hPrinter, job.JobId, 0, IntPtr.Zero, JOB_CONTROL_DELETE);

                                    if (JobCaptured != null)
                                    {
                                        PrintJobCapturedEventArgs args = new PrintJobCapturedEventArgs();
                                        args.JobId = jobId;
                                        args.DocumentTitle = job.pDocument ?? ("Job #" + jobId);
                                        args.SpoolFilePath = splPath;
                                        args.SpoolBytes = bytes;
                                        XlsPrinter.Core.AppLogger.Log(string.Format("SpoolWatcher: Firing JobCaptured event for job #{0}", jobId));
                                        JobCaptured(this, args);
                                    }
                                }
                                else
                                {
                                    XlsPrinter.Core.AppLogger.Log(string.Format("SpoolWatcher: Failed to read spool file {0}", splPath));
                                }
                            }
                            else
                            {
                                XlsPrinter.Core.AppLogger.Log(string.Format("SpoolWatcher: Spool file for job #{0} not found in {1}", jobId, spoolFolder));
                            }
                        }
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(pJobBuf);
                }
            }
            finally
            {
                ClosePrinter(hPrinter);
            }
        }

        private static string FindSpoolFileForJob(string spoolFolder, int jobId)
        {
            if (!Directory.Exists(spoolFolder))
                return null;

            // Standard Windows spooler file naming: 00005.SPL or FP00005.SPL
            string fiveDigit = Path.Combine(spoolFolder, jobId.ToString("D5") + ".SPL");
            if (File.Exists(fiveDigit)) return fiveDigit;

            string fpFiveDigit = Path.Combine(spoolFolder, "FP" + jobId.ToString("D5") + ".SPL");
            if (File.Exists(fpFiveDigit)) return fpFiveDigit;

            try
            {
                string[] files = Directory.GetFiles(spoolFolder, "*.SPL");
                string suffix = jobId.ToString("D5") + ".SPL";
                for (int i = 0; i < files.Length; i++)
                {
                    if (files[i].EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                        return files[i];
                }
            }
            catch { }

            return null;
        }

        private static byte[] TryReadSharedFile(string path)
        {
            for (int attempt = 0; attempt < 5; attempt++)
            {
                try
                {
                    using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    {
                        byte[] buf = new byte[fs.Length];
                        int read = 0;
                        while (read < buf.Length)
                        {
                            int n = fs.Read(buf, read, buf.Length - read);
                            if (n <= 0) break;
                            read += n;
                        }
                        return buf;
                    }
                }
                catch
                {
                    Thread.Sleep(150);
                }
            }
            return null;
        }

        public void Dispose()
        {
            StopWatching();
        }
    }
}