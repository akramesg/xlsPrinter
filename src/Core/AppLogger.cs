using System;
using System.IO;
using System.Threading;

namespace XlsPrinter.Core
{
    public static class AppLogger
    {
        private static readonly object _lock = new object();
        private static string _logPath;

        static AppLogger()
        {
            try
            {
                string loc = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string dir = !string.IsNullOrEmpty(loc) ? Path.GetDirectoryName(loc) : AppDomain.CurrentDomain.BaseDirectory;
                if (!Directory.Exists(dir) || dir.IndexOf("System32", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    dir = Path.GetTempPath();
                }
                _logPath = Path.Combine(dir, "xls_debug.log");
            }
            catch
            {
                _logPath = Path.Combine(Path.GetTempPath(), "xls_debug.log");
            }
        }

        public static string LogFilePath
        {
            get { return _logPath; }
        }

        public static void Log(string message)
        {
            try
            {
                lock (_lock)
                {
                    string line = string.Format("[{0:yyyy-MM-dd HH:mm:ss.fff}] [T{1:D2}] {2}\r\n",
                        DateTime.Now, Thread.CurrentThread.ManagedThreadId, message);
                    File.AppendAllText(_logPath, line);
                }
            }
            catch { }
        }
    }
}

