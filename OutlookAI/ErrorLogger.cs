using System;
using System.IO;

namespace OutlookAI
{
    /// <summary>
    /// Simple error logging utility for OutlookAI
    /// </summary>
    public static class ErrorLogger
    {
        private static readonly object _logLock = new object();
        private static string _logFilePath;

        static ErrorLogger()
        {
            string logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "OutlookAI",
                "Logs");

            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }

            _logFilePath = Path.Combine(logDirectory, $"OutlookAI_{DateTime.Now:yyyyMMdd}.log");
        }

        /// <summary>
        /// Logs an error message to the log file
        /// </summary>
        public static void LogError(string message, Exception ex = null)
        {
            if (!ThisAddIn.userdata?.LogErrors ?? true)
                return;

            try
            {
                lock (_logLock)
                {
                    using (StreamWriter writer = new StreamWriter(_logFilePath, append: true))
                    {
                        writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR: {message}");

                        if (ex != null)
                        {
                            writer.WriteLine($"  Exception: {ex.GetType().Name}");
                            writer.WriteLine($"  Message: {ex.Message}");
                            writer.WriteLine($"  StackTrace: {ex.StackTrace}");

                            if (ex.InnerException != null)
                            {
                                writer.WriteLine($"  Inner Exception: {ex.InnerException.GetType().Name}");
                                writer.WriteLine($"  Inner Message: {ex.InnerException.Message}");
                            }
                        }

                        writer.WriteLine();
                    }
                }
            }
            catch
            {
                // Fail silently - don't let logging errors crash the add-in
                System.Diagnostics.Debug.WriteLine($"Failed to write to log file: {message}");
            }
        }

        /// <summary>
        /// Logs an informational message to the log file
        /// </summary>
        public static void LogInfo(string message)
        {
            if (!ThisAddIn.userdata?.LogErrors ?? true)
                return;

            try
            {
                lock (_logLock)
                {
                    using (StreamWriter writer = new StreamWriter(_logFilePath, append: true))
                    {
                        writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] INFO: {message}");
                    }
                }
            }
            catch
            {
                // Fail silently
                System.Diagnostics.Debug.WriteLine($"Failed to write to log file: {message}");
            }
        }

        /// <summary>
        /// Logs a warning message to the log file
        /// </summary>
        public static void LogWarning(string message)
        {
            if (!ThisAddIn.userdata?.LogErrors ?? true)
                return;

            try
            {
                lock (_logLock)
                {
                    using (StreamWriter writer = new StreamWriter(_logFilePath, append: true))
                    {
                        writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] WARNING: {message}");
                    }
                }
            }
            catch
            {
                // Fail silently
                System.Diagnostics.Debug.WriteLine($"Failed to write to log file: {message}");
            }
        }

        /// <summary>
        /// Gets the path to the current log file
        /// </summary>
        public static string GetLogFilePath()
        {
            return _logFilePath;
        }

        /// <summary>
        /// Deletes log files older than the specified number of days
        /// </summary>
        public static void CleanupOldLogs(int daysToKeep = 30)
        {
            try
            {
                string logDirectory = Path.GetDirectoryName(_logFilePath);
                var logFiles = Directory.GetFiles(logDirectory, "OutlookAI_*.log");

                foreach (var logFile in logFiles)
                {
                    FileInfo fileInfo = new FileInfo(logFile);
                    if (fileInfo.LastWriteTime < DateTime.Now.AddDays(-daysToKeep))
                    {
                        File.Delete(logFile);
                    }
                }
            }
            catch
            {
                // Fail silently
            }
        }
    }
}
