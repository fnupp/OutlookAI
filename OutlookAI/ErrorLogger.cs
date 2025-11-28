using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace OutlookAI
{
    /// <summary>
    /// Simple error logging utility for OutlookAI
    /// </summary>
    public static class ErrorLogger
    {
        private static readonly object _logLock = new object();
        private static string _logFilePath;
        private static readonly AsyncLocal<CorrelationContext> _correlationContext = new AsyncLocal<CorrelationContext>();

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
                        string correlationId = _correlationContext.Value?.CorrelationId ?? "";
                        writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{correlationId}] ERROR: {message}");

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
                        string correlationId = _correlationContext.Value?.CorrelationId ?? "";
                        writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{correlationId}] INFO: {message}");
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
                        string correlationId = _correlationContext.Value?.CorrelationId ?? "";
                        writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{correlationId}] WARNING: {message}");
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

        /// <summary>
        /// Starts a new correlation scope for grouping related log entries
        /// </summary>
        public static IDisposable BeginCorrelation(string operationName = null)
        {
            var correlationId = GenerateCorrelationId();
            var context = new CorrelationContext(correlationId, operationName);
            _correlationContext.Value = context;

            LogInfo($"=== BEGIN: {operationName ?? "Operation"} ===");
            return new CorrelationScope(context);
        }

        /// <summary>
        /// Captures the current correlation context for propagation across Task.Run() boundaries
        /// </summary>
        public static CorrelationContext CaptureContext()
        {
            return _correlationContext.Value;
        }

        /// <summary>
        /// Restores a captured correlation context inside Task.Run()
        /// </summary>
        public static void RestoreContext(CorrelationContext context)
        {
            _correlationContext.Value = context;
        }

        /// <summary>
        /// Generates a random 8-character correlation ID
        /// </summary>
        private static string GenerateCorrelationId()
        {
            const string chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
            var random = new Random(Guid.NewGuid().GetHashCode());
            return new string(Enumerable.Range(0, 8)
                .Select(_ => chars[random.Next(chars.Length)])
                .ToArray());
        }

        /// <summary>
        /// IDisposable scope for automatic BEGIN/END correlation logging
        /// </summary>
        private class CorrelationScope : IDisposable
        {
            private readonly CorrelationContext _context;

            public CorrelationScope(CorrelationContext context)
            {
                _context = context;
            }

            public void Dispose()
            {
                if (_context != null)
                {
                    var elapsed = DateTime.Now - _context.StartTime;
                    ErrorLogger.LogInfo($"=== END: {_context.OperationName ?? "Operation"} (Duration: {elapsed.TotalMilliseconds:F0}ms) ===");
                    _correlationContext.Value = null;
                }
            }
        }
    }
}
