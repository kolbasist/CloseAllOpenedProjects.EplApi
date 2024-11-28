using Eplan.EplApi.Base;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace MyEplanAPI.Service
{
    internal static class Logger
    {
        private static readonly string logFilePath = "C:\\Users\\Public\\Eplan\\Common\\EplanAPILog.txt";

        // Send message to EPLAN log and write to a file
        public static void SendMSGToEplanLog(string msg)
        {
            // Write to EPLAN log
            BaseException exc = new BaseException(msg, MessageLevel.Message);
            exc.FixMessage();

            // Write to log file
            LogToFile(msg);
        }

        // Write message to a text file
        private static void LogToFile(string msg)
        {
            try
            {
                // Get directory from the path
                string directoryPath = Path.GetDirectoryName(logFilePath);

                if (!string.IsNullOrEmpty(directoryPath) && !Directory.Exists(directoryPath))
                {
                    // Create directories if they do not exist
                    Directory.CreateDirectory(directoryPath);
                }

                // Check if the file exists
                if (!File.Exists(logFilePath))
                {
                    // Create the file if it does not exist
                    using (FileStream fs = File.Create(logFilePath))
                    {
                        // Write a log header if needed
                        string header = "Timestamp                  [Process]        [Method]       Message";
                        byte[] info = new UTF8Encoding(true).GetBytes(header + Environment.NewLine);
                        fs.Write(info, 0, info.Length);
                    }
                }

                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string processName = Process.GetCurrentProcess().ProcessName;
                string methodName = GetCallingMethodName();

                string logMessage = $"{timestamp} [{processName}] [{methodName}] {msg}";

                // Append message to log file
                File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // Handle exceptions during log file writing
                Console.WriteLine($"Error writing log: {ex.Message}");
            }
        }

        // Get the name of the calling method for logging purposes
        private static string GetCallingMethodName()
        {
            StackTrace stackTrace = new StackTrace();

            // Index 2, because 0 - current method, 1 - LogToFile method, 2 - calling method
            if (stackTrace.FrameCount > 2)
            {
                var method = stackTrace.GetFrame(2).GetMethod();
                return $"{method.DeclaringType.FullName}.{method.Name}";
            }

            return "Unknown method";
        }
    }
}
