using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using MyEplanAPI.Service;

namespace Warden
{
    internal class WindowLogger
    {
        // Import necessary functions from user32.dll
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxLength);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder text, int maxLength);

        // Log all windows related to the current process
        public void LogAllWindows()
        {
            Logger.SendMSGToEplanLog("[WindowLogger.LogAllWindows] Invoked");

            uint eplanProcessId = (uint)Process.GetCurrentProcess().Id;

            EnumWindows((hWnd, lParam) =>
            {
                uint processId;
                GetWindowThreadProcessId(hWnd, out processId);

                if (processId == eplanProcessId)
                {
                    // Get window properties
                    bool isVisible = IsWindowVisible(hWnd);
                    bool isEnabled = IsWindowEnabled(hWnd);

                    StringBuilder windowText = new StringBuilder(256);
                    GetWindowText(hWnd, windowText, windowText.Capacity);

                    StringBuilder className = new StringBuilder(256);
                    GetClassName(hWnd, className, className.Capacity);

                    // Formulate window information string
                    string windowInfo = ($"Handle: {hWnd}, Title: \"{windowText}\", Class: \"{className}\", Visible: {isVisible}, Enabled: {isEnabled}");

                    // Log window information
                    Logger.SendMSGToEplanLog($"[WindowLogger.LogAllWindows] {windowInfo}");
                }
                return true;
            }, IntPtr.Zero);
        }
    }
}
