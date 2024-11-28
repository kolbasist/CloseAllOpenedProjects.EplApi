using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MyEplanAPI.Service;

namespace Warden
{
    internal class DialogWindowCloser
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

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxLength);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private const uint WM_CLOSE = 0x0010;

        // Start the task to close dialog windows
        public void StartDialogClosingTask()
        {
            Logger.SendMSGToEplanLog("[DialogWindowCloser.StartDialogClosingTask] Starting dialog closing task.");

            Task.Run(() =>
            {
                try
                {
                    CloseDialogsUntilEditorIsActive();
                }
                catch (Exception ex)
                {
                    Logger.SendMSGToEplanLog($"[DialogWindowCloser.StartDialogClosingTask] Exception: {ex.Message}");
                }
            });
        }

        // Close all dialog windows until the main editor window is active
        public void CloseDialogsUntilEditorIsActive()
        {
            Logger.SendMSGToEplanLog("[DialogWindowCloser.CloseDialogsUntilEditorIsActive] Invoked");

            uint eplanProcessId = (uint)Process.GetCurrentProcess().Id;
            int maxAttempts = 10;
            int attempt = 0;
            const string editorWindowClass = "AfxMDIFrame140u";

            do
            {
                Logger.SendMSGToEplanLog($"[DialogWindowCloser.CloseDialogsUntilEditorIsActive] Attempt {attempt + 1} to close dialogs.");

                IntPtr activeWindowHandle = IntPtr.Zero;
                bool foundActiveWindow = false;
                bool editorWindowEnabled = false;

                // Enumerate windows and update state
                EnumWindows((hWnd, lParam) =>
                {
                    uint processId;
                    GetWindowThreadProcessId(hWnd, out processId);

                    if (processId == eplanProcessId && IsWindowVisible(hWnd))
                    {
                        StringBuilder className = new StringBuilder(256);
                        GetClassName(hWnd, className, className.Capacity);

                        // Check if this is the editor window
                        if (className.ToString() == editorWindowClass)
                        {
                            editorWindowEnabled = IsWindowEnabled(hWnd);
                            Logger.SendMSGToEplanLog("[DialogWindowCloser.CloseDialogsUntilEditorIsActive] Editor window found.");
                            return false;
                        }

                        // If we find an active dialog window
                        if (className.ToString() == "#32770" && IsWindowEnabled(hWnd))
                        {
                            activeWindowHandle = hWnd;
                            foundActiveWindow = true;
                            Logger.SendMSGToEplanLog($"[DialogWindowCloser.CloseDialogsUntilEditorIsActive] Active dialog window found: Handle: {hWnd}");
                            return false;
                        }
                    }
                    return true;
                }, IntPtr.Zero);

                // If editor window is enabled, break the loop
                if (editorWindowEnabled)
                {
                    Logger.SendMSGToEplanLog("[DialogWindowCloser.CloseDialogsUntilEditorIsActive] Editor window is active. Exiting loop.");
                    break;
                }

                // If no active dialog window is found
                if (!foundActiveWindow)
                {
                    Logger.SendMSGToEplanLog("[DialogWindowCloser.CloseDialogsUntilEditorIsActive] No active dialog windows found. Exiting loop.");
                    break;
                }

                // Attempt to close the active dialog window
                if (foundActiveWindow && activeWindowHandle != IntPtr.Zero)
                {
                    try
                    {
                        CloseWindow(activeWindowHandle);
                        Thread.Sleep(500);

                        if (!IsWindowVisible(activeWindowHandle))
                        {
                            Logger.SendMSGToEplanLog("[DialogWindowCloser.CloseDialogsUntilEditorIsActive] Successfully closed active dialog window.");
                            attempt = 0;
                        }
                        else
                        {
                            Logger.SendMSGToEplanLog("[DialogWindowCloser.CloseDialogsUntilEditorIsActive] Failed to close the window. It is still visible.");
                            attempt++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.SendMSGToEplanLog($"[DialogWindowCloser.CloseDialogsUntilEditorIsActive] Exception while closing window: {ex.Message}");
                        attempt++;
                    }
                }

                if (attempt >= maxAttempts)
                {
                    Logger.SendMSGToEplanLog("[DialogWindowCloser.CloseDialogsUntilEditorIsActive] Maximum number of attempts reached. Exiting loop.");
                    break;
                }

                Thread.Sleep(500);
            } while (true);
        }

        // Close a specific window by handle
        private void CloseWindow(IntPtr hWnd)
        {
            try
            {
                if (hWnd != IntPtr.Zero)
                {
                    SetForegroundWindow(hWnd);
                    Thread.Sleep(200);
                    PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                    Logger.SendMSGToEplanLog($"[DialogWindowCloser.CloseWindow] Sent WM_CLOSE message to window: {GetWindowTitle(hWnd)}");
                }
            }
            catch (Exception ex)
            {
                Logger.SendMSGToEplanLog($"[DialogWindowCloser.CloseWindow] Error closing window: {ex.Message}");
            }
        }

        // Get the window title by handle
        private string GetWindowTitle(IntPtr hWnd)
        {
            StringBuilder windowText = new StringBuilder(256);
            GetWindowText(hWnd, windowText, windowText.Capacity);
            return windowText.ToString();
        }
    }
}
