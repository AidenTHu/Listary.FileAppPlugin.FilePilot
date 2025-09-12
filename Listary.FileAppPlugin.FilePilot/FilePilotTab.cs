using FlaUI.Core.WindowsAPI;
using FlaUI.Core.Input;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Listary.FileAppPlugin.FilePilot
{
    public class FilePilotTab : IFileTab, IGetFolder, IOpenFolder
    {
        // Win32 API declarations
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsZoomed(IntPtr hWnd);

        private const int SW_SHOW = 5;
        private const int SW_MAXIMIZE = 3;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public async Task<string> GetCurrentFolder()
        {
            // FilePilot does not expose the current folder via Windows Messages or Automation APIs.
            // Since there is no reliable way to get the current folder from File Pilot using SendKeys or MouseEvents, return null.
            return null;
        }

        public async Task<bool> OpenFolder(string path)
        {
            string originalClipboard = null;
            bool clipboardHasText = false;
            try
            {
                // Save the current clipboard content
                if (System.Windows.Clipboard.ContainsText())
                {
                    originalClipboard = System.Windows.Clipboard.GetText();
                    clipboardHasText = true;
                }

                // Sets the clipboard to the folder path
                System.Windows.Clipboard.SetText(path);

                FocusFilePilotWindow(FindWindow("File Pilot", null));

                // Opens the File Pilot folder path input box
                Keyboard.Press(VirtualKeyShort.CONTROL);
                Keyboard.Type(VirtualKeyShort.KEY_L);
                Keyboard.Release(VirtualKeyShort.CONTROL);
                await Task.Delay(5);

                // Pastes the path into the input box
                Keyboard.Press(VirtualKeyShort.CONTROL);
                Keyboard.Type(VirtualKeyShort.KEY_V);
                Keyboard.Release(VirtualKeyShort.CONTROL);

                // Presses enter to input the path. Wait till the file path is entered before pressing enter.
                await Task.Delay(80);
                Keyboard.Type(VirtualKeyShort.RETURN);

                // Restores the original clipboard content
                if (clipboardHasText)
                {
                    System.Windows.Clipboard.SetText(originalClipboard);
                }

                return true;
            }
            catch (Exception)
            {
                // Attempts to restore clipboard even on error
                if (clipboardHasText)
                {
                    try
                    {
                        System.Windows.Clipboard.SetText(originalClipboard);
                    }
                    catch { }
                }
                return false;
            }
        }

        public void FocusFilePilotWindow(IntPtr hwnd)
        {
            ShowWindow(hwnd, IsZoomed(hwnd) ? SW_MAXIMIZE : SW_SHOW);
            SetForegroundWindow(hwnd);
        }
    }
}