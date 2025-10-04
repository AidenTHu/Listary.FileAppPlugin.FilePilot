using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Shapes;

namespace Listary.FileAppPlugin.FilePilot {
    public class FilePilotTab : IFileTab, IGetFolder, IOpenFolder {
        public FilePilotTab(string path) { }
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

        private static readonly HashSet<string> pathCache = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private int GetFileAndFolderCount(string folderPath) {
            try {
                if (!Directory.Exists(folderPath)) {
                    return 0;
                }

                int fileCount = Directory.GetFiles(folderPath, "*", SearchOption.TopDirectoryOnly).Length;
                int folderCount = Directory.GetDirectories(folderPath, "*", SearchOption.TopDirectoryOnly).Length;
                return fileCount + folderCount;
            } catch {
                return 0;
            }
        }

        public async Task<string> GetCurrentFolder() {
            // FilePilot does not expose the current folder via Windows Messages or Automation APIs.
            // Since there is no reliable way to get the current folder from File Pilot using SendKeys or MouseEvents, return null.
            return null;
        }

        public async Task<bool> OpenFolder(string path) {
            string originalClipboard = null;
            bool clipboardHasText = false;
            try {
                // Save the current clipboard content
                if (System.Windows.Clipboard.ContainsText()) {
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

                await Task.Delay(20);

                // Pastes the path into the input box
                Keyboard.Press(VirtualKeyShort.CONTROL);
                Keyboard.Type(VirtualKeyShort.KEY_V);
                Keyboard.Release(VirtualKeyShort.CONTROL);

                await Task.Delay(calculateDelay(path));

                Keyboard.Type(VirtualKeyShort.RETURN);

                // Restores the original clipboard content
                if (clipboardHasText) {
                    System.Windows.Clipboard.SetText(originalClipboard);
                }

                return true;
            }
            catch (Exception) {
                // Attempts to restore clipboard even on error
                if (clipboardHasText) {
                    try {
                        System.Windows.Clipboard.SetText(originalClipboard);
                    }
                    catch { }
                }
                return false;
            }
        }

        private int calculateDelay(string path) {
            int baseDelay = 40;
            if (!pathCache.Contains(path)) {
                pathCache.Add(path);

                // First time opening a folder in File Pilot is slower, add extra delay
                baseDelay += 200;
            }

            // Calculate file and folder count and set delay using scaling factor
            int fileAndFolderCount = GetFileAndFolderCount(path);
            double scaleFactor;

            if (fileAndFolderCount > 1000) {
                scaleFactor = .1;
            }
            else if (fileAndFolderCount > 500) {
                scaleFactor = .58;
            }
            else if (fileAndFolderCount > 50) {
                scaleFactor = 1.3;
            }
            else {
                scaleFactor = .018;
            }

            return baseDelay + (int)(fileAndFolderCount * scaleFactor);
        }

        public void FocusFilePilotWindow(IntPtr hwnd) {
            ShowWindow(hwnd, IsZoomed(hwnd) ? SW_MAXIMIZE : SW_SHOW);
            SetForegroundWindow(hwnd);
        }
    }
}