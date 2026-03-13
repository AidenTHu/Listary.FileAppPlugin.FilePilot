using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
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

        public async Task<string> GetCurrentFolder() {
            // FilePilot does not expose the current folder via Windows Messages or Automation APIs.
            // Since there is no reliable way to get the current folder from File Pilot using SendKeys or MouseEvents, return null.
            return null;
        }

        public async Task<bool> OpenFolder(string path) {
            string originalClipboard = null; 
           StringCollection originalFileList = null;
            try {
                // Save the current clipboard content
                // Prefer preserving file-drop lists (copied files)
                if (System.Windows.Clipboard.ContainsFileDropList()) {
                    try {
                        originalFileList = System.Windows.Clipboard.GetFileDropList();
                    }
                    catch {}
                }
                if (System.Windows.Clipboard.ContainsText()) {
                    try {
                        originalClipboard = System.Windows.Clipboard.GetText();
                    }
                    catch {}
                }

                // Sets the clipboard to the folder path
                System.Windows.Clipboard.SetText(path);

                this.FocusFilePilotWindow(FindWindow("File Pilot", null));

                // Opens the File Pilot folder path input box
                Keyboard.Press(VirtualKeyShort.CONTROL);
                Keyboard.Type(VirtualKeyShort.KEY_L);
                Keyboard.Release(VirtualKeyShort.CONTROL);

                await Task.Delay(20);

                // Pastes the path into the input box
                Keyboard.Press(VirtualKeyShort.CONTROL);
                Keyboard.Type(VirtualKeyShort.KEY_V);
                Keyboard.Release(VirtualKeyShort.CONTROL);

                await Task.Delay(this.calculateDelay(path));

                Keyboard.Type(VirtualKeyShort.RETURN);
                return true;
            }
            catch (Exception) {
                return false;
            }
            finally {
                try {
                    if (originalClipboard != null) {
                        System.Windows.Clipboard.SetText(originalClipboard);
                    }
                    if (originalFileList != null) {
                        System.Windows.Clipboard.SetFileDropList(originalFileList);
                    }
                }
                catch {}
            }
        }

        private int calculateDelay(string path) {
            if (!pathCache.Contains(path)) {
                pathCache.Add(path);
                // First time opening a folder in File Pilot is slower, add extra delay
                return 200;
            } else {
                return 40;
            }
        }

        public void FocusFilePilotWindow(IntPtr hwnd) {
            ShowWindow(hwnd, IsZoomed(hwnd) ? SW_MAXIMIZE : SW_SHOW);
            SetForegroundWindow(hwnd);
        }
    }
}