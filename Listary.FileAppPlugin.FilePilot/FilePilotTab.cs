using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

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

        // We'll use SendKeys to send keyboard shortcuts (no external dependencies)

        private static readonly HashSet<string> pathCache = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public async Task<string> GetCurrentFolder() {
            // FilePilot does not expose the current folder via Windows Messages or Automation APIs.
            // Since there is no reliable way to get the current folder from File Pilot using SendKeys or MouseEvents, return null.
            return null;
        }

        private static void RunOnSta(Action action) {
            var t = new Thread(() => {
                try { action(); }
                catch { }
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
        }

        public async Task<bool> OpenFolder(string path) {
            string originalClipboard = null; 
            StringCollection originalFileList = null;
            try {
                // Save the current clipboard content on an STA thread
                    try {
                    originalFileList = GetClipboardFileDropList();
                    }
                catch { originalFileList = null; }

                    try {
                    originalClipboard = GetClipboardText();
                    }
                catch { originalClipboard = null; }

                // Sets the clipboard to the folder path
                SetClipboardText(path);

                this.FocusFilePilotWindow(FindWindow("File Pilot", null));

                // Perform the keystrokes on an STA thread using SendKeys (requires STA)
                RunOnSta(() => {
                    // Opens the File Pilot folder path input box (Ctrl+L)
                    System.Windows.Forms.SendKeys.SendWait("^l");
                    Thread.Sleep(20);

                    // Pastes the path into the input box (Ctrl+V)
                    System.Windows.Forms.SendKeys.SendWait("^v");
                    Thread.Sleep(this.calculateDelay(path));

                    // Press Enter
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                });
                return true;
            }
            catch (Exception) {
                return false;
            }
            finally {
                try {
                    // Prefer restoring file-drop list if it existed, otherwise restore text
                    if (originalFileList != null) {
                        SetClipboardFileDropList(originalFileList);
                    }
                    else if (originalClipboard != null) {
                        SetClipboardText(originalClipboard);
                    }
                }
                catch {}
            }
        }

        // Clipboard helpers that run on STA threads to avoid COM/clipboard apartment issues
        private static string GetClipboardText() {
            string result = null;
            var t = new Thread(() => {
                try {
                    if (System.Windows.Forms.Clipboard.ContainsText()) result = System.Windows.Forms.Clipboard.GetText();
                }
                catch { result = null; }
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
            return result;
        }

        private static StringCollection GetClipboardFileDropList() {
            StringCollection result = null;
            var t = new Thread(() => {
                try {
                    if (System.Windows.Forms.Clipboard.ContainsFileDropList()) result = System.Windows.Forms.Clipboard.GetFileDropList();
                }
                catch { result = null; }
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
            return result;
        }

        private static void SetClipboardText(string text) {
            var t = new Thread(() => {
                try { System.Windows.Forms.Clipboard.SetText(text); }
                catch { }
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
        }

        private static void SetClipboardFileDropList(StringCollection list) {
            var t = new Thread(() => {
                try { System.Windows.Forms.Clipboard.SetFileDropList(list); }
                catch { }
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
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