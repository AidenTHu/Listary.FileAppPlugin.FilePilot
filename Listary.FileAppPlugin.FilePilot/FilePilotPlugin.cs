using System;
using System.Threading.Tasks;

namespace Listary.FileAppPlugin.FilePilot
{
    public class FilePilotPlugin : IFileAppPlugin
    {
        private IFileAppPluginHost _host;

        public bool IsOpenedFolderProvider => true;
        
        public bool IsQuickSwitchTarget => true;
        
        public bool IsSharedAcrossApplications => false;

        public SearchBarType SearchBarType => SearchBarType.Floating;
        
        public async Task<bool> Initialize(IFileAppPluginHost host)
        {
            return true;
        }

        public IFileWindow BindFileWindow(IntPtr hWnd)
        {
            return Win32Utils.GetClassName(hWnd) == "File Pilot" ? new FilePilotWindow(_host) : null;
        }
    }
}
