using System;
using System.Threading.Tasks;

namespace Listary.FileAppPlugin.FilePilot
{
    public class FilePilotWindow : IFileWindow
    {
        private IFileAppPluginHost _host;

        public IntPtr Handle { get; }

        public FilePilotWindow(IFileAppPluginHost host) {}

        public async Task<IFileTab> GetCurrentTab()
        {
            return new FilePilotTab();
        }
    }
}
