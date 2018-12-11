using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Runtime.InteropServices;
using Task = System.Threading.Tasks.Task;

namespace Bas.OpenPowerShellHere
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class OpenPowerShellHereCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("0dc193e5-18d2-4acc-8438-d01ddc43b95a");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenPowerShellHereCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private OpenPowerShellHereCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static OpenPowerShellHereCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in OpenPowerShellHereCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new OpenPowerShellHereCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            var folderPath = GetSelectedItemFolderPath();

            var powershellPaths = new[]
            {
                Path.Combine(Environment.GetEnvironmentVariable("ProgramW6432"), "PowerShell\\6\\pwsh.exe"), // Try powershell 6 first
                "pwsh.exe", // hopefully if it's in another folder, that folder is in the PATH.
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell\\v1.0\\powershell.exe"), // otherwise try Powershell 1
                "powershell.exe"
            };

            foreach (var powershellPath in powershellPaths)
            {
                var process = new System.Diagnostics.Process();
                process.StartInfo.FileName = powershellPath;
                //process.StartInfo.Arguments = options;
                process.StartInfo.WorkingDirectory = folderPath;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = false;

                try
                {
                    process.Start();
                    break; // If powershell was launched successfully, we can exit the loop and the function.
                }
                catch
                {
                    continue; // if not, try again with the next path.
                }
            }
        }

        private static string GetSelectedItemFolderPath()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var monitorSelection = Package.GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            var solution = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution;

            IVsMultiItemSelect multiItemSelect = null;
            IntPtr hierarchyPtr = IntPtr.Zero;
            IntPtr selectionContainerPtr = IntPtr.Zero;

            try
            {
                var hr = monitorSelection.GetCurrentSelection(out hierarchyPtr, out uint itemId, out multiItemSelect, out selectionContainerPtr);

                if (multiItemSelect == null)
                {
                    if (hierarchyPtr.ToInt32() == 0)
                    {
                        solution.GetSolutionInfo(out string solutionPath, out _, out _);
                        return solutionPath;
                    }
                    else
                    {
                        var hierarchy = Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy;

                        ((IVsProject)hierarchy).GetMkDocument(itemId, out string itemFullPath);

                        return Path.GetDirectoryName(itemFullPath);
                    }
                }
            }
            finally
            {
                if (selectionContainerPtr != IntPtr.Zero)
                {
                    Marshal.Release(selectionContainerPtr);
                }

                if (hierarchyPtr != IntPtr.Zero)
                {
                    Marshal.Release(hierarchyPtr);
                }
            }

            return null;
        }
    }
}
