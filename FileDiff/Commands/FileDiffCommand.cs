using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace FileDiff
{
    internal sealed class FileDiffCommand
    {
        public const int CommandId = 0x0100;

        public static readonly Guid CommandSet = new Guid("dc74d78f-f5fa-4944-a3bd-b621c50ba12b");

        private readonly AsyncPackage package;

        private FileDiffCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static FileDiffCommand Instance
        {
            get;
            private set;
        }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new FileDiffCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            var dte1 = ServiceProvider.GetServiceAsync(typeof(DTE));
            DTE2 dte = dte1.Result as DTE2;

            string file1, file2;
            if (CanFilesBeCompared(dte, out file1, out file2))
            {
                dte.ExecuteCommand("Tools.DiffFiles", $"\"{file1}\" \"{file2}\"");
            }

        }
        private static bool CanFilesBeCompared(DTE2 dte, out string file1, out string file2)
        {
            var items = GetSelectedFiles(dte);
            file1 = "";
            file2 = "";
            bool val = false;
            switch(items.Count())
            {
                case 1:
                    file1 = items.ElementAtOrDefault(0);
                    var dialog = new OpenFileDialog();
                    dialog.InitialDirectory = Path.GetDirectoryName(file1);
                    dialog.ShowDialog();
                    file2 = dialog.FileName;
                    val = !string.IsNullOrEmpty(file1) && !string.IsNullOrEmpty(file2);
                    break;
                case 2:
                    file1 = items.ElementAtOrDefault(0);
                    file2 = items.ElementAtOrDefault(1);
                    val = !string.IsNullOrEmpty(file1) && !string.IsNullOrEmpty(file2);
                    break;
                default:
                    break;
            }
            
            
            return val;

        }
        public static IEnumerable<string> GetSelectedFiles(DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var items = (Array)dte.ToolWindows.SolutionExplorer.SelectedItems;
            return from item in items.Cast<UIHierarchyItem>()
                   let pi = item.Object as ProjectItem
                   select pi.FileNames[1];
        }
    }
}
