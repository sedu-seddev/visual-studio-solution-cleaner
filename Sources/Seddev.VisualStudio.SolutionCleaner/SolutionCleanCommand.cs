namespace Seddev.VisualStudio.SolutionCleaner
{
    using Microsoft.VisualStudio;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using Seddev.VisualStudio.SolutionCleaner.Models;
    using System;
    using System.ComponentModel.Design;
    using System.IO;
    using System.Windows.Forms;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SolutionCleanCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("7a81dddc-354a-43e7-b698-dfc0b41737b8");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SolutionCleanCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private SolutionCleanCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static SolutionCleanCommand Instance
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
            // Switch to the main thread - the call to AddCommand in SolutionCleanCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SolutionCleanCommand(package, commandService);
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
            ThreadHelper.ThrowIfNotOnUIThread();
            OutputWindowTextWriter outputWindowTextWriter = null;

            try
            {
                outputWindowTextWriter = GetOutputWindowTextWriter();
                FullCleanSolution(outputWindowTextWriter);
            }
            catch (Exception ex)
            {
                outputWindowTextWriter?.WriteLine("FATAL> Error while cleaning the solution!");
                outputWindowTextWriter?.WriteLine($"FATAL> {ex.Message}");
                outputWindowTextWriter?.WriteLine($"FATAL> {ex.StackTrace}");

                outputWindowTextWriter?.Close();
                outputWindowTextWriter?.Dispose();
            }            
        }

        /// <summary>
        /// Creates a build output window pane text writer. Caller has to dispose the writer.
        /// </summary>
        /// <returns>OutputWindowTextWriter of type VSConstants.GUID_BuildOutputWindowPane.</returns>
        private static OutputWindowTextWriter GetOutputWindowTextWriter()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
            var windowPaneGuid = VSConstants.GUID_BuildOutputWindowPane;
            outWindow.GetPane(ref windowPaneGuid, out var buildOutputPane);
            
            buildOutputPane.Activate();

            var outputWindowTextWriter = new OutputWindowTextWriter(buildOutputPane);
            return outputWindowTextWriter;
        }

        private static void FullCleanSolution(OutputWindowTextWriter outputWindowTextWriter)
        {
            outputWindowTextWriter.WriteLine("Full clean started...");

            var solutionDirectory = Directory.GetCurrentDirectory();
            var cleaningProcess = new CleaningProcess(solutionDirectory);

            foreach (var project in cleaningProcess.SolutionProjects)
            {
                try
                {                                        
                    outputWindowTextWriter.WriteLine($"{project.Id}> ------ Full clean started: Project: {project.Name} ------");

                    CleanDirectory(outputWindowTextWriter, project, "bin");
                    CleanDirectory(outputWindowTextWriter, project, "obj");

                    cleaningProcess.SuccessCount++;
                }
                catch (Exception ex)
                {
                    cleaningProcess.FailedCount++;

                    outputWindowTextWriter.WriteLine($"ERROR: {project.Id}> Error while cleaning {project.Name}");
                    outputWindowTextWriter.WriteLine($"ERROR: {project.Id}> {ex.Message}");
                    outputWindowTextWriter.WriteLine($"ERROR: {project.Id}> {ex.StackTrace}");
                }

                outputWindowTextWriter.WriteLine($"========== Full clean: {cleaningProcess.SuccessCount} succeeded, {cleaningProcess.FailedCount} failed ==========");
            }
        }

        private static void CleanDirectory(OutputWindowTextWriter outputWindowTextWriter, Project project, string folderToClean)
        {
            var pathToClean = Path.Combine(project.Path, folderToClean);

            if (Directory.Exists(pathToClean))
            {
                Directory.Delete(pathToClean, true);
                outputWindowTextWriter.WriteLine($"{project.Id}> {folderToClean} folder cleaned for project {project.Name}");
            }
            else
            {
                outputWindowTextWriter.WriteLine($"{project.Id}> No {folderToClean} folder exists for project {project.Name}");
            }
        }
    }
}
