using System;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell.Settings;
using Microsoft.Xrm.Sdk.Query;
using Task = System.Threading.Tasks.Task;
using EnvDTE80;
using EnvDTE;

namespace WebResourcePlugin
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ConnectToDynamics
    {
        /// <summary>
        /// Connect command ID.
        /// </summary>
        public const int ConnectCommandId = 0x0100;

        /// <summary>
        /// Publish command ID.
        /// </summary>
        public const int PublishCommandId = 0x0102;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("5bc53386-793b-4d0a-a6ca-cb36519f1873");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConnectToDynamics"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ConnectToDynamics(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
            
            SettingsManager = new ShellSettingsManager(package);

            var menuCommandId = new CommandID(CommandSet, ConnectCommandId);
            var menuItem = new MenuCommand(this.ExecuteConnect, menuCommandId);
            commandService.AddCommand(menuItem);

            menuCommandId = new CommandID(CommandSet, PublishCommandId);
            menuItem = new MenuCommand(this.ExecutePublish, menuCommandId);
            commandService.AddCommand(menuItem);
        }

        public bool IsConnected => Service != null && Service.IsReady;

        public CrmServiceClient Service { get; private set; } = null;
        public ShellSettingsManager SettingsManager { get; } = null;

        public bool MakeConnection(string url, string domain, string user, string password, bool save)
        {
            var connectionString =
                $"RequireNewInstance=True; AuthType=AD; Url={url}; Domain={domain}; Username={user}; Password={password};";

            Service = new CrmServiceClient(connectionString);

            if (IsConnected && save)
            {
                var settings = ConnectionSettings;
                settings.Domain = domain;
                settings.User = user;
                settings.Url = url;
                ConnectionSettings = settings;
            }

            return IsConnected;
        }

        public string[] GetSolutions()
        {
            if (!IsConnected)
            {
                return new string[0];
            }

            var queryGetSolutions = new QueryExpression
            {
                EntityName = "solution",
                ColumnSet = new ColumnSet("friendlyname"),
                Criteria = new FilterExpression(),
            };

            var solutions = Service.RetrieveMultiple(queryGetSolutions);

            // Filter to only show Edison solutions
            var entities = solutions.Entities.Where(e => e.GetAttributeValue<string>("friendlyname").Contains("Edison"));
            return entities.Select(e => e.GetAttributeValue<string>("friendlyname")).ToArray();
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ConnectToDynamics Instance
        {
            get;
            private set;
        }

        public SettingsModel ConnectionSettings
        {
            get
            {
                var readSettings = SettingsManager.GetReadOnlySettingsStore(SettingsScope.UserSettings);
                if (!readSettings.CollectionExists("DynamicsConnection"))
                {
                    return new SettingsModel();
                }

                var settings = readSettings.GetPropertyNames("DynamicsConnection").ToArray();
                var model = new SettingsModel();

                if (settings.Contains("Domain"))
                {
                    model.Domain = readSettings.GetString("DynamicsConnection", "Domain");
                }

                if (settings.Contains("User"))
                {
                    model.User = readSettings.GetString("DynamicsConnection", "User");
                }

                if (settings.Contains("Url"))
                {
                    model.Url = readSettings.GetString("DynamicsConnection", "Url");
                }

                if (settings.Contains("Solution"))
                {
                    model.Solution = readSettings.GetString("DynamicsConnection", "Solution");
                }

                return model;
            }
            set
            {
                var writableSettings = SettingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);
                if (!writableSettings.CollectionExists("DynamicsConnection"))
                {
                    writableSettings.CreateCollection("DynamicsConnection");
                }
                
                writableSettings.SetString("DynamicsConnection", "Domain", value.Domain ?? string.Empty);
                writableSettings.SetString("DynamicsConnection", "User", value.User ?? string.Empty);
                writableSettings.SetString("DynamicsConnection", "Url", value.Url ?? string.Empty);
                writableSettings.SetString("DynamicsConnection", "Solution", value.Solution ?? string.Empty);
            }
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        // ReSharper disable once UnusedMember.Local
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => this.package;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ConnectToDynamics's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ConnectToDynamics(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void ExecuteConnect(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var connectWindow = new ConnectWindow();
            connectWindow.ShowDialog();
        }

        private void ExecutePublish(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsConnected)
            {
                ExecuteConnect(sender, e);
            }

            if (IsConnected)
            {
                var projectItem = GetSelectedProjectItem();
                if (projectItem == null)
                {
                    return;
                }

                // Save the item
                projectItem.Document.Save();

                // Read the content
                var content = File.ReadAllText(projectItem.Document.FullName);

                VsShellUtilities.ShowMessageBox(
                    this.package,
                    content,
                    "PublishToDynamics",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private ProjectItem GetSelectedProjectItem()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var applicationObjectTask = this.ServiceProvider.GetServiceAsync(typeof(SDTE));
            var applicationObject = applicationObjectTask.Result as DTE2;
            if (applicationObject == null)
            {
                return null;
            }

            var solutionExplorer = applicationObject.ToolWindows.SolutionExplorer;
            var items = solutionExplorer.SelectedItems as EnvDTE.UIHierarchyItem[];
            if (items == null || items.Length == 0)
            {
                return null;
            }

            var file = items.FirstOrDefault();
            return file?.Object as ProjectItem;
        }
    }
}
