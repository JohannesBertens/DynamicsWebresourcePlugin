using EnvDTE;
using EnvDTE80;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell.Settings;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.Threading;
using Task = System.Threading.Tasks.Task;

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
        /// Update command ID.
        /// </summary>
        public const int UpdateCommandId = 0x0104;

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

            menuCommandId = new CommandID(CommandSet, UpdateCommandId);
            menuItem = new MenuCommand(this.ExecuteUpdate, menuCommandId);
            commandService.AddCommand(menuItem);
        }

        public bool IsConnected => Service != null && Service.IsReady;

        public CrmServiceClient Service { get; private set; } = null;
        public ShellSettingsManager SettingsManager { get; } = null;

        public string SelectedSolution { get; set; } = null;

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

        private Guid GetPublisher(string publisherName)
        {
            var queryPublisher = new QueryExpression("publisher");
            queryPublisher.Criteria.AddCondition("friendlyname", ConditionOperator.Equal, publisherName);
            var result = Service.RetrieveMultiple(queryPublisher);

            if (result.Entities.Count != 1)
            {
                return default;
            }

            var publisher = result.Entities.FirstOrDefault();
            if (publisher == null)
            {
                return default;
            }

            return publisher.Id;
        }


        public string[] GetSolutions()
        {
            if (!IsConnected)
            {
                return new string[0];
            }

            var microsoftPublisherId = GetPublisher("MicrosoftCorporation");

            var queryGetSolutions = new QueryExpression
            {
                EntityName = "solution",
                ColumnSet = new ColumnSet(true)
            };
            queryGetSolutions.Criteria.AddCondition("ismanaged", ConditionOperator.Equal, false);
            queryGetSolutions.Criteria.AddCondition("isvisible", ConditionOperator.Equal, true);
            queryGetSolutions.Criteria.AddCondition("publisherid", ConditionOperator.NotIn, microsoftPublisherId);

            var solutions = Service.RetrieveMultiple(queryGetSolutions);
            return solutions.Entities.Select(e => e.GetAttributeValue<string>("friendlyname")).ToArray();
        }

        private Guid GetGuidForSolution(string solution)
        {
            if (!IsConnected)
            {
                return default;
            }

            var queryGetSolutions = new QueryExpression
            {
                EntityName = "solution",
                NoLock = true,
                TopCount = 1,
                ColumnSet = new ColumnSet()
            };
            queryGetSolutions.Criteria.AddCondition("friendlyname", ConditionOperator.Equal, solution);

            var solutions = Service.RetrieveMultiple(queryGetSolutions);
            if (solution == null || solutions.Entities == null || solutions.Entities.Count == 0)
            {
                return default;
            }

            var solutionId = solutions.Entities.FirstOrDefault();
            return solutionId?.Id ?? default;
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
            // Switch to the main thread - the call to AddCommand in ConnectToDynamics's constructor requires the UI thread.
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

            Connect();
        }

        private IVsOutputWindowPane GetGeneralPane()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var output = (IVsOutputWindowPane)this.package.GetServiceAsync(typeof(SVsGeneralOutputWindowPane)).Result;

            return output;
        }

        private void Connect()
        {
            var connectWindow = new ConnectWindow();
            connectWindow.ShowDialog();
        }

        private void EnsureConnected()
        {
            if (!IsConnected)
            {
                Connect();
            }
        }

        private void WriteToOutput(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var generalPane = GetGeneralPane();
            generalPane.OutputString(message);
            generalPane.Activate();
        }

        private void ExecuteUpdate(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            WriteToOutput("Executing Update from Dynamics.");

            WriteToOutput("Ensuring Connection." + Environment.NewLine);
            EnsureConnected();
            if (IsConnected)
            {
                WriteToOutput("We are connected!" + Environment.NewLine);
                var fileName = GetSelectedFileName();
                WriteToOutput($"File name to check: {fileName}." + Environment.NewLine);
                if (string.IsNullOrEmpty(fileName))
                {
                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        "Can not find file to update",
                        "Error!",
                        OLEMSGICON.OLEMSGICON_CRITICAL,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    return;
                }

                var content = UpdateFromDynamics(fileName);
                WriteToOutput($"Got content from Dynamics, length: {content.Length}." + Environment.NewLine);
                if (string.IsNullOrEmpty(content))
                {
                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        fileName,
                        "Failed to Update from Dynamics",
                        OLEMSGICON.OLEMSGICON_CRITICAL,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    return;
                }

                File.WriteAllText(fileName, content);
                WriteToOutput($"Wrote content to file, updating selected file." + Environment.NewLine);
                RefreshSelectedFile();
            }
        }

        private void ExecutePublish(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            EnsureConnected();
            if (IsConnected)
            {
                var fileName = GetSelectedFileName();
                if (string.IsNullOrEmpty(fileName))
                {
                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        "Can not find file to publish",
                        "Error!",
                        OLEMSGICON.OLEMSGICON_CRITICAL,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    return;
                }

                // Read the content
                var content = File.ReadAllText(fileName);
                if (!PublishContentToDynamics(content, fileName))
                {
                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        fileName,
                        "Failed to Publish To Dynamics",
                        OLEMSGICON.OLEMSGICON_CRITICAL,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }
            }
        }

        private string UpdateFromDynamics(string fileName)
        {
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            var webResource = GetWebResourceFromSolution(fileNameWithoutExtension, SelectedSolution, true);
            if (webResource == null)
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    "Searched using ",
                    "Failed to find unique WebResource To Update",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return string.Empty;
            }

            return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(webResource.GetAttributeValue<string>("content")));
        }

        private Entity GetWebResourceFromSolution(string name, string solution, bool getContent = false)
        {
            var columnSet = new ColumnSet("name");
            if (getContent)
            {
                columnSet.AddColumn("content");
            }

            var request = new QueryExpression("webresource") {ColumnSet = columnSet, NoLock = true};
            request.Criteria.AddCondition("webresourcetype", ConditionOperator.Equal, 3);
            request.Criteria.AddCondition("name", ConditionOperator.Equal, name);

            var solutionGuid = GetGuidForSolution(solution);
            var solutionLink = new LinkEntity("webresource", "solutioncomponent", "webresourceid", "objectid", JoinOperator.Inner);
            solutionLink.LinkCriteria.AddCondition("solutionid", ConditionOperator.Equal, solutionGuid);
            request.LinkEntities.Add(solutionLink);

            var results = Service.RetrieveMultiple(request);
            if (results.Entities.Count != 1)
            {
                WriteToOutput($"Got {results.Entities.Count} WebResource results:" + Environment.NewLine);
                foreach (var result in results.Entities)
                {
                    WriteToOutput($" {result["name"]}" + Environment.NewLine);
                }

                return null;
            }

            return results.Entities.FirstOrDefault();
        }

        private bool PublishContentToDynamics(string content, string fileName)
        {
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            var webResource = GetWebResourceFromSolution(fileNameWithoutExtension, SelectedSolution);
            if (webResource == null)
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    $"Searched using {fileName}",
                    "Failed to find unique WebResource To Update",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return false;
            }

            // Update content
            var webResourceToUpdate = new Entity("webresource");
            webResourceToUpdate["webresourceid"] = webResource["webresourceid"];
            webResourceToUpdate["content"] = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(content));
            
            // Save
            Service.Update(webResourceToUpdate);

            // Publish
            var webResourceXml = $"<importexportxml><webresources><webresource>{webResourceToUpdate.Id}</webresource></webresources></importexportxml>";
            var publishXmlRequest = new PublishXmlRequest
            {
                ParameterXml = string.Format(webResourceXml)
            };
            Service.Execute(publishXmlRequest);

            return true;
        }

        private void RefreshSelectedFile()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            var projectItem = GetSelectedProjectItem();
            if (projectItem == null)
            {
                return;
            }

            if (projectItem.IsOpen)
            {
               // If it has a Document, close it
                projectItem.Document?.Close(vsSaveChanges.vsSaveChangesNo);
            }

            // Re-open to show changes
            if (!projectItem.IsOpen)
            {
                try
                {
                    var window = projectItem.Open();
                    window.Activate();
                }
                catch (Exception)
                {
                    // Cannot open it - FINE!
                }
            }
        }

        private ProjectItem GetSelectedProjectItem()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var applicationObjectTask = this.ServiceProvider.GetServiceAsync(typeof(SDTE));
            if (!(applicationObjectTask.Result is DTE2 applicationObject))
            {
                return null;
            }

            var solutionExplorer = applicationObject.ToolWindows.SolutionExplorer;
            if (!(solutionExplorer.SelectedItems is UIHierarchyItem[] items) || items.Length == 0)
            {
                return null;
            }

            var file = items.FirstOrDefault();
            return file?.Object as ProjectItem;
        }

        private string GetSelectedFileName()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var projectItem = GetSelectedProjectItem();
            return projectItem?.FileNames[1];
        }
    }
}
