using System;
using System.IO;
using FluentAssertions;
using Microsoft.Test.Apex.Services;
using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.Shell;
using Microsoft.Test.Apex.VisualStudio.Debugger;
using Microsoft.Test.Apex.VisualStudio.Solution;
using Microsoft.Test.Apex.VisualStudio.FolderView;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Test.Apex.Services.StringGeneration;
using Microsoft.Test.Apex.VisualStudio.Shell.ToolWindows;
using Microsoft.VisualStudio.PlatformUI.Apex.NewProjectDialog;
using Microsoft.VisualStudio.PlatformUI.Apex.NewItemDialog;
using Microsoft.Test.Apex.VisualStudio.Settings;

namespace VsIde.IdeExp.GetToCodeTests
{
    [TestClass]
    public class NewProjectTests : VisualStudioHostTest
    {
        private NewProjectDialogTestService newProjectDialogTestService;
        private readonly VsSafeFileNameStringGenerator stringGenerator = new VsSafeFileNameStringGenerator() { MinLength = 3, MaxLength = 10 };

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();
            this.newProjectDialogTestService = this.GetAndCheckService<NewProjectDialogTestService>();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(5 * 60 * 1000)]
        public void CreateNewItem()
        {
            string templateID = "Microsoft.JavaScript.NodejsWebApp";
            NewItemDialogTestService newItemDialogTestService = this.GetAndCheckService<NewItemDialogTestService>();
            CreateNewProject(templateID);
            SolutionService solution = this.VisualStudio.ObjectModel.Solution;
            using (this.Scope.Enter("Add a new item to solution node."))
            {
                solution.SolutionExplorer.RootItem.Select();
                NewItemDialogTestExtension testExtension = newItemDialogTestService.GetExtension();

                testExtension.InvokeDialog();
                const int timeoutInSeconds = 5;
                bool result = testExtension.Verify.IsNewItemDialogDisplayWithinTimeSpan(TimeSpan.FromSeconds(timeoutInSeconds));
                _ = result.Should().BeTrue($"Failed to show new item dialog in {timeoutInSeconds} seconds.");
                
            }

        }

        [TestMethod]
        [Owner("IDE Experience")]
        public void TestItemTemplateLoad()
        {
            SolutionService solution = this.VisualStudio.ObjectModel.Solution;
            solution.CreateEmptySolution();
            var folder = solution.AddSolutionFolder();
            _ = folder.AddProject(ProjectLanguage.CSharp, ProjectTemplate.ConsoleApplication);

            NewItemDialogTestService newItemDialogTestService = this.GetAndCheckService<NewItemDialogTestService>();

            using (this.Scope.Enter(TestExecutionPoint.Scenario, "Open add new item dialog."))
            {
                NewItemDialogTestExtension testExtension = newItemDialogTestService.GetExtension();

                testExtension.InvokeDialog();

                const int timeoutInSeconds = 5;
                bool result = testExtension.Verify.IsNewItemDialogDisplayWithinTimeSpan(TimeSpan.FromSeconds(timeoutInSeconds));
                _ = result.Should().BeTrue($"Failed to show new item dialog in {timeoutInSeconds} seconds.");

                bool result1 = testExtension.Verify.IsTemplateLoadCompleteWithinTimeSpan(TimeSpan.FromSeconds(timeoutInSeconds));
                _ = result1.Should().BeTrue($"Failed to load item template in {timeoutInSeconds} seconds.");

                testExtension.CloseDialog();
            }
        }

        [TestMethod, Owner("IDE Experience"), Timeout(5 * 60 * 1000)]
        public void RestoreDocumentsTestMethod()
        {
            string templateID = "Microsoft.Web.Mvc``C#";

            CreateNewProject(templateID);
            SolutionService solution = this.VisualStudio.ObjectModel.Solution;
            SolutionExplorerService exporerService = solution.SolutionExplorer;

            this.Services.Synchronization.WaitFor(TimeSpan.FromSeconds(60), () => solution.LoadedProjectCount == 1);

            SolutionService solutionService = VisualStudio.ObjectModel.Solution;
            ProjectTestExtension project = solutionService.Projects[0];
            string path = project.FullPath;
            project[@"Views\_ViewImports.cshtml"].Open();
            project["Program.cs"].Open();
            string programPath = project["Program.cs"].FullPath;
            string viewPath = project[@"Views\_ViewImports.cshtml"].FullPath;
            solutionService.Close();
            Delay(2);

            solutionService.OpenProject(path);
            this.Verify.IsTrue(this.VisualStudio.ObjectModel.WindowManager.IsDocumentOpen(programPath, DocumentLogicalView.Code));
            this.Verify.IsTrue(this.VisualStudio.ObjectModel.WindowManager.IsDocumentOpen(viewPath, DocumentLogicalView.Code));
            VisualStudio.Stop();
        }
        [TestMethod, Owner("IDE Experience"), Timeout(5 * 60 * 1000)]
        public void EnterExitDebugMode()
        {
            string templateID = "microsoft.CSharp.WPFApplication";
            CreateNewProject(templateID);
            SolutionService solution = this.VisualStudio.ObjectModel.Solution;
            ProjectTestExtension project = solution.Projects[0];
            string mainWindowPath = project["MainWindow.xaml"].FullPath;
            DebuggerService debugger = this.VisualStudio.ObjectModel.Debugger;
            debugger.Start();
            DebuggerMode debugMode = debugger.CurrentMode;
            this.Verify.IsTrue(debugMode == DebuggerMode.RunMode);
            debugger.Stop();
            Delay(2);
            this.Verify.IsTrue(this.VisualStudio.ObjectModel.WindowManager.IsDocumentOpen(mainWindowPath, DocumentLogicalView.Code));
            this.VisualStudio.Stop();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(500 * 60 * 1000)]
        public void ChangingSettings()
        {
            string templateID = "microsoft.CSharp.ConsoleApplication";
            CreateNewProject(templateID);

            SolutionService solution = this.VisualStudio.ObjectModel.Solution;
            uint oldtabsize = this.VisualStudio.ObjectModel.Settings.TextEditor.GetTabSize(FileType.CS);
            uint oldindentsize = this.VisualStudio.ObjectModel.Settings.TextEditor.GetIndentSize(FileType.CS);
            this.VisualStudio.ObjectModel.Settings.TextEditor.SetIndentStyle(TextEditorSettings.IndentStyle.Smart, FileType.CS);
            uint newtabszie = 7;
            this.VisualStudio.ObjectModel.Settings.TextEditor.SetTabSize(newtabszie, FileType.CS);
            uint oldtabsize01 = this.VisualStudio.ObjectModel.Settings.TextEditor.GetTabSize(FileType.CS);
            SolutionExplorerService exporerService = solution.SolutionExplorer;
            this.VisualStudio.ObjectModel.Settings.TextEditor.SetTabOption(TextEditorSettings.TabOption.InsertSpaces, FileType.CS);
            
        }

        [TestMethod, Owner("IDE Experience"), Timeout(5 * 60 * 1000)]
        public void IdeShowAllFilesTest()
        {
            string templateID = "microsoft.CSharp.ConsoleApplication";
            CreateNewProject(templateID);

            SolutionService solution = this.VisualStudio.ObjectModel.Solution;
            SolutionExplorerService exporerService = solution.SolutionExplorer;

            this.Services.Synchronization.WaitFor(TimeSpan.FromSeconds(60), () => solution.LoadedProjectCount == 1);
            exporerService.ShowAllFiles();

            exporerService.Select("bin");
            SolutionExplorerItemTestExtension projectBin = exporerService.FindItemRecursive("bin");
            this.Verify.IsNotNull(projectBin);
            SolutionExplorerItemTestExtension projectObj = exporerService.FindItemRecursive("obj");
            this.Verify.IsNotNull(projectObj);
            exporerService.ShowAllFiles();
            VisualStudio.Stop();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(5 * 60 * 1000)]
        public void FolderSolutionViewSwitch()
        {
            string templateID = "Microsoft.Common.Console``C#";
            CreateNewProject(templateID);

            SolutionService solutionService = VisualStudio.ObjectModel.Solution;
            this.Services.Synchronization.WaitFor(TimeSpan.FromSeconds(60), () => solutionService.LoadedProjectCount == 1);
            SolutionExplorerService exporerService = solutionService.SolutionExplorer;
            string displayname = this.VisualStudio.ObjectModel.Solution.DisplayName;
            string name = this.VisualStudio.ObjectModel.Solution.Name;
            string filename = this.VisualStudio.ObjectModel.Solution.FileName;
            exporerService.ToggleToFolderView();
            FolderViewService folderViewService = VisualStudio.ObjectModel.FolderView;
            folderViewService.Verify.IsInFolderView();
            if (folderViewService.FolderViewTree != null)
            {
                string[] paths = folderViewService.FolderViewTree.GetAllItemPaths();

                bool isContainsln = false;
                for (int i = 0; i < paths.Length; i++)
                {
                    if (paths[i].Contains("sln"))
                    {
                        isContainsln = true;
                        break;
                    }
                }
                this.Verify.IsTrue(isContainsln);
                FolderViewItemTestExtension solutionItem;
                foreach (string pathitem in paths)
                {
                    folderViewService.FolderViewTree.Select(pathitem);
                    if (pathitem.Contains("sln"))
                    {
                        solutionItem = folderViewService.FolderViewTree.SelectedItem;
                        solutionItem.DoubleClick();
                        folderViewService.Verify.IsNotInFolderView();
                    }
                }
            }

            this.Verify.IsTrue(solutionService.LoadedProjectCount == 1);
            VisualStudio.Stop();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(5 * 60 * 1000)]
        public void NPDFilterProjectTemplates()
        {
            IKeyboardAutomationService keyboardService = this.VisualStudio.Get<IKeyboardAutomationService>();
            string searchTemplate = "React";
            NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
            CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();
            createProjectDialogTestExtension.SearchTemplates(searchTemplate);
            int count = createProjectDialogTestExtension.GetTemplateShownCount();
            for (int i = 0; i < count; i++)
            {
                this.Verify.IsTrue(createProjectDialogTestExtension.SelectedProjectTemplateName.Contains(searchTemplate));
                if (i == 0)
                {
                    Delay(1);
                    keyboardService.TypeKey(KeyboardKey.Tab);
                    Delay(1);
                    keyboardService.TypeKey(KeyboardKey.Tab);
                }
                if (i != count - 1)
                {
                    Delay(1);
                    keyboardService.TypeKey(KeyboardKey.Down);
                }
            }
            searchTemplate = "Angular";
            Delay(2);
            createProjectDialogTestExtension.SearchTemplates(searchTemplate);
            Delay(2);
            int angularCount = createProjectDialogTestExtension.GetTemplateShownCount();
            this.Verify.IsTrue(angularCount == 2);

            Delay(2);
            searchTemplate = "AA";
            createProjectDialogTestExtension.SearchTemplates(searchTemplate);
            this.Verify.IsTrue(createProjectDialogTestExtension.GetTemplateShownCount() == 0);
            createProjectDialogTestExtension.ClearSearch();
            ConfigProjectDialogTestExtension configPage = createProjectDialogTestExtension.NavigateToConfigProjectPage();
            configPage.FinishConfiguration();
            configPage.FinishConfiguration();
            Delay(2);
            this.VisualStudio.Stop();
        }

        private void CreateNewProject(string templateID)
        {
            NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
            CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();

            string projectLocation = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "source", "repos", stringGenerator.Generate());
            string solutionName = stringGenerator.Generate();
            string projectName = stringGenerator.Generate();
            createProjectDialogTestExtension.Verify.ProjectCreationDialogVisibilityIs(true).Should().BeTrue();

            if (!createProjectDialogTestExtension.TrySelectTemplateWithId(templateID))
                return;

            ConfigProjectDialogTestExtension configTestExtension = createProjectDialogTestExtension.NavigateToConfigProjectPage();
            configTestExtension.Verify.ProjectConfigurationDialogVisibilityIs(true).Should().BeTrue();

            // Assign valid SolutionName
            configTestExtension.SolutionName = solutionName;

            // Assign valid ProjectName
            configTestExtension.ProjectName = projectName;

            // Assign valid Location
            configTestExtension.Location = projectLocation;

            configTestExtension.FinishConfiguration();
            if (!templateID.Contains("Application") && !templateID.Contains("Node"))
                configTestExtension.FinishConfiguration();

        }

        private void Delay(int delayTime)
        {
            DateTime start = DateTime.Now;
            int spanSecond;
            do
            {
                TimeSpan span = DateTime.Now - start;
                spanSecond = span.Seconds;
            }
            while (spanSecond < delayTime);
        }
        private T GetAndCheckService<T>() where T : class
        {
            var testService = this.VisualStudio.Get<T>();
            this.ThrowIfNull(testService, $"{typeof(T).Name} not found.");
            return testService;
        }

        private void ThrowIfNull(object obj, string message)
        {
            obj.Should().NotBeNull(because: message);
        }

    }
}
