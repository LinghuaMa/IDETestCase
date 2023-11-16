using System;
using System.Linq;
using VsIde.IdeExp.GetToCodeTests;
using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.Solution;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.PlatformUI.Apex.GetToCode;
using Microsoft.Test.Apex.VisualStudio.Shell.ToolWindows;
using Microsoft.Test.Apex.VisualStudio.Shell;
using System.IO;
using Microsoft.Test.Apex.VisualStudio.Settings;

namespace VsIde.IdeExp.GetToCodeTests
{
    [TestClass]
    public class IdeUserScenarioTests : GetToCodeTestBase
    {
        protected override StartWindowStartupOption StartWindowStartupOption => StartWindowStartupOption.UseVisualStudioSetting;

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();
        }

        [TestMethod]
        [Owner("IDE Experience")]
        public void CollapseTheSelectedNodeUsingContextMenu()
        {
            string clonePath;
            string cloneURL = @"https://github.com/IDEExperienceTestSolution/CollapseCommandTestApp.git";
            SolutionExplorerService solutionExplorerService = GetSolutionExplorer(cloneURL, out clonePath).SolutionExplorer;

            // Select a project node which is expanded
            try
            {
                solutionExplorerService.Select("BlazorApp1");
            }
            catch
            {
                Delay(3);
                solutionExplorerService.Select("BlazorApp1");
            }

            SolutionExplorerItemTestExtension projectNode = solutionExplorerService.FindItemRecursive("BlazorApp1");
            projectNode.Expand();
            this.Verify.IsTrue(projectNode.IsExpanded, "We expect the node of BlazorApp1 is expanded");

            // Collapse the node with command button in context menu by right click the node         
            CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
            CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
            commandBtns.First(b => b.Name == "Collapse All Descendants").Execute();
            this.Verify.IsTrue(!projectNode.IsExpanded, "We expect the node of BlazorApp1 is collapsed");

            // Verify the command button of 'Collapse All Descendants' in context menu is disable when the node is collapsed
            solutionExplorerService.Select("BlazorApp1");
            projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
            commandBtns = projectContextMenu.CommandBarButtonsRecursive;
            this.Verify.IsTrue(!commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);

            // Expand a solution folder
            solutionExplorerService.Select("Solution Items");
            SolutionExplorerItemTestExtension solutionFolderNode = solutionExplorerService.FindItemRecursive("Solution Items");
            solutionFolderNode.Expand();
            this.Verify.IsTrue(solutionFolderNode.IsExpanded, "We expect the node of 'Solution Items' is expanded");

            // Collapse the node with command button in context menu by right click the node
            CommandBarTestExtension solutionFolderContextMenu = VisualStudio.Get<CommandBarsService>().SolutionFolderContextMenu;
            CommandBarButtonTestExtension[] solutionFolderCommandBtns = solutionFolderContextMenu.CommandBarButtonsRecursive;
            solutionFolderCommandBtns.First(b => b.Name == "Collapse All Descendants").Execute();
            this.Verify.IsTrue(!solutionFolderNode.IsExpanded, "We expect the node of 'Solution Items' is collapsed");

            VisualStudio.Stop();

            SetPermissionsToNormalOnDotGitFolder(clonePath);
            Directory.Delete(clonePath, true);
        }

        [TestMethod, Owner("IDE Experience")]
        public void FilteringToChangedFiles()
        {
            string clonePath;
            string cloneURL = "https://github.com/Microsoft/TailwindTraders-Desktop";
            SolutionService solutionService = GetSolutionExplorer(cloneURL, out clonePath);
            string solutionPath = Path.Combine(clonePath, @"Source\CouponReader.NETFramework.sln");
            solutionService.WaitForFullyLoadedOnOpen = true;
            solutionService.WaitForCpsProjectsFullyLoaded = true;
            solutionService.Open(solutionPath);
            solutionService.WaitForFullyLoaded();
            this.Services.Synchronization.WaitFor(TimeSpan.FromSeconds(60), () =>
            {
                return this.VisualStudio.ObjectModel.Solution.IsOpen;
            });
            string projectPath = Path.Combine(clonePath, @"Source\\CouponReader.Tests\\CouponReader.Tests.csproj");
            ProjectTestExtension project = solutionService.OpenProject(projectPath);
            TextEditorDocumentWindowTestExtension editor = project["CouponsServiceTests.cs"].GetDocumentAsTextEditor();

            editor.Editor.Edit.InsertText(12, 13, Environment.NewLine);
            editor.Editor.Edit.InsertText(12, 13, @"//Test01");
            editor.Save();
            solutionService.Close();
            Delay(2);

            project = solutionService.OpenProject(projectPath);
            editor = project["CouponsServiceTests.cs"].GetDocumentAsTextEditor();
            editor.Editor.Caret.MoveToAndSelectExpression("Test01");
            this.Verify.IsTrue(editor.Editor.Caret.Verify.IsAtLine(12));
            this.SetStartupOption(EnvironmentSettings.StartupOption.ShowEmptyEnvironment);
            VisualStudio.Stop();
            SetPermissionsToNormalOnDotGitFolder(clonePath);
            Directory.Delete(clonePath, true);
        }

        [TestMethod, Owner("IDE Experience")]
        public void VerifyVisibilityOfCommandForCollapsedNode()
        {
            string clonePath;
            string cloneURL = @"https://github.com/IDEExperienceTestSolution/CollapseCommandTestApp.git";
            SolutionExplorerService solutionExplorerService = GetSolutionExplorer(cloneURL, out clonePath).SolutionExplorer;
            Delay(3);
            using (Scope.Enter("[Solution root] Verify the Collapse All Descendants command is visible and always enabled"))
            {
                solutionExplorerService.RootItem.Select();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
                commandBtns.First(b => b.Name == "Collapse All Descendants").Execute();
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            using (Scope.Enter("[Solution Folder] Verify the Collapse All Descendants command is visible but disabled"))
            {
                solutionExplorerService.Select("Solution Items");
                SolutionExplorerItemTestExtension solutionFolderNode = solutionExplorerService.FindItemRecursive("Solution Items");
                solutionFolderNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                this.Verify.IsTrue(solutionFolderNode.IsExpanded, "We expect the node of 'Solution Items' is collapsed");
                commandBtns.First(b => b.Name == "Collapse All Descendants").Execute();
                this.Verify.IsFalse(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            using (Scope.Enter("[Project root] Verify the Collapse All Descendants command is visible but disabled"))
            {
                solutionExplorerService.Select("BlazorApp1");
                SolutionExplorerItemTestExtension projectNode = solutionExplorerService.FindItemRecursive("BlazorApp1");
                projectNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                commandBtns.First(b => b.Name == "Collapse All Descendants").Execute();
                this.Verify.IsFalse(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }

            using (Scope.Enter("[Project Forder] Verify the Collapse All Descendants command is visible but disabled"))
            {
                solutionExplorerService.Select("Data");
                SolutionExplorerItemTestExtension projectFolderNode = solutionExplorerService.FindItemRecursive("Data");
                projectFolderNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                commandBtns.First(b => b.Name == "Collapse All Descendants").Execute();
                this.Verify.IsFalse(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            using (Scope.Enter("[A descendant nodes] Verify the Collapse All Descendants command is visible but disabled"))
            {
                solutionExplorerService.Select("appsettings.json");
                SolutionExplorerItemTestExtension projectFolderNode = solutionExplorerService.FindItemRecursive("appsettings.json");
                projectFolderNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                commandBtns.First(b => b.Name == "Collapse All Descendants").Execute();
                this.Verify.IsFalse(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            VisualStudio.Stop();
            SetPermissionsToNormalOnDotGitFolder(clonePath);
            Directory.Delete(clonePath, true);
        }

        [TestMethod, Owner("IDE Experience")]
        public void VerifyVisibilityOfCommandForExpandedNode()
        {
            string clonePath;
            string cloneURL = @"https://github.com/IDEExperienceTestSolution/CollapseCommandTestApp.git";
            SolutionExplorerService solutionExplorerService = GetSolutionExplorer(cloneURL, out clonePath).SolutionExplorer;
            Delay(3);
            using (Scope.Enter("[Solution root] Verify the Collapse All Descendants command is enabled"))
            {
                solutionExplorerService.RootItem.Select();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            using (Scope.Enter("[Solution Folder] Verify the Collapse All Descendants command is enabled"))
            {
                solutionExplorerService.Select("Solution Items");
                SolutionExplorerItemTestExtension solutionFolderNode = solutionExplorerService.FindItemRecursive("Solution Items");
                solutionFolderNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                this.Verify.IsTrue(solutionFolderNode.IsExpanded, "We expect the node of 'Solution Items' is collapsed");
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            using (Scope.Enter("[Project root] Verify the Collapse All Descendants command is enabled"))
            {
                solutionExplorerService.Select("BlazorApp1");
                SolutionExplorerItemTestExtension projectNode = solutionExplorerService.FindItemRecursive("BlazorApp1");
                projectNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }

            using (Scope.Enter("[Project Forder] Verify the Collapse All Descendants command is enabled"))
            {
                solutionExplorerService.Select("Data");
                SolutionExplorerItemTestExtension projectFolderNode = solutionExplorerService.FindItemRecursive("Data");
                projectFolderNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            using (Scope.Enter("[A descendant nodes] Verify the Collapse All Descendants command is enabled"))
            {
                solutionExplorerService.Select("appsettings.json");
                SolutionExplorerItemTestExtension projectFolderNode = solutionExplorerService.FindItemRecursive("appsettings.json");
                projectFolderNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            VisualStudio.Stop();
            SetPermissionsToNormalOnDotGitFolder(clonePath);
            Directory.Delete(clonePath, true);
        }

        [TestMethod, Owner("IDE Experience")]
        public void VerifyVisibilityOfCommandWhenCollapsedThenExpandedAgain()
        {
            string clonePath;
            string cloneURL = @"https://github.com/IDEExperienceTestSolution/CollapseCommandTestApp.git";
            SolutionExplorerService solutionExplorerService = GetSolutionExplorer(cloneURL, out clonePath).SolutionExplorer;
            Delay(3);
            using (Scope.Enter("[Solution root] Verify the Collapse All Descendants command is enabled"))
            {
                solutionExplorerService.RootItem.Select();
                solutionExplorerService.RootItem.Expand();  
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                commandBtns.First(c => c.Name == "Collapse All Descendants").Execute();
                solutionExplorerService.RootItem.Expand();
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            using (Scope.Enter("[Solution Folder] Verify the Collapse All Descendants command is enabled"))
            {
                solutionExplorerService.Select("Solution Items");
                SolutionExplorerItemTestExtension solutionFolderNode = solutionExplorerService.FindItemRecursive("Solution Items");
                solutionFolderNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                commandBtns.First(c => c.Name == "Collapse All Descendants").Execute();
                solutionFolderNode.Expand();
                this.Verify.IsTrue(solutionFolderNode.IsExpanded, "We expect the node of 'Solution Items' is expanded");
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            using (Scope.Enter("[Project root] Verify the Collapse All Descendants command is enabled"))
            {
                solutionExplorerService.Select("BlazorApp1");
                SolutionExplorerItemTestExtension projectNode = solutionExplorerService.FindItemRecursive("BlazorApp1");
                projectNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                commandBtns.First(c => c.Name == "Collapse All Descendants").Execute();
                projectNode.Expand();
                this.Verify.IsTrue(projectNode.IsExpanded, "We expect the node of 'Project root' is expanded");
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }

            using (Scope.Enter("[Project Forder] Verify the Collapse All Descendants command is enabled"))
            {
                solutionExplorerService.Select("Data");
                SolutionExplorerItemTestExtension projectFolderNode = solutionExplorerService.FindItemRecursive("Data");
                projectFolderNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                commandBtns.First(c => c.Name == "Collapse All Descendants").Execute();
                projectFolderNode.Expand();
                this.Verify.IsTrue(projectFolderNode.IsExpanded, "We expect the node of 'Project Forder' is expanded");
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            using (Scope.Enter("[A descendant nodes] Verify the Collapse All Descendants command is enabled"))
            {
                solutionExplorerService.Select("appsettings.json");
                SolutionExplorerItemTestExtension projectFolderNode = solutionExplorerService.FindItemRecursive("appsettings.json");
                projectFolderNode.Expand();
                CommandBarTestExtension projectContextMenu = VisualStudio.Get<CommandBarsService>().ProjectContextMenu;
                CommandBarButtonTestExtension[] commandBtns = projectContextMenu.CommandBarButtonsRecursive;
                this.Verify.IsTrue(commandBtns.First(c => c.Name == "Collapse All Descendants").IsEnabled);
            }
            VisualStudio.Stop();
            SetPermissionsToNormalOnDotGitFolder(clonePath);
            Directory.Delete(clonePath, true);
        }

        [TestMethod, Owner("IDE Experience")]
        public void OpenLargeSolutionCase()
        {
            string clonePath;
            string cloneURL= "https://github.com/dotnet/project-system";
            SolutionService solutionService = GetSolutionExplorer(cloneURL, out clonePath);
            string solutionPath = Path.Combine(clonePath, @"ProjectSystem.sln");
            solutionService.WaitForFullyLoadedOnOpen = true;
            solutionService.WaitForCpsProjectsFullyLoaded = true;
            solutionService.Open(solutionPath);
            solutionService.WaitForFullyLoaded();
            this.Services.Synchronization.WaitFor(TimeSpan.FromSeconds(10), () =>
            {
                return this.VisualStudio.ObjectModel.Solution.IsOpen;
            });
            int projectCount= solutionService.LoadedProjectCount;
            this.Verify.IsTrue(projectCount == 16);
            solutionService.Verify.LoadedProjectCountEquals(16);
            solutionService.Verify.ProjectCountEquals(16);
            SolutionExplorerService solutionExplorerService = solutionService.SolutionExplorer;
            string rootName= solutionExplorerService.RootItem.Name;
            solutionExplorerService.Verify.RootItemNameEquals("Solution 'ProjectSystem' ‎\0(16 of 16 projects)");
            solutionExplorerService.Verify.SolutionTextContains("ProjectSystem");
            VisualStudio.Stop();
            SetPermissionsToNormalOnDotGitFolder(clonePath);
            Directory.Delete(clonePath, true);
        }

        private SolutionService GetSolutionExplorer(string cloneURL, out string clonePath)
        {
            GetToCodeDialogTestExtension g2cDialogTestExtension = this.GetToCodeService.GetCurrentGetToCodeDialog();
            CloneDialogTestExtension cloneDialogTestExtension = g2cDialogTestExtension.NavigateToCloneDialog();
            cloneDialogTestExtension.SetCloneUrlTo(cloneURL);
            clonePath = cloneDialogTestExtension.CloneLocalPath;
            cloneDialogTestExtension.BeginClone();
            cloneDialogTestExtension.WaitForCloneToComplete();
            this.GetToCodeService.Verify.GetToCodeDialogIsClosedWithinTimeSpan(TimeSpan.FromSeconds(5));

            this.Services.Synchronization.WaitFor(() => this.VisualStudio.ObjectModel.Solution != null);
            return this.VisualStudio.ObjectModel.Solution;
        }


        private void SetPermissionsToNormalOnDotGitFolder(string path)
        {
            try
            {
                var directoryInfo = new DirectoryInfo(path);
                directoryInfo.Attributes &= ~FileAttributes.ReadOnly;
                string[] files = Directory.GetFiles(path);

                foreach (string file in files)
                {
                    File.SetAttributes(file, FileAttributes.Normal);
                }

                string[] directories = Directory.GetDirectories(path);

                foreach (string directory in directories)
                {
                    SetPermissionsToNormalOnDotGitFolder(directory);
                }
            }
            catch (Exception e)
            {
                this.Logger.WriteWarning(e.ToString());
            }
        }

        private void SetStartupOption(EnvironmentSettings.StartupOption option)
        {
            var settings = this.GetAndCheckService<SettingsService>();
            settings.Environment.General.OnStartup = option;
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
    }
}
