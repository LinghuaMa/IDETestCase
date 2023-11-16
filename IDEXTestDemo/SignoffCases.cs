//using Microsoft.Test.Apex.Services.StringGeneration;
//using Microsoft.Test.Apex.VisualStudio;
//using Microsoft.Test.Apex.VisualStudio.Solution;
//using Microsoft.VisualStudio.TestTools.UnitTesting;
//using System;
//using System.Collections.Generic;
//using FluentAssertions;
//using Microsoft.Test.Apex.VisualStudio.Shell;
//using Microsoft.VisualStudio.PlatformUI.Apex.NewProjectDialog;
//using System.Threading;
//using Microsoft.Test.Apex.VisualStudio.Shell.ToolWindows;
//using Microsoft.Test.Apex.Services;
//using Microsoft.VisualStudio.PlatformUI.Apex.NewItemDialog;
//using Microsoft.VisualStudio.PlatformUI.Apex.GetToCode;
//using System.IO;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

using Microsoft.Test.Apex.Services.StringGeneration;
using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.FeatureFlags;
using Microsoft.Test.Apex.VisualStudio.FolderView;
using Microsoft.Test.Apex.VisualStudio.Settings;
using Microsoft.Test.Apex.VisualStudio.Shell;
using Microsoft.Test.Apex.VisualStudio.Solution;
using Microsoft.VisualStudio.PlatformUI.Apex.GetToCode;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.PlatformUI.Apex.NewProjectDialog;
using FluentAssertions;

namespace VsIde.IdeExp.SignoffTests
{
    [TestClass]
    public class SignoffCases : VisualStudioHostTest
    {
        private NewProjectDialogTestService newProjectDialogTestService;
        //private NewItemDialogTestService newItemDialogTestService;
        //private StartPageTestService startPageTestService = null;
        protected GetToCodeTestService GetToCodeService;

        private List<string> foldersToCleanUp = new List<string>();
        private readonly VsSafeFileNameStringGenerator stringGenerator = new VsSafeFileNameStringGenerator() { MinLength = 3, MaxLength = 10 };

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();
            this.newProjectDialogTestService = this.GetAndCheckService<NewProjectDialogTestService>();
            //newItemDialogTestService= this.GetAndCheckService<NewItemDialogTestService>();
            //startPageTestService= this.GetAndCheckService<StartPageTestService>();
            this.GetToCodeService = this.GetAndCheckService<GetToCodeTestService>();
        }

        [TestMethod, Owner("IDE Experience")]
        public void ValidateAutogenerationOfLocalPathFromURIs()
        {
            string repoName = stringGenerator.Generate();
            string ValidCloneUrl = $"https://github.com/octocat/{repoName}";
            string ValidSSHUrl = $"git@github.com:octocat/{repoName}";
            const string InvalidCloneUrl = "   "; // We only validate empty string, null and whitespace

            GetToCodeDialogTestExtension g2cDialogTestExtension = this.GetToCodeService.GetCurrentGetToCodeDialog();
            CloneDialogTestExtension cloneDialogTestExtension = g2cDialogTestExtension.NavigateToCloneDialog();

            cloneDialogTestExtension.Verify.CloneDialogIsOpenEquals(true);
            cloneDialogTestExtension.Verify.CanCloneEquals(false);

            string defaultLocalPath = cloneDialogTestExtension.CloneLocalPath;
            string expectedLocalPath = Path.Combine(defaultLocalPath, repoName);

            cloneDialogTestExtension.SetCloneUrlTo(InvalidCloneUrl);
            cloneDialogTestExtension.Verify.CloneLocalPathEquals(defaultLocalPath);
            cloneDialogTestExtension.Verify.CanCloneEquals(false);

            cloneDialogTestExtension.SetCloneUrlTo(ValidCloneUrl);
            cloneDialogTestExtension.Verify.CloneLocalPathEquals(expectedLocalPath);
            cloneDialogTestExtension.Verify.CanCloneEquals(true);

            cloneDialogTestExtension.SetCloneUrlTo(ValidSSHUrl);
            cloneDialogTestExtension.Verify.CloneLocalPathEquals(expectedLocalPath);
            cloneDialogTestExtension.Verify.CanCloneEquals(true);

            g2cDialogTestExtension.CloseDialog();
        }

        [TestMethod]
        [Owner("IDE Experience")]
        public void OpenLargeSolution()
        {
            string repoName = stringGenerator.Generate();
            string ValidCloneUrl = $"https://github.com/octocat/{repoName}";
            string ValidSSHUrl = $"git@github.com:octocat/{repoName}";
            const string InvalidCloneUrl = "   "; // We only validate empty string, null and whitespace

            GetToCodeDialogTestExtension g2cDialogTestExtension = this.GetToCodeService.GetCurrentGetToCodeDialog();
            CloneDialogTestExtension cloneDialogTestExtension = g2cDialogTestExtension.NavigateToCloneDialog();
            CloneDialogTestExtension cloneDialog = this.GetToCodeService.GetCurrentGetToCodeDialog().NavigateToCloneDialog();
            cloneDialogTestExtension.SetCloneUrlTo("https://github.com/dotnet/project-system");
            cloneDialogTestExtension.SetCloneLocalPathTo("c:\\test");
            cloneDialogTestExtension.BeginClone();
            cloneDialogTestExtension.WaitForCloneToComplete();

        }

        [TestMethod]
        [Owner("IDE Experience")]
        public void RestoreDocumentsTestMethod()
        {
            NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
            CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();

            string templateID = "Microsoft.Web.Mvc``C#";
            string projectLocation = Environment.ExpandEnvironmentVariables($@"%USERPROFILE%\source\repos\{stringGenerator.Generate()}");
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

            this.foldersToCleanUp.Add(projectLocation);

            configTestExtension.FinishConfiguration();
            if (!templateID.Contains("Application"))
                configTestExtension.FinishConfiguration();
            Delay(9);
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

        [TestMethod]
        [Owner("IDE Experience")]
        public void CreateNewItem()
        {
            NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
            CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();

            string templateID = "Microsoft.JavaScript.NodejsWebApp";
            string projectLocation = Environment.ExpandEnvironmentVariables($@"%USERPROFILE%\source\repos\{stringGenerator.Generate()}");
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

            this.foldersToCleanUp.Add(projectLocation);

            configTestExtension.FinishConfiguration();
            //if (!templateID.Contains("Application"))
            //    configTestExtension.FinishConfiguration();
            Delay(9);

            //TabItem ;

            SolutionService solutionService=VisualStudio.ObjectModel.Solution;
            VisualStudio.ObjectModel.Solution.SolutionExplorer.RootItem.Select();
            VisualStudio.ObjectModel.Solution.SolutionExplorer.RootItem.RightClick();
            Thread.Sleep(TimeSpan.FromSeconds(5));
            //for (int i = 0; i <= 6; i++)
            //{
            //    this.VisualStudio.ObjectModel.Solution.SolutionExplorer.KeyboardService.TypeKey(KeyboardModifier.Control, KeyboardKey.Tab);
            //    Thread.Sleep(TimeSpan.FromMilliseconds(500)); // Give a half second delay between each keypress so none of the keypresses are ignored
            //}
            //this.VisualStudio.ObjectModel.Solution.SolutionExplorer.KeyboardService.TypeKey(KeyboardKey.Enter);
            ProjectTestExtension project = solutionService.Projects[0];
            //project
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

        //public void ClickContextMenuOnItem(SolutionExplorerItemTestExtension item, int menuItemIndex)
        //{
        //    item.RightClick();
        //    Thread.Sleep(TimeSpan.FromSeconds(5));
        //    for (int i = 0; i <= menuItemIndex; i++)
        //    {
        //        this.VisualStudio.ObjectModel.Solution.SolutionExplorer.KeyboardService.TypeKey(KeyboardModifier.Control, KeyboardKey.Tab);
        //        Thread.Sleep(TimeSpan.FromMilliseconds(500)); // Give a half second delay between each keypress so none of the keypresses are ignored
        //    }
        //    this.VisualStudio.ObjectModel.Solution.SolutionExplorer.KeyboardService.TypeKey(KeyboardKey.Enter);
        //}

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

