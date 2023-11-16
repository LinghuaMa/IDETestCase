using System;

using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.Shell;
//using Microsoft.VisualStudio.Ide.Tests;
using Microsoft.VisualStudio.PlatformUI.Apex.GetToCode;
using Microsoft.VisualStudio.PlatformUI.Apex.NewProjectDialog;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VsIde.IdeExp.GetToCodeTests
{
    /// <summary>
    /// This test class contains Start window tests relying on
    /// <see cref="VisualStudioHostBaseConfiguration.StartWindowStartupOption"/> to force Start window on start up,
    /// regardless of current user setting.
    /// </summary>
    [TestClass]
    public class StartWindowEnabledUserScenarioTests : GetToCodeTestBase
    {
        // This command set should be in sync with Microsoft.VisualStudio.PlatformUI.MruPackageGuids in
        // src\env\shell\UIInternal\Packages\MRU\Package\Guids.cs
        private readonly Guid recentProjectsAndSolutionsCommandSet = new Guid("{001bd6e5-e2cd-4f47-98d4-2d39215359fc}");
        private const uint firstMruItemCmdId = 1;

        [TestMethod]
        [Owner(GetToCodeTestBase.IdeExperienceEngineersAlias)]
        [Timeout(GetToCodeTestBase.FiveMinutesInMs)]
        //[TestCategory(IDETestCategories.Optimization)]
        public void ClosingStartWindowShouldShowMainWindow()
        {
            GetToCodeDialogTestExtension testExtension = this.GetToCodeService.GetCurrentGetToCodeDialog();
            testExtension.CloseDialog();

            this.VisualStudio.ObjectModel.MainWindow.Verify.IsShown();
        }

        [TestMethod]
        [Owner(GetToCodeTestBase.IdeExperienceEngineersAlias)]
        [Timeout(GetToCodeTestBase.FifteenMinutesInMs)]
        //[TestCategory(IDETestCategories.Optimization)]
        public void SearchTemplateCreateProjectAndOpenProjectFromMruShouldSucceed()
        {
            GetToCodeDialogTestExtension getToCode = this.GetToCodeService.GetCurrentGetToCodeDialog();

            CreateProjectDialogTestExtension createProject = getToCode.NavigateToNewProjectDialog();

            this.SearchTemplates(createProject, "asp .net", isFirstSearch: true);

            // Search for C# language tag to ensure results on any VS localized builds. This test creates a project
            // and loads it from MRU later so mock templates data does not work well.
            this.SearchTemplates(createProject, "net framework");

            if (!createProject.TrySelectTemplateWithId("microsoft.CSharp.ConsoleApplication"))
                return;

            ConfigProjectDialogTestExtension configProject = createProject.NavigateToConfigProjectPage();
            if (!configProject.Verify.ProjectConfigurationDialogVisibilityIs(true))
                return;

            configProject.Location = this.TestExecutionDirectory.FullName;
            configProject.ProjectInSolutionDirectory = false;
            if (!configProject.Verify.CanFinishConfigurationEquals(true))
                return;

            configProject.FinishConfiguration();

            this.WaitForSolutionLoad();

            //Check if the top MRU item in the File > Recent Projects and Solutions menu is enabled.
            CommandQueryResult result = this.VisualStudio.ObjectModel.Commanding.QueryStatusCommand(recentProjectsAndSolutionsCommandSet, firstMruItemCmdId);
            Assert.IsTrue(result.IsEnabled, "Top MRU item was expected to be found in the File > Recent Projects and solutions menu.");
            Assert.IsTrue(result.IsVisible, "Top MRU item was expected to be visible in the File > Recent Projects and solutions menu.");

            this.RestartVisualStudio();

            getToCode = this.GetToCodeService.GetCurrentGetToCodeDialog();
            ProjectMRUTestExtension projectMru = getToCode.ProjectMRU;
            ProjectMRUItemTestExtension lastAccessedItem = projectMru.LastAccessedItem;
            lastAccessedItem.OpenItem();

            this.WaitForSolutionLoad();
        }

        /// <remarks>
        /// Inherits parent process environment so in-proc data collector can work.
        /// </remarks>
        protected override bool InheritParentProcessEnvironment => true;

        protected override StartWindowStartupOption StartWindowStartupOption => StartWindowStartupOption.Enable;

        private void SearchTemplates(CreateProjectDialogTestExtension createProject, string searchText, bool isFirstSearch = false)
        {
            // allow more time on first search, due the the search index need rebuild.
            createProject.SearchTemplates(searchText, timeOutInSeconds: isFirstSearch ? 60 : 10);
            this.WaitForIdle();
        }

        /// <summary>
        /// Waits for process CPU idle and UI thread idle to ensure UI code is included in optimization profile.
        /// </summary>
        private void WaitForIdle()
        {
            this.VisualStudio.WaitForProcessCpuIdle(TimeSpan.FromSeconds(10));
            this.GetToCodeService.WaitForUIThreadIdle(TimeSpan.FromSeconds(10));
        }

        private void WaitForSolutionLoad()
        {
            this.WaitForIdle();

            this.VisualStudio.ObjectModel.Solution.WaitForFullyLoaded();
            this.VisualStudio.ObjectModel.Solution.Verify.IsFullyLoaded();
        }
    }
}
