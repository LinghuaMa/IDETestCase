using System;

using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.Shell;
using Microsoft.VisualStudio.PlatformUI.Apex.NewProjectDialog;
using Microsoft.VisualStudio.PlatformUI.Apex.StartPage;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VsIde.IdeExp.StartPageTests
{
    [TestClass]
    public class StartPageEnabledUserScenarioTests : VisualStudioHostTest
    {
        private const int FifteenMinutesInMs = 15 * 60 * 1000;
        private const string IdeExperienceEngineersAlias = "IDE Experience";

        private readonly Guid startPageGuid = new Guid("e506b91c-c606-466a-90a9-123d1d1e12b3");

        // This command set should be in sync with Microsoft.VisualStudio.PlatformUI.MruPackageGuids in
        // src\env\shell\UIInternal\Packages\MRU\Package\Guids.cs
        private readonly Guid recentProjectsAndSolutionsCommandSet = new Guid("{001bd6e5-e2cd-4f47-98d4-2d39215359fc}");
        private const uint firstMruItemCmdId = 1;

        private StartPageTestService startPageService;
        private WindowManagementService windowManager;

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            this.windowManager = this.GetAndCheckService<WindowManagementService>();
            this.windowManager.FindToolWindow(startPageGuid, showWindow: true);
            this.startPageService = this.GetAndCheckService<StartPageTestService>();
        }

        [TestCleanup]
        public override void TestCleanup()
        {
            base.TestCleanup();

            this.VisualStudio.Stop();
        }

        [TestMethod]
        [Owner(IdeExperienceEngineersAlias)]
        [Timeout(FifteenMinutesInMs)]
        public void SearchTemplateCreateProjectAndOpenProjectFromMruShouldSucceed()
        {
            StartPageTestExtension startPage = this.startPageService.GetCurrentStartPage();

            CreateProjectDialogTestExtension createProject = startPage.NavigateToNewProjectDialog();

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

            startPage = this.startPageService.GetCurrentStartPage();
            StartPageProjectMRUTestExtension projectMru = startPage.ProjectMRUTestExtension;
            ProjectMRUItemTestExtension lastAccessedItem = projectMru.LastAccessedItem;
            lastAccessedItem.OpenItem();

            this.WaitForSolutionLoad();
        }

        protected override VisualStudioHostConfiguration GetHostConfiguration()
        {
            return new VisualStudioHostConfiguration
            {
                StartWindowStartupOption = StartWindowStartupOption.Disable,
            };
        }

        private T GetAndCheckService<T>() where T : class
        {
            var testService = this.VisualStudio.Get<T>();
            this.Verify.IsNotNull(testService, $"{typeof(T).Name} not found.");
            return testService;
        }

        private void SearchTemplates(CreateProjectDialogTestExtension createProject, string searchText, bool isFirstSearch = false)
        {
            // allow more time on first search, due the the search index need rebuild.
            createProject.SearchTemplates(searchText, timeOutInSeconds: isFirstSearch ? 60 : 10);
            this.VisualStudio.WaitForProcessCpuIdle(TimeSpan.FromSeconds(10));
        }

        private void RestartVisualStudio()
        {
            this.VisualStudio.Stop(); // the base class will automatically restart VS when the VisualStudio property getter is next invoked

            this.windowManager = this.GetAndCheckService<WindowManagementService>();
            this.windowManager.FindToolWindow(startPageGuid, showWindow: true);
            this.startPageService = this.GetAndCheckService<StartPageTestService>();
        }

        private void WaitForSolutionLoad()
        {
            this.VisualStudio.WaitForProcessCpuIdle(TimeSpan.FromSeconds(10));

            this.VisualStudio.ObjectModel.Solution.WaitForFullyLoaded();
            this.VisualStudio.ObjectModel.Solution.Verify.IsFullyLoaded();
        }
    }
}
