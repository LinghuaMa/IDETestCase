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

using Task = System.Threading.Tasks.Task;

namespace VsIde.IdeExp.GetToCodeTests
{
    /// <summary>
    /// VS IDE Apex tests for GetToCode
    ///
    /// Note: Open folder tests are not tested here. They are being manually tested by CTI.
    /// </summary>
    [TestClass]
    public class UserScenarioTests : GetToCodeTestBase
    {
        private readonly List<string> foldersToCleanUp = new List<string>();

        private readonly VsSafeFileNameStringGenerator stringGenerator = new VsSafeFileNameStringGenerator() { MinLength = 3, MaxLength = 20 };
        private readonly TimeSpan dialogShowTimeout = TimeSpan.FromSeconds(5);

        // This command set should be in sync with Microsoft.VisualStudio.PlatformUI.MruPackageGuids in
        // src\env\shell\UIInternal\Packages\MRU\Package\Guids.cs
        private readonly Guid recentProjectsAndSolutionsCommandSet = new Guid("{001bd6e5-e2cd-4f47-98d4-2d39215359fc}");

        internal const int TwoMinutesInMs = 1000 * 60 * 2;
        internal const int ThreeMinutesInMs = 1000 * 60 * 3;

        private const uint firstMruItemCmdId = 1;
        private const uint cmdId = 219;
        private Guid solutionCommandGroupGuid = new Guid("5efc7975-14bc-11cf-9b2b-00aa00573819");

        internal const string RepoUrl = "https://github.com/octocat/spoon-knife.git";

        protected override StartWindowStartupOption StartWindowStartupOption => StartWindowStartupOption.UseVisualStudioSetting;

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            var settings = this.GetAndCheckService<SettingsService>();
            if (settings.Environment.General.OnStartup != EnvironmentSettings.StartupOption.GetToCode)
            {
                this.SetStartupOption(EnvironmentSettings.StartupOption.GetToCode);

                var featureFlag = this.GetAndCheckService<FeatureFlagService>();

                // hide start page for G2C test
                if (featureFlag.GetFeatureFlag(@"VS.Core.WelcomePage"))
                {
                    featureFlag.SetFeatureFlag(@"VS.Core.WelcomePage", false);
                }

                this.RestartVisualStudio(); // resets this.GetToCodeService
            }
        }

        [TestCleanup]
        public override void TestCleanup()
        {
            this.Delay(3);
            base.TestCleanup();

            foreach (string folder in this.foldersToCleanUp)
            {
                int deleteFolderRetryCount = 0;
                int maxRetryCount = 3;
                bool succeeded = false;
                while (!succeeded)
                {
                    try
                    {
                        Directory.Delete(folder, recursive: true);
                        succeeded = true;
                    }
                    catch (IOException) when (deleteFolderRetryCount++ < maxRetryCount)
                    {
                        Thread.Sleep(500);
                    }
                }
            }
        }

        [TestMethod, Owner("IDE Experience"), Timeout(ThreeMinutesInMs)]
        public void CheckProjectMRU()
        {
            string solutionFullPath = string.Empty;

            using (this.Scope.Enter("Create a test solution"))
            {
                solutionFullPath = this.CreateTestSolution();

                if (string.IsNullOrEmpty(solutionFullPath))
                {
                    Assert.Fail("Failed to create a test solution.");
                    return;
                }
            }

            using (this.Scope.Enter("Open solution and then close VS"))
            {
                this.SetStartupOption(EnvironmentSettings.StartupOption.ShowEmptyEnvironment);
                this.RestartVisualStudio();

                this.VisualStudio.ObjectModel.Solution.Open(solutionFullPath);
                this.VerifySolutionIsOpen(solutionFullPath);
                this.SetStartupOption(EnvironmentSettings.StartupOption.GetToCode);

                this.CloseSolutionAndGetToCode();
                this.RestartVisualStudio();
            }

            using (this.Scope.Enter("Verify the last access solution"))
            {
                GetToCodeDialogTestExtension dialog = this.GetToCodeService.GetCurrentGetToCodeDialog();
                ProjectMRUTestExtension projectMRUTestExtension = dialog.ProjectMRU;
                ProjectMRUItemTestExtension projectMRUItemTestExtension = projectMRUTestExtension.LastAccessedItem;

                Assert.IsTrue(projectMRUItemTestExtension.Verify.ItemPathEquals(Path.GetDirectoryName(solutionFullPath)));
                Assert.IsTrue(projectMRUItemTestExtension.Verify.ItemNameEquals(Path.GetFileName(solutionFullPath)));
                Assert.IsTrue(projectMRUItemTestExtension.Verify.ItemIsPinnedEquals(false));

                projectMRUTestExtension.TogglePinStatusForLastAccessedItem();
                projectMRUItemTestExtension = projectMRUTestExtension.LastAccessedItem;
                Assert.IsTrue(projectMRUItemTestExtension.Verify.ItemIsPinnedEquals(true));

                projectMRUTestExtension.TogglePinStatusForLastAccessedItem();
                projectMRUItemTestExtension = projectMRUTestExtension.LastAccessedItem;
                Assert.IsTrue(projectMRUItemTestExtension.Verify.ItemIsPinnedEquals(false));

                projectMRUItemTestExtension.OpenItem();
                this.VerifySolutionIsOpen(solutionFullPath);

                this.CloseSolutionAndGetToCode();
                projectMRUItemTestExtension.RemoveItemFromMRUList();
            }
        }

        [TestMethod, Owner("IDE Experience"), Timeout(ThreeMinutesInMs)]
        // This test guards against regressing this feedback ticket: https://devdiv.visualstudio.com/DevDiv/_workitems/edit/593358
        public void CheckMRUCommand()
        {
            string solutionFullPath = string.Empty;

            using (this.Scope.Enter("Create a test solution"))
            {
                solutionFullPath = this.CreateTestSolution();

                if (string.IsNullOrEmpty(solutionFullPath))
                {
                    Assert.Fail("Failed to create a test solution.");
                    return;
                }
            }

            using (this.Scope.Enter("Open solution and then close VS"))
            {
                this.VisualStudio.ObjectModel.Solution.Open(solutionFullPath);
                this.VerifySolutionIsOpen(solutionFullPath);

                this.SetStartupOption(EnvironmentSettings.StartupOption.ShowEmptyEnvironment);
                this.RestartVisualStudio();
            }

            using (this.Scope.Enter("Check if the top MRU item in the File > Recent Projects and Solutions menu is visible."))
            {
                CommandQueryResult result = this.VisualStudio.ObjectModel.Commanding.QueryStatusCommand(recentProjectsAndSolutionsCommandSet, firstMruItemCmdId);
                Assert.IsTrue(result.IsVisible, "Top MRU item was expected to be found in the File > Recent Projects and solutions menu.");
            }
        }

        [TestMethod, Owner("IDE Experience"), Timeout(ThreeMinutesInMs)]
        public void CheckSearchMRU()
        {
            string solutionFullPath = string.Empty;
            GetToCodeDialogTestExtension dialog;
            ProjectMRUTestExtension projectMRUTestExtension;

            using (this.Scope.Enter("Create new test solution"))
            {
                solutionFullPath = this.CreateTestSolution();

                if (string.IsNullOrEmpty(solutionFullPath))
                {
                    Assert.Fail("Failed to create a test solution.");
                    return;
                }
            }

            using (this.Scope.Enter("Open solution and then close VS"))
            {
                this.SetStartupOption(EnvironmentSettings.StartupOption.ShowEmptyEnvironment);
                this.RestartVisualStudio();

                this.VisualStudio.ObjectModel.Solution.Open(solutionFullPath);
                this.VerifySolutionIsOpen(solutionFullPath);
                this.SetStartupOption(EnvironmentSettings.StartupOption.GetToCode);

                this.CloseSolutionAndGetToCode();
                this.RestartVisualStudio();
            }

            using (this.Scope.Enter("Verify the last access solution"))
            {
                dialog = this.GetToCodeService.GetCurrentGetToCodeDialog();
                projectMRUTestExtension = dialog.ProjectMRU;

                // start 1 search
                Assert.IsTrue(projectMRUTestExtension.SearchMRUItems(Path.GetFileNameWithoutExtension(solutionFullPath), timeOutInSeconds: 15), "Search exceeded timeout.");
                Assert.IsTrue(projectMRUTestExtension.MRUListContains(solutionFullPath), "MRU list doesn't contain newly added solution.");
                Assert.IsTrue(projectMRUTestExtension.SearchResultsContain(solutionFullPath), "Solution is not in the search results.");
                Assert.IsTrue(projectMRUTestExtension.TopSearchResultEquals(solutionFullPath), "Solution is not the first of the search results.");

                this.CloseSolutionAndGetToCode();
                projectMRUTestExtension.RemoveAllItemsFromMRU();
            }
        }

        [TestMethod, Owner("IDE Experience"), Timeout(FiveMinutesInMs)]
        public void ClonePublicGitRepo()
        {
            // Spoon-Knife is a sample repo on GitHub that's only there for demonstration purposes
            const string ValidCloneUrl = "https://github.com/octocat/Spoon-Knife";
            const string InvalidCloneUrl = "   "; // We only validate empty string, null and whitespace
            string cloneLocalPath = Environment.ExpandEnvironmentVariables($@"%USERPROFILE%\Source\Repos\{stringGenerator.Generate()}");

            GetToCodeDialogTestExtension g2cDialogTestExtension = this.GetToCodeService.GetCurrentGetToCodeDialog();
            CloneDialogTestExtension cloneDialogTestExtension = g2cDialogTestExtension.NavigateToCloneDialog();

            cloneDialogTestExtension.Verify.CloneDialogIsOpenEquals(true);
            cloneDialogTestExtension.Verify.CanCloneEquals(false);

            cloneDialogTestExtension.SetCloneUrlTo(InvalidCloneUrl);
            cloneDialogTestExtension.SetCloneLocalPathTo(cloneLocalPath);

            cloneDialogTestExtension.Verify.CloneUrlEquals(InvalidCloneUrl);
            cloneDialogTestExtension.Verify.CloneLocalPathEquals(cloneLocalPath);
            cloneDialogTestExtension.Verify.CanCloneEquals(false);

            cloneDialogTestExtension.SetCloneUrlTo(ValidCloneUrl);
            cloneDialogTestExtension.SetCloneLocalPathTo(cloneLocalPath);

            cloneDialogTestExtension.Verify.CloneUrlEquals(ValidCloneUrl);
            cloneDialogTestExtension.Verify.CloneLocalPathEquals(cloneLocalPath);
            cloneDialogTestExtension.Verify.CanCloneEquals(true);

            cloneDialogTestExtension.BeginClone();

            cloneDialogTestExtension.WaitForCloneToComplete();

            this.GetToCodeService.Verify.GetToCodeDialogIsClosedWithinTimeSpan(this.dialogShowTimeout);

            this.Services.Synchronization.WaitFor(TimeSpan.FromSeconds(10), () =>
            {
                return this.VisualStudio.ObjectModel.Solution.IsOpen;
            });

            Assert.IsTrue(this.VisualStudio.ObjectModel.Solution.IsOpen);

            this.CleanupRepo(cloneLocalPath);
        }

        [TestMethod, Owner("IDE Experience"), Timeout(TwoMinutesInMs)]
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

        [TestMethod, Owner("IDE Experience"), Timeout(TwoMinutesInMs)]
        public void CloneFromCodeContainerProvider()
        {
            GetToCodeDialogTestExtension g2cDialogTestExtension = this.GetToCodeService.GetCurrentGetToCodeDialog();
            CloneDialogTestExtension cloneDialogTestExtension = g2cDialogTestExtension.NavigateToCloneDialog();

            Assert.IsTrue(cloneDialogTestExtension.Verify.CloneDialogIsOpenEquals(true));

            // Add a new provider
            string acquiredFolderName = Environment.ExpandEnvironmentVariables($@"%USERPROFILE%\Source\Repos\{stringGenerator.Generate()}");
            string providerName = stringGenerator.Generate();
            cloneDialogTestExtension.AddTestCodeContainerProvider(providerName, acquiredFolderName);

            IEnumerable<CodeContainerProviderTestExtension> codeContainerProviders = cloneDialogTestExtension.GetInstalledCodeContainerProviders();

            // The first provider should always be the Microsoft owned Azure Repos provider
            CodeContainerProviderTestExtension azureRepos = codeContainerProviders.FirstOrDefault();

            Assert.IsNotNull(azureRepos, "Azure repos was expected to be installed");
            azureRepos.Verify.NameIs("Azure DevOps");

            CodeContainerProviderTestExtension testProvider = codeContainerProviders.FirstOrDefault(ccp => ccp.Name.Equals(providerName, StringComparison.CurrentCulture));
            Assert.IsNotNull(testProvider);

            testProvider.Invoke();
            testProvider.WaitForCodeContainerAcquisitionComplete();

            this.Services.Synchronization.WaitFor(TimeSpan.FromSeconds(10), () =>
            {
                return this.VisualStudio.ObjectModel.Solution.IsOpen;
            });

            Assert.IsTrue(this.VisualStudio.ObjectModel.Solution.IsOpen);
            this.VisualStudio.ObjectModel.MainWindow.Verify.IsShown();

            this.CleanupRepo(acquiredFolderName);
        }

        [TestMethod, Owner("IDE Experience"), Timeout(TwoMinutesInMs)]
        public void GetToCodeDialogShouldNotAppearIfDifferentStartupOptionSelected()
        {
            this.CloseGetToCodeDialog();
            this.SetStartupOption(EnvironmentSettings.StartupOption.ShowEmptyEnvironment);
            this.RestartVisualStudio();
            Assert.IsTrue(this.GetToCodeService.Verify.GetToCodeDialogIsNotVisible());
            this.VisualStudio.ObjectModel.MainWindow.Verify.IsShown();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(TwoMinutesInMs)]
        [DeploymentItem(@"Assets", "Assets")]
        public void LaunchingVsWithSolutionOnCommandLineShouldBypassGetToCodeDialog()
        {
            string solutionFilePath = this.GetVanillaProject();

            // Add quotes to prevent whitespace in solutionFilePath from break the command
            this.VisualStudioCommandLine = $" \"{solutionFilePath}\"";
            this.RestartVisualStudio();
            this.VisualStudio.ObjectModel.MainWindow.Verify.IsShown();
            Assert.IsTrue(this.GetToCodeService.Verify.GetToCodeDialogIsNotVisible());
            this.VerifySolutionIsOpen(solutionFilePath);
            this.CloseSolutionAndGetToCode();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(TwoMinutesInMs)]
        public void GetToCodeDialogShouldHaveClickableButtons()
        {
            GetToCodeDialogTestExtension dialog = this.GetToCodeService.GetCurrentGetToCodeDialog();
            Assert.IsTrue(dialog.Verify.AllActionsAreEnabled());
            this.CloseGetToCodeDialog();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(TwoMinutesInMs)]
        public void OpenSolutionFromOpenLocalProjectDialog()
        {
            this.RestartVisualStudio();

            GetToCodeDialogTestExtension startWindow = this.GetToCodeService.GetCurrentGetToCodeDialog();
            OpenProjectDialogTestExtension openProjectDialog = startWindow.NavigateToOpenProjectDialog();

            Assert.IsTrue(openProjectDialog.Verify.IsOpenProjectSolutionDialog(expected: true));
            openProjectDialog.CloseDialog();

            Assert.IsTrue(this.GetToCodeService.Verify.GetToCodeDialogIsVisible());

            this.CloseGetToCodeDialog();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(TwoMinutesInMs)]
        [DeploymentItem(@"Assets", "Assets")]
        public void GetToCodeShowsOnCloseSolution()
        {
            this.RestartVisualStudio();

            string solutionFilePath = this.GetVanillaProject();

            // open the GetToCodeDialog then open a solution.
            GetToCodeDialogTestExtension dialog = this.GetToCodeService.GetCurrentGetToCodeDialog();
            this.OpenSolutionFromGetToCode(dialog, solutionFilePath);

            // Verify the G2C dialog is no longer visible and the solution was
            // opened successfully
            Assert.IsTrue(this.GetToCodeService.Verify.GetToCodeDialogIsNotVisible());
            this.VerifySolutionIsOpen(solutionFilePath);
            this.VisualStudio.ObjectModel.MainWindow.Verify.IsShown();

            // close the solution and verify that GetToCode opens.
            this.CloseSolution();

            Assert.IsTrue(this.GetToCodeService.Verify.GetToCodeDialogIsVisibleWithinTimeSpan(this.dialogShowTimeout));
            GetToCodeDialogTestExtension startWindow = this.GetToCodeService.GetCurrentGetToCodeDialog();
            startWindow.CloseDialog();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(TwoMinutesInMs)]
        [DeploymentItem(@"Assets", "Assets")]
        public void GetToCodeShowsOnCloseSolutionWhenVSOpenedViaCLI()
        {
            string solutionFilePath = this.GetVanillaProject();

            // Add quotes to prevent whitespace in solutionFilePath from break the command
            this.VisualStudioCommandLine = $" \"{solutionFilePath}\"";
            this.RestartVisualStudio();
            this.VisualStudio.ObjectModel.MainWindow.Verify.IsShown();
            Assert.IsTrue(this.GetToCodeService.Verify.GetToCodeDialogIsNotVisible());
            this.VerifySolutionIsOpen(solutionFilePath);

            this.CloseSolution();

            Assert.IsTrue(this.GetToCodeService.Verify.GetToCodeDialogIsVisibleWithinTimeSpan(this.dialogShowTimeout));
            GetToCodeDialogTestExtension startWindow = this.GetToCodeService.GetCurrentGetToCodeDialog();
            startWindow.CloseDialog();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(TwoMinutesInMs)]
        public void SwitchingSolutionsWithGetToCodeOptionSetDoesNotShowGetToCode()
        {
            string solutionFilePath = CreateTestSolution();
            string solutionFilePath2 = CreateTestSolution();
            this.RestartVisualStudio();

            // open the GetToCodeDialog then open first solution.
            GetToCodeDialogTestExtension dialog = this.GetToCodeService.GetCurrentGetToCodeDialog();
            this.OpenSolutionFromGetToCode(dialog, solutionFilePath);

            // Verify the G2C dialog is no longer visible and the solution was
            // opened successfully
            Assert.IsTrue(this.GetToCodeService.Verify.GetToCodeDialogIsNotVisible());
            this.VerifySolutionIsOpen(solutionFilePath);
            this.VisualStudio.ObjectModel.MainWindow.Verify.IsShown();

            // switch to the next solution and verify G2C doesn't open up.
            this.OpenSolution(solutionFilePath2);
            this.VerifySolutionIsOpen(solutionFilePath2);
            Assert.IsTrue(this.GetToCodeService.Verify.GetToCodeDialogIsNotVisible());
            this.CloseSolutionAndGetToCode();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(TwoMinutesInMs)]
        public void LaunchNewProjectDialog()
        {
            GetToCodeDialogTestExtension testExtension = this.GetToCodeService.GetCurrentGetToCodeDialog();
            var npdTestExtension = testExtension.NavigateToNewProjectDialog();
            npdTestExtension.Verify.IsNewProjectDialogDisplayed();
        }

        [TestMethod, Owner("IDE Experience")]
        public void OpenLargeSolutionCase()
        {
            const string ValidCloneUrl = "https://github.com/dotnet/project-system";

            GetToCodeDialogTestExtension g2cDialogTestExtension = this.GetToCodeService.GetCurrentGetToCodeDialog();
            CloneDialogTestExtension cloneDialogTestExtension = g2cDialogTestExtension.NavigateToCloneDialog();

            cloneDialogTestExtension.SetCloneUrlTo(ValidCloneUrl);
            cloneDialogTestExtension.Verify.CloneUrlEquals(ValidCloneUrl);
            cloneDialogTestExtension.Verify.CanCloneEquals(true);
            cloneDialogTestExtension.BeginClone();
            cloneDialogTestExtension.WaitForCloneToComplete();
            this.GetToCodeService.Verify.GetToCodeDialogIsClosedWithinTimeSpan(this.dialogShowTimeout);
            this.Services.Synchronization.WaitFor(TimeSpan.FromSeconds(10), () =>
            {
                return this.VisualStudio.ObjectModel.Solution.IsOpen;
            });

            Assert.IsTrue(this.VisualStudio.ObjectModel.Solution.IsOpen);
            string defaultLocalPath = cloneDialogTestExtension.CloneLocalPath;
            string solutionPath = Path.Combine(defaultLocalPath, @"ProjectSystem.sln");

            this.SetStartupOption(EnvironmentSettings.StartupOption.ShowEmptyEnvironment);
            this.RestartVisualStudio();
            SolutionService solutionService = this.VisualStudio.ObjectModel.Solution;
            solutionService.WaitForFullyLoadedOnOpen = true;
            solutionService.WaitForCpsProjectsFullyLoaded = true;
            solutionService.Open(solutionPath);
            solutionService.WaitForFullyLoaded();
            this.Services.Synchronization.WaitFor(TimeSpan.FromSeconds(60), () =>
            {
                return this.VisualStudio.ObjectModel.Solution.IsOpen;
            });

            solutionService.Verify.LoadedProjectCountEquals(16);
            solutionService.Verify.ProjectCountEquals(16);
            this.SetStartupOption(EnvironmentSettings.StartupOption.ShowEmptyEnvironment);
            this.RestartVisualStudio();
            this.CleanupRepo(defaultLocalPath);
            this.Delay(3);
        }

        [TestMethod, Owner("IDE Experience")]
        public void FilteringToChangedFiles()
        {
            const string ValidCloneUrl = "https://github.com/Microsoft/TailwindTraders-Desktop";

            GetToCodeDialogTestExtension g2cDialogTestExtension = this.GetToCodeService.GetCurrentGetToCodeDialog();
            CloneDialogTestExtension cloneDialogTestExtension = g2cDialogTestExtension.NavigateToCloneDialog();

            cloneDialogTestExtension.SetCloneUrlTo(ValidCloneUrl);
            cloneDialogTestExtension.Verify.CloneUrlEquals(ValidCloneUrl);
            cloneDialogTestExtension.Verify.CanCloneEquals(true);
            cloneDialogTestExtension.BeginClone();
            cloneDialogTestExtension.WaitForCloneToComplete();
            this.GetToCodeService.Verify.GetToCodeDialogIsClosedWithinTimeSpan(this.dialogShowTimeout);
            this.Services.Synchronization.WaitFor(TimeSpan.FromSeconds(10), () =>
            {
                return this.VisualStudio.ObjectModel.Solution.IsOpen;
            });

            Assert.IsTrue(this.VisualStudio.ObjectModel.Solution.IsOpen);
            string defaultLocalPath = cloneDialogTestExtension.CloneLocalPath;
            string solutionPath = Path.Combine(defaultLocalPath, @"Source\CouponReader.NETFramework.sln");

            this.SetStartupOption(EnvironmentSettings.StartupOption.ShowEmptyEnvironment);
            this.RestartVisualStudio();
            SolutionService solutionService = this.VisualStudio.ObjectModel.Solution;
            solutionService.WaitForFullyLoadedOnOpen = true;
            solutionService.WaitForCpsProjectsFullyLoaded = true;
            solutionService.Open(solutionPath);
            solutionService.WaitForFullyLoaded();
            this.Services.Synchronization.WaitFor(TimeSpan.FromSeconds(60), () =>
            {
                return this.VisualStudio.ObjectModel.Solution.IsOpen;
            });
            string projectPath = Path.Combine(defaultLocalPath, @"Source\\CouponReader.Tests\\CouponReader.Tests.csproj");
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
            this.RestartVisualStudio();
            this.CleanupRepo(defaultLocalPath);
            this.Delay(3);
        }

        private void SetStartupOption(EnvironmentSettings.StartupOption option)
        {
            var settings = this.GetAndCheckService<SettingsService>();
            settings.Environment.General.OnStartup = option;
        }

        private void CloseGetToCodeDialog()
        {
            if (GetToCodeService.Verify.GetToCodeDialogIsVisible())
            {
                GetToCodeService.GetCurrentGetToCodeDialog().CloseDialog();
            }
        }

        private string GetVanillaProject()
        {
            string path = Path.Combine(this.TestExecutionDirectory.FullName, @"Assets\HelloWorldConsole\HelloWorldConsole.sln");
            Assert.IsTrue(Verify.IO.FileExists(path));

            return path;
        }

        /// <summary>
        /// Creates a solution and returns the full path of the .sln file.
        /// </summary>
        /// <returns></returns>
        private string CreateTestSolution(string solutionName = null)
        {
            var solution = this.GetAndCheckService<SolutionService>();

            if (string.IsNullOrEmpty(solutionName))
            {
                solution.CreateDefaultProject();
            }
            else
            {
                solution.CreateEmptySolution(solutionName);
            }

            solution.Save();
            this.foldersToCleanUp.Add(Path.GetDirectoryName(solution.FilePath));
            return solution.FilePath;
        }

        private void VerifySolutionIsOpen(string solutionFilePath)
        {
            var solution = this.GetAndCheckService<SolutionService>();
            Assert.IsTrue(solution.Verify.IsOpen());
            Assert.IsTrue(this.Verify.Strings.AreEqual(solutionFilePath, solution.FilePath), $"Solution path is '{solution.FilePath}' ('{solutionFilePath}' expected)");
        }

        private void OpenSolution(string solutionFilePath)
        {
            var solution = this.GetAndCheckService<SolutionService>();
            solution.Open(solutionFilePath);
        }

        private void CloseSolution()
        {
            // Do close solution in a different thread due to Start Window is a modal window and will block the process
            Task task = Task.Run(() =>
            {
                CommandingService cmdService = this.GetAndCheckService<CommandingService>();
                cmdService.ExecuteCommand(solutionCommandGroupGuid, cmdId, null);
            });

            task.GetAwaiter().OnCompleted(() =>
            {
                if (task.IsFaulted)
                {
                    Assert.Fail(task.Exception.Message);
                }
            });
        }

        private void VerifyFolderViewIsOpen(string folder)
        {
            var folderView = this.GetAndCheckService<FolderViewService>();
            Assert.IsTrue(folderView.Verify.IsInFolderView(), "Solution Explorer is not in folder view");
            this.VerifySolutionIsOpen(folder); // TODO: does this check work in folder mode?
        }

        private void CleanupRepo(string cloneLocalPath)
        {
            while (Directory.Exists(cloneLocalPath))
            {
                try
                {
                    Directory.Delete(cloneLocalPath, true);
                }
                catch (UnauthorizedAccessException)
                {
                    this.SetPermissionsToNormalOnDotGitFolder(cloneLocalPath);
                }
                catch (IOException)
                {
                }
            }
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

        // This is a hack to FAKE the idea of opening an existing solution
        // from the GetToCode.
        //
        // NOTE: The OpenProjectDialog isn't GetToCode Code, so it will be
        // tested via unit testing and we'll assume it just works.
        // Maybe later, we will write a TestContract wrapping it so that
        // we can use the Apex Framework to open the OpenProjectDialog from
        // the GetToCode Dialog, and then open a project, but for now,
        // we'll just open the solution directly, and then close GetToCode.
        private void OpenSolutionFromGetToCode(GetToCodeDialogTestExtension dialog, string solutionFilePath)
        {
            OpenSolution(solutionFilePath);
            dialog.CloseDialog();
        }

        /// <summary>
        /// Close an open solution, wait for GetToCode to initialize, and then
        /// close GetToCode.
        ///
        /// NOTE: If a solution is open when VS is closed, the first thing that
        /// happens is CloseSolution() is called.  If the shutdown behaviour
        /// of VS doesn't happen fast enough, it's possible for GetToCode to be
        /// popped open.  This doesn't happen in normal use cases (outside of
        /// Apex), but Apex is a different scenario, and can cause this issue.
        /// Thus, call this method when you might be closing a solution either
        /// directly or indirectly.  It will make sure GetToCode does block VS.
        /// </summary>
        private void CloseSolutionAndGetToCode()
        {
            // Manually close the solution currently open.
            this.CloseSolution();

            // Wait for get to code to be initialized, otherwise we might call
            // this.GetToCodeService.CloseGetToCode() too soon.
            this.GetToCodeService.WaitForGetToCodeToBeOpenWithinTimeSpan(dialogShowTimeout);
            // Close get to code so that we can interact with VS again.  While
            // get to code is open (in the reentrant case), VS is disabled due
            // to get to code being a modal window.
            this.GetToCodeService.CloseGetToCode();
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
