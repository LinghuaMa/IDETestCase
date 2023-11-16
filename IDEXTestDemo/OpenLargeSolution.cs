using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

using FluentAssertions;

using Microsoft.Test.Apex.Services.StringGeneration;
using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.FeatureFlags;
using Microsoft.Test.Apex.VisualStudio.FolderView;
using Microsoft.Test.Apex.VisualStudio.Settings;
using Microsoft.Test.Apex.VisualStudio.Shell;
using Microsoft.Test.Apex.VisualStudio.Solution;
using Microsoft.VisualStudio.PlatformUI.Apex.GetToCode;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using VsIde.IdeExp.GetToCodeTests;

using Task = System.Threading.Tasks.Task;

namespace VsIde.IdeExp.GetToCodeTests
{
    [TestClass]
    public class OpenLargeSolution : GetToCodeTestBase
    {
        private readonly TimeSpan dialogShowTimeout = TimeSpan.FromSeconds(5);
        protected override StartWindowStartupOption StartWindowStartupOption => StartWindowStartupOption.UseVisualStudioSetting;
        private readonly List<string> foldersToCleanUp = new List<string>();
        private readonly VsSafeFileNameStringGenerator stringGenerator = new VsSafeFileNameStringGenerator() { MinLength = 3, MaxLength = 20 };

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();
            SettingsService settings = this.GetAndCheckService<SettingsService>();
            if (settings.Environment.General.OnStartup != EnvironmentSettings.StartupOption.GetToCode)
            {
                this.SetStartupOption(EnvironmentSettings.StartupOption.GetToCode);

                FeatureFlagService featureFlag = this.GetAndCheckService<FeatureFlagService>();

                // hide start page for G2C test
                if (featureFlag.GetFeatureFlag(@"VS.Core.StartPage"))
                {
                    featureFlag.SetFeatureFlag(@"VS.Core.StartPage", false);
                }

                this.RestartVisualStudio(); // resets this.GetToCodeService
            }
        }

        [TestCleanup]
        public override void TestCleanup()
        {
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

        private const uint cmdId = 219;
        private Guid solutionCommandGroupGuid = new Guid("5efc7975-14bc-11cf-9b2b-00aa00573819");

        //private void CloseSolution()
        //{
        //    // Do close solution in a different thread due to Start Window is a modal window and will block the process
        //    Task task = Task.Run(() =>
        //    {
        //        CommandingService cmdService = this.GetAndCheckService<CommandingService>();
        //        cmdService.ExecuteCommand(solutionCommandGroupGuid, cmdId, null);
        //    });

        //    task.GetAwaiter().OnCompleted(() =>
        //    {
        //        if (task.IsFaulted)
        //        {
        //            Assert.Fail(task.Exception.Message);
        //        }
        //    });
        //}

        private void SetStartupOption(EnvironmentSettings.StartupOption option)
        {
            SettingsService settings = this.GetAndCheckService<SettingsService>();
            settings.Environment.General.OnStartup = option;
        }
    }
}
