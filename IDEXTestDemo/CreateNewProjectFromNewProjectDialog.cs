using Microsoft.Test.Apex;
using Microsoft.Test.Apex.Services.StringGeneration;
using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.Solution;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Threading;

using FluentAssertions;
using Microsoft.Test.Apex.VisualStudio.Shell;
using Microsoft.VisualStudio.PlatformUI.Apex.NewProjectDialog;
using Microsoft.VisualStudio.TemplateProviders.Templates;

namespace CreateNewProjectFromNPD
{
    [TestClass]
    public class CreateNewProjectFromNPD : VisualStudioHostTest
    {
        private NewProjectDialogTestService newProjectDialogTestService;

        private List<string> foldersToCleanUp = new List<string>();
        private readonly VsSafeFileNameStringGenerator stringGenerator = new VsSafeFileNameStringGenerator() { MinLength = 3, MaxLength = 10 };

        string[] templateId = { "Microsoft.Common.Console``C#", "Microsoft.Common.Console``VB", "Microsoft.Common.WinForms``C#", "Microsoft.Common.WinForms``VB", "Microsoft.Common.WPF``C#", "Microsoft.Web.RazorPages``C#", "AzureFunctions``C#", "Microsoft.CSharp.ConsoleApplication", "Microsoft.VisualBasic.Windows.WpfApplication" };
        private static TimeSpan defaultOpenSolutionTimeout = TimeSpan.FromSeconds(2);

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();
            this.newProjectDialogTestService = this.GetAndCheckService<NewProjectDialogTestService>();
        }

        [TestMethod]
        [Owner("IDE Experience")]
        public void CreateNewProjectFromNewProjectDialog()
        {
            foreach (string id in templateId)
            {
                string projectLocation = Environment.ExpandEnvironmentVariables($@"%USERPROFILE%\source\repos\{stringGenerator.Generate()}");
                string solutionName = stringGenerator.Generate();
                string projectName = stringGenerator.Generate();

                NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
                CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();
                createProjectDialogTestExtension.Verify.ProjectCreationDialogVisibilityIs(true).Should().BeTrue();
                //createProjectDialogTestExtension.SearchTemplates();
                if (!createProjectDialogTestExtension.TrySelectTemplateWithId(id))
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

                if (!id.Contains("Application"))
                    configTestExtension.FinishConfiguration();

                this.VerifySolutionIsLoaded(Path.Combine(projectLocation, solutionName, solutionName + ".sln"));

                this.VisualStudio.ObjectModel.Solution.SaveAndClose();
            }
        }

        [TestMethod]
        [Owner("IDE Experience")]
        public void CreateBlankSolution()
        {
            string projectLocation = Environment.ExpandEnvironmentVariables($@"%USERPROFILE%\source\repos\{stringGenerator.Generate()}");
            string solutionName = stringGenerator.Generate();
            using (Scope.Enter("[Blank Solution] NPD configuration page behavior when \"Create new project\""))
            {
                NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
                CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();
                createProjectDialogTestExtension.Verify.ProjectCreationDialogVisibilityIs(true).Should().BeTrue();
                createProjectDialogTestExtension.SearchTemplates("Blank Solution");
                createProjectDialogTestExtension.GetTemplateShownCount().Should().Be(11);
                ConfigProjectDialogTestExtension configTestExtension = createProjectDialogTestExtension.NavigateToConfigProjectPage();
                configTestExtension.Verify.ProjectConfigurationDialogVisibilityIs(true).Should().BeTrue();

                // Assign valid SolutionName
                configTestExtension.SolutionName = solutionName;
                var aaa= configTestExtension.SolutionName;
                // Verify that does not show "Project name" in config page
                configTestExtension.Verify.IsNull(configTestExtension.ProjectName);
                // Assign valid Location
                configTestExtension.Location = projectLocation;
                configTestExtension.FinishConfiguration();
                this.VerifySolutionIsLoaded(Path.Combine(projectLocation, "Solution1", solutionName + ".sln"));
                this.VisualStudio.ObjectModel.Solution.SaveAndClose();
            }
            using (Scope.Enter("[Project] NPD configuration page behavior"))
            {
                
            }
        }


        [TestCleanup]
        public override void TestCleanup()
        {
            base.TestCleanup();
            this.VisualStudio.Stop();

            foreach (string folder in this.foldersToCleanUp)
            {
                if (Directory.Exists(folder))
                {
                    Directory.Delete(folder, recursive: true);
                }
            }
        }

        protected override VisualStudioHostConfiguration GetHostConfiguration()
        {
            return new HostConfiguration
            {
                StartWindowStartupOption = StartWindowStartupOption.Disable,
            };
        }

        /// <summary>
        /// Custom Visual Studio host configuration to ensure Microsoft.VisualStudio.UI.Internal.Apex is included in
        /// Apex MEF composition
        /// </summary>
        private class HostConfiguration : VisualStudioHostConfiguration
        {
            public override IEnumerable<string> CompositionAssemblies
            {
                get
                {
                    HashSet<string> assemblies = new HashSet<string>(base.CompositionAssemblies);
                    _ = assemblies.Add(typeof(NewProjectDialogTestService).Assembly.Location);
                    return assemblies;
                }
            }
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

        private void VerifySolutionOpened(TimeSpan verifyAfterInSec)
        {
            Thread.Sleep(verifyAfterInSec);
            this.VisualStudio.ObjectModel.Solution.IsOpen.Should().BeTrue(because: $"solution should be opened in {verifyAfterInSec} seconds");
        }

        private void VerifySolutionIsLoaded(string solutionFilePath)
        {
            // Ensure the solution is opened before verify fully loaded
            this.VerifySolutionOpened(defaultOpenSolutionTimeout);

            this.VisualStudio.ObjectModel.Solution.WaitForFullyLoaded(singleProjectLoadTimeout: TimeSpan.FromSeconds(10));
            this.VisualStudio.ObjectModel.Solution.FilePath.Should().Be(solutionFilePath);
        }
    }
}
