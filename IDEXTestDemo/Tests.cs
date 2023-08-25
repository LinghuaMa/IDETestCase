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


namespace Example
{
    [TestClass]
    public class Tests : VisualStudioHostTest
    {
        /// <summary>
        /// Gets the directory where the test is currently executing from.
        /// </summary>
 

        [TestClass]
        public class NPDUserScenarioTests : VisualStudioHostTest
        {
            //public string TestExecutionDirectory { get => Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); }
            private NewProjectDialogTestService newProjectDialogTestService;

            private List<string> foldersToCleanUp = new List<string>();
            private readonly VsSafeFileNameStringGenerator stringGenerator = new VsSafeFileNameStringGenerator() { MinLength = 3, MaxLength = 10 };
            private const string ConsoleAppTemplateId = "Microsoft.CSharp.ConsoleApplication";
            private const string WPFTemplateId = "Microsoft.VisualBasic.Windows.WpfApplication";
            private const string AzureFunctionTemplateId = "Microsoft.AzureFunctions.ProjectTemplate.CSharp";
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
                string projectLocation = Environment.ExpandEnvironmentVariables($@"%USERPROFILE%\Source\Repos\{stringGenerator.Generate()}");
                string solutionName = stringGenerator.Generate();
                string projectName = stringGenerator.Generate();

                NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
                CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();
                createProjectDialogTestExtension.Verify.ProjectCreationDialogVisibilityIs(true).Should().BeTrue();

                string templateId = WPFTemplateId;

                if (!createProjectDialogTestExtension.TrySelectTemplateWithId(templateId))
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

                this.VerifySolutionIsLoaded(Path.Combine(projectLocation, solutionName, solutionName + ".sln"));
            }

            [TestMethod]
            [Owner("IDE Experience")]
            public void BlankSolutionTemplateListIsSearchableFromSearchBox()
            {
                NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
                CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();

                // Clear filters and search term before count the template
                createProjectDialogTestExtension.ClearFilters();
                createProjectDialogTestExtension.ClearSearch();

                const string searchText = "Blank Solution";

                int totalExtensions = createProjectDialogTestExtension.GetTemplateShownCount();
                totalExtensions.Should().BeGreaterThan(0);

                // Search "Blank Solution"
                createProjectDialogTestExtension.SearchTemplates(searchText);

                createProjectDialogTestExtension.Verify.HasTemplates().Should().BeTrue(because: $"Should find result when search '{searchText}'");

                int totalSearchResult = createProjectDialogTestExtension.GetTaggedTemplateShownCount();             

                createProjectDialogTestExtension.ClearFilters();

                newProjectDialogTestExtension.CloseDialog();
            }

            [TestMethod]
            [Owner("IDE Experience")]
            public void ConsoleTemplateListIsSearchableFromSearchBox()
            {
                NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
                CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();

                // Clear filters and search term before count the template
                createProjectDialogTestExtension.ClearFilters();
                createProjectDialogTestExtension.ClearSearch();

                const string searchText = "Console";

                int totalExtensions = createProjectDialogTestExtension.GetTemplateShownCount();
                totalExtensions.Should().BeGreaterThan(0);

                // Search "Console"
                createProjectDialogTestExtension.SearchTemplates(searchText);

                Thread.Sleep(5000);
                createProjectDialogTestExtension.Verify.HasTemplates().Should().BeTrue(because: $"Should find result when search '{searchText}'");

                int totalSearchResult = createProjectDialogTestExtension.GetTaggedTemplateShownCount();

                // With .NET Desktop workload installed, we suppose at least 3 Console templates (C#, VB and CPP) are there
                totalSearchResult.Should().BeGreaterOrEqualTo(2, because: $"At least 2 WPF related templates(C# and VB) should be found. Actually {totalSearchResult}");

                // Apply C# filter
                createProjectDialogTestExtension.ApplyFilters(TemplateTagType.Language, "C#");

                int CSharpSearchResult = createProjectDialogTestExtension.GetTaggedTemplateShownCount();
                CSharpSearchResult.Should().Be(totalSearchResult, because: "Apply filter on search results should not remove any results");

                // Apply VB filter
                createProjectDialogTestExtension.ApplyFilters(TemplateTagType.Language, "Visual Basic");

                int VBSearchResult = createProjectDialogTestExtension.GetTaggedTemplateShownCount();
                VBSearchResult.Should().Be(totalSearchResult, because: "Apply filter on search results should not remove any results");

                // Apply C++ filter
                createProjectDialogTestExtension.ApplyFilters(TemplateTagType.Language, "C++");

                int CPPSearchResult = createProjectDialogTestExtension.GetTaggedTemplateShownCount();
                VBSearchResult.Should().Be(totalSearchResult, because: "Apply filter on search results should not remove any results");
                
                createProjectDialogTestExtension.ClearFilters();

                newProjectDialogTestExtension.CloseDialog();
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
}