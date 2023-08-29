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
using Microsoft.VisualStudio.TemplateProviders.TemplateProviderApi;

namespace TemplateListIsSearchableFromSearchBox
{
    [TestClass]
    public class TemplateListIsSearchableTest : VisualStudioHostTest
    {
        private NewProjectDialogTestService newProjectDialogTestService;

        private List<string> foldersToCleanUp = new List<string>();
        private readonly VsSafeFileNameStringGenerator stringGenerator = new VsSafeFileNameStringGenerator() { MinLength = 3, MaxLength = 10 };

        string[] searchText = { "Blank Solution", "Console", "WPF", "Azure"};
        private static TimeSpan defaultOpenSolutionTimeout = TimeSpan.FromSeconds(2);

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();
            this.newProjectDialogTestService = this.GetAndCheckService<NewProjectDialogTestService>();
        }

        [TestMethod]
        [Owner("IDE Experience")]
        public void TemplateListIsSearchableFromSearchBox()
        {
            foreach (string txt in searchText)
            {
                NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
                CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();

                // Clear filters and search term before count the template
                createProjectDialogTestExtension.ClearFilters();
                createProjectDialogTestExtension.ClearSearch();

                //const string searchText = "Blank Solution";

                int totalExtensions = createProjectDialogTestExtension.GetTemplateShownCount();
                totalExtensions.Should().BeGreaterThan(0);

                // Search "Blank Solution"
                createProjectDialogTestExtension.SearchTemplates(txt);

                createProjectDialogTestExtension.Verify.HasTemplates().Should().BeTrue(because: $"Should find result when search '{searchText}'");

                int totalSearchResult = createProjectDialogTestExtension.GetTaggedTemplateShownCount();

                createProjectDialogTestExtension.ClearFilters();

                newProjectDialogTestExtension.CloseDialog();
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
