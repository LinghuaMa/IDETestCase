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

        string[] searchText = { "Blank Solution", "Console", "WPF", "Azure" };
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
            NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
            CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();

            foreach (string txt in searchText)
            {
                createProjectDialogTestExtension.SearchTemplates(txt);

                createProjectDialogTestExtension.Verify.HasTemplates().Should().BeTrue(because: $"Should find result when search '{searchText}'");

                int totalSearchResult = createProjectDialogTestExtension.GetTaggedTemplateShownCount();

                totalSearchResult.Should().BeGreaterOrEqualTo(2, because: $"At least 2 WPF related templates(C# and VB) should be found. Actually {totalSearchResult}");

                // Clear filters and search term before count the template
                createProjectDialogTestExtension.ClearFilters();
                createProjectDialogTestExtension.ClearSearch();
            }

            newProjectDialogTestExtension.CloseDialog();
            VisualStudio.Stop();
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

