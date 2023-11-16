using Microsoft.Test.Apex.Services.StringGeneration;
using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.Solution;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using FluentAssertions;
using Microsoft.Test.Apex.VisualStudio.Shell;
using Microsoft.VisualStudio.PlatformUI.Apex.NewProjectDialog;

namespace VsIde.IdeExp.GetToCodeTests
{
    [TestClass]
    public class NewProjectTests : VisualStudioHostTest
    {
        private NewProjectDialogTestService newProjectDialogTestService;
        private readonly VsSafeFileNameStringGenerator stringGenerator = new VsSafeFileNameStringGenerator() { MinLength = 3, MaxLength = 10 };

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            this.newProjectDialogTestService = this.GetAndCheckService<NewProjectDialogTestService>();
        }

        [TestMethod, Owner("IDE Experience"), Timeout(5 * 60 * 1000)]
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
