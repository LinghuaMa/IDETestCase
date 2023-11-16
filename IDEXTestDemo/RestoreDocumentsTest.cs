using Microsoft.Test.Apex.Services.StringGeneration;
using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.Solution;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using FluentAssertions;
using Microsoft.Test.Apex.VisualStudio.Shell;
using Microsoft.VisualStudio.PlatformUI.Apex.NewProjectDialog;
using System.IO;

namespace Vside.IdeExp.SignoffTests
{
    [TestClass]
    public class RestoreDocumentsTest : VisualStudioHostTest
    {
        private NewProjectDialogTestService newProjectDialogTestService;

        private List<string> foldersToCleanUp = new List<string>();
        private readonly VsSafeFileNameStringGenerator stringGenerator = new VsSafeFileNameStringGenerator() { MinLength = 3, MaxLength = 10 };

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();
            this.newProjectDialogTestService = this.GetAndCheckService<NewProjectDialogTestService>();
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
        public void CreateBlankSolution()
        {
            string solutionLocation = Environment.ExpandEnvironmentVariables($@"%USERPROFILE%\source\repos\{stringGenerator.Generate()}");
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
                var aaa = configTestExtension.SolutionName;
                // Verify that does not show "Project name" in config page
                configTestExtension.Verify.IsNull(configTestExtension.ProjectName);
                // Assign valid Location
                configTestExtension.Location = solutionLocation;
                configTestExtension.FinishConfiguration();
                Delay(10);
                this.VisualStudio.ObjectModel.Solution.FilePath.Should().Be(Path.Combine(solutionLocation, "Solution1", solutionName + ".sln"));
                //this.VisualStudio.ObjectModel.Solution.SaveAndClose();
                //VisualStudio.Stop();
                //SetPermissionsToNormalOnDotGitFolder(solutionLocation);
                //Directory.Delete(solutionLocation, true);
            }
            using (Scope.Enter("[Project] NPD configuration page behavior"))
            {
                NewProjectDialogTestExtension newProjectDialogTestExtension = this.newProjectDialogTestService.GetExtension();
                CreateProjectDialogTestExtension createProjectDialogTestExtension = newProjectDialogTestExtension.InvokeNewProjectDialog();
                createProjectDialogTestExtension.Verify.ProjectCreationDialogVisibilityIs(true).Should().BeTrue();
                string templateID = "Microsoft.JavaScript.NodejsWebApp";
                if (!createProjectDialogTestExtension.TrySelectTemplateWithId(templateID))
                    return;
                ConfigProjectDialogTestExtension configTestExtension = createProjectDialogTestExtension.NavigateToConfigProjectPage();
                configTestExtension.Verify.ProjectConfigurationDialogVisibilityIs(true).Should().BeTrue();

                // Assign valid SolutionName
                configTestExtension.SolutionName = solutionName;
                var aaa = configTestExtension.SolutionName;
                // Verify that does not show "Project name" in config page
                configTestExtension.Verify.IsNull(configTestExtension.ProjectName);
                // Assign valid Location
                configTestExtension.Location = solutionLocation;
                configTestExtension.FinishConfiguration();
                Delay(10);
                //this.VisualStudio.ObjectModel.Solution.FilePath.Should().Be(Path.Combine(solutionLocation, "Solution1", solutionName + ".sln"));
                //this.VisualStudio.ObjectModel.Solution.SaveAndClose();
                VisualStudio.Stop();
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
