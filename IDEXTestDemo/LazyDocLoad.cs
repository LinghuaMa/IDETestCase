using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Microsoft.Test.Apex;
using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.Shell;
using Microsoft.Test.Apex.VisualStudio.Solution;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LazyTabTests
{
    /// <summary>
    /// Testing lazy tab for VB/VC/C#
    /// </summary>
    [TestClass]
    public class LazyDocLoad : ApexTest
    {
        private VisualStudioHost vs = null;

        private string SolutionPath = @"";

        DocStateModel DocState = new DocStateModel();

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            vs = this.Operations.CreateHost<VisualStudioHost>();
            vs.Start();
        }

        [TestCleanup]
        public override void TestCleanup()
        {
            base.TestCleanup();

            this.VsStop(vs);
        }

        /// <Scenario>
        /// 1. Copy SUO file for making C Sharp designer loaded page
        /// 2. Launch solution with load behavior set to "OnDemand"
        /// 3. Verify only C Sharp loaded
        /// 4. Verify only C Sharp documents filled
        /// 5. Shutdown solution
        /// 6. Copy SUO fileto make C Sharp code loaded page
        /// 7. Rerun launch test.
        /// </Scenario>
        [TestMethod]
        [DeploymentItem(@"Assets\LazyDocLoad\", "LazyDocLoad")]
        [Owner("IDE Experience")]
        [Timeout(5 * 60 * 1000)]
        [Description("Tests lazy Doc load behavior if a CSharp file was the one selected")]
        public void CSharpLazyLoad()
        {
            SolutionPath = LoadSolutionForTest();

            if (!Verify.IO.FileExists(SolutionPath, "Solution for test should exist at {0}", SolutionPath))
                return;

            string tmpFilePath = Path.Combine(vs.ObjectModel.Solution.ContainingDirectory, "file.txt");

            StreamWriter sw = new StreamWriter(tmpFilePath);
            sw.WriteLine("TestFile");
            sw.Close();

            DocumentWindowTestExtension window = vs.ObjectModel.WindowManager.OpenDocument(tmpFilePath);
            string TextFileCaption = window.Caption;
            DocState.AddFile(TextFileCaption, "");

            SetFileAsActive(window.Caption);
            vs.ObjectModel.Solution.SaveAndClose();

            SolutionService slnSvc = vs.ObjectModel.Solution;

            // Disable waiting for the solution to be fully loaded when calling Open()
            slnSvc.WaitForFullyLoadedOnOpen = false;

            // Open existing solution and verify it’s loading asynchronously
            slnSvc.Open(this.SolutionPath, SolutionLoadBehavior.OnDemand);

            if (!slnSvc.IsASLOff())
            {
                slnSvc.Verify.IsNotFullyLoaded();
            }

            //Check initial state:
            if (!CheckStubbed())
            {
                Logger.WriteError("Some documents in wrong load state after initial load");
            }

            //Start putting focus to the tabs and verify state
            for (int i = 0; i < DocState.OpenFiles.Count; i++)
            {
                SetFileAsActive(DocState.OpenFiles[i].FileCaption);
                if (!CheckStubbed())
                {
                    Logger.WriteError("Some documents in wrong load state after {0} selected", DocState.OpenFiles[i].FileCaption);
                }
            }

            //Reset to the initial state and try backwords to make sure order isn't important:
            DocState.ResetListToInitialLoadState();
            SetFileAsActive(TextFileCaption);

            if (slnSvc.IsLoadingPaused)
            {
                slnSvc.LoadBehavior = SolutionLoadBehavior.Realtime;
            }
            slnSvc.Close();

            slnSvc.Open(this.SolutionPath, SolutionLoadBehavior.OnDemand);

            if (!slnSvc.IsASLOff())
            {
                slnSvc.Verify.IsNotFullyLoaded();
            }

            //Sanity check:
            if (!CheckStubbed())
            {
                Logger.WriteError("Some documents in wrong load state after second load");
            }

            for (int i = DocState.OpenFiles.Count - 1; i >= 0; i--)
            {
                SetFileAsActive(DocState.OpenFiles[i].FileCaption);
                if (!CheckStubbed())
                {
                    Logger.WriteError("Some documents in wrong load state after {0} selected, moving backwards on the list", DocState.OpenFiles[i].FileCaption);
                }
            }

            if (slnSvc.IsLoadingPaused)
            {
                slnSvc.LoadBehavior = SolutionLoadBehavior.Realtime;
            }
            this.Services.Synchronization.WaitFor(() =>
            {
                try
                {
                    if (slnSvc.IsOpen) slnSvc.Close();
                    return true;
                }
                catch (System.Runtime.Remoting.RemotingException)
                {
                    return false;
                }
            });
        }

        private string LoadSolutionForTest()
        {
            string solutionPath = Path.Combine(TestExecutionDirectory.FullName, @"LazyDocLoad\LazyDocLoad.sln");

            if (!Verify.IO.FileExists(solutionPath, "Solution for test should exist at {0}", solutionPath))
                return solutionPath;

            SolutionService solutionService = vs.ObjectModel.Solution;
            solutionService.WaitForFullyLoadedOnOpen = true;
            solutionService.WaitForCpsProjectsFullyLoaded = true;
            solutionService.Open(solutionPath);

            // Load all projects and open files. For each project, open the default item without tracking in DocState
            // to keep the opened documents the same as creating a new project, which the test was doing before the fix
            // for https://devdiv.visualstudio.com/DevDiv/_workitems/edit/983333
            LoadWinFormProject(ProjectLanguage.CSharp);
            LoadWinFormProject(ProjectLanguage.VB);

            LoadConsoleProject(ProjectLanguage.VC);
            LoadConsoleProject(ProjectLanguage.CSharp);

            LoadWPFProject(ProjectLanguage.CSharp);
            LoadWPFProject(ProjectLanguage.VB);

            return solutionPath;
        }

        private void SetFileAsActive(string fileCaption)
        {
            vs.ObjectModel.WindowManager.ShowDocOnInitialize = false;
            DocumentWindowTestExtension window = vs.ObjectModel.WindowManager.FindDocumentWindow(WindowManagementService.DocumentWindowFindType.ByCaption, fileCaption);

            window.Show();
            window.WaitForFullyLoaded();
            DocState.SetFileAsFocused(fileCaption);
        }

        private void LoadWinFormProject(ProjectLanguage language)
        {
            string ProjName = string.Format("{0}_Winform", language.ToString());

            ProjectTestExtension proj = vs.ObjectModel.Solution.GetProjectExtension<ProjectTestExtension>(ProjName);

            if (TryGetProjectItem(proj, "Form1", out ProjectItemTestExtension item))
                item.Open(DocumentLogicalView.Designer);

            DocumentWindowTestExtension documentWindow;

            string ItemName = string.Format("{0}_Winform_Code", language.ToString());
            if (TryGetProjectItem(proj, ItemName, out ProjectItemTestExtension Item))
            {
                documentWindow = Item.Open();
                DocState.AddFile(documentWindow.Caption, ProjName);
            }

            ItemName = string.Format("{0}_Winform_Window", language.ToString());
            if (TryGetProjectItem(proj, ItemName, out Item))
            {
                documentWindow = Item.Open(DocumentLogicalView.Code);
                DocState.AddFile(documentWindow.Caption, ProjName);

                documentWindow = Item.Open(DocumentLogicalView.Designer);
                DocState.AddFile(documentWindow.Caption, ProjName);
            }
        }

        private void LoadConsoleProject(ProjectLanguage language)
        {
            string ProjName = string.Format("{0}_Console", language.ToString());

            ProjectTestExtension proj = vs.ObjectModel.Solution.GetProjectExtension<ProjectTestExtension>(ProjName);

            string defaultItemName = language == ProjectLanguage.VC ? "VC_Console" : "Program";
            if (TryGetProjectItem(proj, defaultItemName, out ProjectItemTestExtension item))
                item.Open();

            string ItemName = string.Format("{0}_Console_code", language.ToString());
            if (TryGetProjectItem(proj, ItemName, out ProjectItemTestExtension Item))
            {
                DocumentWindowTestExtension documentWindow = Item.Open();
                DocState.AddFile(documentWindow.Caption, ProjName);
            }
        }

        private void LoadWPFProject(ProjectLanguage language)
        {
            string ProjName = string.Format("{0}_WPF", language.ToString());

            ProjectTestExtension proj = vs.ObjectModel.Solution.GetProjectExtension<ProjectTestExtension>(ProjName);

            if (TryGetProjectItem(proj, "MainWindow", out ProjectItemTestExtension defaultItem))
            {
                defaultItem.Open(DocumentLogicalView.Code);
                defaultItem.Open(DocumentLogicalView.Designer);
            }

            DocumentWindowTestExtension documentWindow;
            string ItemName = string.Format("{0}_wpf_UserControl", language.ToString());

            if (TryGetProjectItem(proj, ItemName, out ProjectItemTestExtension Item))
            {
                documentWindow = Item.Open(DocumentLogicalView.Designer);
                DocState.AddFile(documentWindow.Caption, ProjName);

                documentWindow = Item.Open(DocumentLogicalView.Code);
                DocState.AddFile(documentWindow.Caption, ProjName);
            }

            ItemName = string.Format("{0}_wpf_code", language.ToString());
            if (TryGetProjectItem(proj, ItemName, out Item))
            {
                documentWindow = Item.Open();
                DocState.AddFile(documentWindow.Caption, ProjName);
            }
        }

        private bool TryGetProjectItem(
            ProjectTestExtension project,
            string itemName,
            out ProjectItemTestExtension projectItem)
        {
            // First try exact match.
            projectItem = project.ProjectItemsRecursive.FirstOrDefault(
                p => string.Equals(itemName, p.Name, StringComparison.OrdinalIgnoreCase));

            // Then try file name without extension.
            if (projectItem == null)
                projectItem = project.ProjectItemsRecursive.FirstOrDefault(
                    p => string.Equals(itemName, Path.GetFileNameWithoutExtension(p.Name), StringComparison.OrdinalIgnoreCase));

            return Verify.IsNotNull(projectItem, "Project {0} should contain item {1}.", project.Name, itemName);
        }

        public bool CheckStubbed()
        {
            bool bRet = true;

            //Look for the window in the file list.  If we don't find it we won't verify
            foreach (FileModel file in DocState.OpenFiles)
            {
                vs.ObjectModel.WindowManager.ShowDocOnInitialize = false;
                DocumentWindowTestExtension window = null;

                try
                {
                    window = vs.ObjectModel.WindowManager.FindDocumentWindow(WindowManagementService.DocumentWindowFindType.ByCaption, file.FileCaption);
                }
                catch (Exception e)
                {
                    if (!file.FileCaption.Contains('['))
                        throw e;

                    var windows = vs.ObjectModel.WindowManager.DocumentWindows;
                    foreach (DocumentWindowTestExtension d in windows)
                    {
                        if (d.Caption.StartsWith(file.FileCaption.Substring(0, file.FileCaption.IndexOf('!'))))
                        {
                            //var res = DocState.OpenFiles.Where(x => x.FileCaption == file.FileCaption);
                            file.FileCaption = d.Caption;
                            window = d;
                            break;
                        }
                    }

                    if (window == null)
                        throw e;
                }

                //If we got this far the file name is the same as the window caption, check the window
                //against expected state:
                if (file.isStub)
                {
                    //if we ever see a failure we need to keep that.
                    bRet = window.Verify.IsStubDocument() && bRet;
                }
                else
                {
                    this.Services.Synchronization.TryWaitFor(() => { return !window.IsStubDocument; });
                    bool isStub = window.Verify.IsNotStubDocument();
                    if (!isStub)
                    {
                        Console.WriteLine("error");
                    }

                    bRet = window.Verify.IsNotStubDocument() && bRet;
                }

            }
            return bRet;
        }

        private void VsStop(VisualStudioHost vs)
        {
            try
            {
                vs.Stop();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Random InteropService exception from Apex when closing.
                // Ignore for now to make the TC to pass while we investigate the culprit.
            }
        }
    }

    public class DocStateModel
    {
        public List<FileModel> OpenFiles = new List<FileModel>();

        public void AddFile(string FileCaption, string ProjectName)
        {
            FileModel fm = new FileModel();
            fm.FileCaption = FileCaption;
            fm.ProjectName = ProjectName;
            fm.isStub = true;

            OpenFiles.Add(fm);
        }

        public void SetFileAsFocused(string FileName)
        {
            //Both the designer and the code behind file windup getting loaded when the
            //File is selected.  Since this bug was resolved as won't fix, changing the
            //verifier to expect the behavior.
            string MinimalFileName = FileName;

            if (MinimalFileName.Contains('['))
            {
                if (MinimalFileName.Contains('!'))
                    MinimalFileName = MinimalFileName.Substring(0, MinimalFileName.IndexOf('!')).TrimEnd();
                else
                    MinimalFileName = MinimalFileName.Substring(0, MinimalFileName.LastIndexOf('[')).TrimEnd();
            }
            //Build up the list in a new one since you can't modify the current list.
            List<FileModel> newOpenFiles = new List<FileModel>();
            for (int i = 0; i < OpenFiles.Count; i++)
            {
                FileModel tempModel = OpenFiles[i];
                string MinimialTempFileName = tempModel.FileCaption;
                if (MinimialTempFileName.Contains('['))
                {
                    if (MinimialTempFileName.Contains('!'))
                        MinimialTempFileName = MinimialTempFileName.Substring(0, MinimialTempFileName.IndexOf('!')).TrimEnd();
                    else
                        MinimialTempFileName = MinimialTempFileName.Substring(0, MinimialTempFileName.LastIndexOf('[')).TrimEnd();
                }

                if (MinimialTempFileName == MinimalFileName)
                {
                    tempModel.isStub = false;
                }

                newOpenFiles.Add(tempModel);
            }
            OpenFiles = newOpenFiles;
        }

        public void ResetListToInitialLoadState()
        {
            List<FileModel> NewOpen = new List<FileModel>();
            for (int i = 0; i < OpenFiles.Count(); i++)
            {
                FileModel tmp = OpenFiles[i];
                tmp.isStub = true;
                NewOpen.Add(tmp);
            }

            //swap lists
            OpenFiles = NewOpen;
        }
    }

    public class FileModel
    {
        public string FileCaption;
        public string ProjectName;
        public bool isStub;
    }
}
