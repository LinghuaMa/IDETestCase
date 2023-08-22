using Microsoft.Test.Apex;
using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.Solution;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Reflection;

namespace Example // Namespace must be the same as assembly name
{
    [TestClass]
    public class Tests : ApexTest // Derive from ApexTest to plug into the Apex framework
    {
        /// <summary>
        /// Gets the directory where the test is currently executing from.
        /// </summary>
        //public string TestExecutionDirectory { get => Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); }

        [TestMethod]
        public void ExampleTest()
        {
            // Start an instance of VS
            var visualStudio = this.Operations.CreateHost<VisualStudioHost>();
            visualStudio.Start();

            // Create a new solution and add a C# console app
           visualStudio.ObjectModel.Solution.CreateEmptySolution("HelloWorld", Path.Combine(@"C:\ApexTest\Example10"));
           //visualStudio.ObjectModel.Solution.CreateDefaultProject();
            var project = visualStudio.ObjectModel.Solution.AddProject(ProjectLanguage.CSharp, ProjectTemplate.ConsoleApplication, "HelloWorld");

            // Add FileName.cs and get a handle to the editor

            var folder = project.AddFolder("New Folder");
            var newItem = folder.AddDefaultProjectItem();
            var newDocument = newItem.GetDocumentAsTextEditor();
            var newEditor = newDocument.Editor;

            // Open Program.cs and get a handle to the editor
            var item = project["Program.cs"];
            var document = item.GetDocumentAsTextEditor();
            var editor = document.Editor;

            // Add some code to the Main method
            editor.Caret.MoveToExpression("static void Main");
            editor.Caret.MoveDown();
            editor.Caret.MoveToEndOfLine();
            editor.KeyboardCommands.Enter();
            editor.KeyboardCommands.Type("Console.WriteLine(\"Hello World!\");");
            editor.KeyboardCommands.Enter();
            editor.KeyboardCommands.Type("System.Threading.Thread.Sleep(TimeSpan.FromSeconds(5));");

            // Save everything and close
            visualStudio.ObjectModel.Solution.SaveAndClose();
            visualStudio.Stop();
        }
    }
}