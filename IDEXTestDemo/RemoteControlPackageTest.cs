// ===================================================================================================--
// 
// Remote Control Service
// 
// ===================================================================================================--
namespace vside.RemoteControl
{
    using Microsoft.Test.Apex;
    using Microsoft.Test.Apex.VisualStudio;
    using Microsoft.Test.Apex.VisualStudio.Services;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Microsoft.Win32;

    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Threading;

    [TestClass]
    [DeploymentItem("Remote.pkgdef")]
    [DeploymentItem("RemoteSettings_Common_17.0.json")]
    public class RemoteControlPackageTest : VisualStudioHostTest
    {
        private string EnvVarValue = string.Empty;
        private static RegistryKey user = Registry.CurrentUser;
        private static RegistryKey s = user.OpenSubKey("Software", true);
        private static RegistryKey m = s.OpenSubKey("Microsoft", true);
        private static RegistryKey vs = m.OpenSubKey("VisualStudio", true);

        [TestInitialize]
        [Timeout(5 * 60 * 1000)]
        public override void TestInitialize()
        {
            base.TestInitialize();

            // caches the value of the environment variable used to point to a 'test' version of the settings json so we can restore it upon Cleanup
            // delete the value for now because we will be setting it throughout the test
            EnvVarValue = Environment.GetEnvironmentVariable("RemoteSettingsFilePath");
            Environment.SetEnvironmentVariable("RemoteSettingsFilePath", "", EnvironmentVariableTarget.User);
        }

        [TestCleanup]
        public override void TestCleanup()
        {
            // restore the test Environment variable
            Environment.SetEnvironmentVariable("RemoteSettingsFilePath", EnvVarValue, EnvironmentVariableTarget.User);

            // call cleanup
            base.TestCleanup();
        }

        [TestMethod]
        [Timeout(5 * 60 * 1000)]
        //[Owner(vside.ProjectTests.TestOwners.PnT)]
        [Description("RemoteControl Service Test")]
        public void RemoteControlServiceTest()
        {
            IVisualStudioPropertyManager propertyManager = this.VisualStudio.Get<IVisualStudioPropertyService>();

            string testDir = this.TestExecutionDirectory.FullName;
            string sourceFile = Path.Combine(testDir, "RemoteSettings_Common_17.0.json");

            // verify service starts and creates expected assets
            using (Scope.Enter("verify service starts and creates expected assets (1st-launch state)"))
            {
                var apexHostConfig = new VisualStudioHostConfiguration() { RestoreUserSettings = false };
                var visualStudio = this.Operations.CreateHost<VisualStudioHost>(apexHostConfig);
                if (visualStudio.IsRunning)
                    visualStudio.Stop();

                // State --> delete RemoteSettings user settings key 
                if (IsRemoteSettingsExist("RemoteSettings"))
                {
                    vs.DeleteSubKeyTree("RemoteSettings", true);
                }

                // State --> delete InetCache (remotesettings.json)
                DeleteInetCache("remotesettings*.cache");

                // Action --> start VS and wait to reach idle
                visualStudio.Start();
                visualStudio.WaitForProcessCpuIdle();

                // Validate --> RemoteSettings user settings key created                       
                this.Verify.IsTrue(IsRemoteSettingsExist("RemoteSettings"), "RemoteSettings key exist in the registry");

                // Validate --> Remote hive created and has sub-keys
                IList<string> keyRemoteHive = (IList<string>)propertyManager.GetRemoteSubCollectionNames(null);
                if (this.Verify.IsNotEmpty(keyRemoteHive, "{0} sub key exists", "RemoteStoreHive"))
                {
                    this.Verify.IsTrue(keyRemoteHive.Count > 0, "{0} sub key has {1} sub keys", "RemoteStoreHive", keyRemoteHive.Count);
                }

                visualStudio.Stop();
            }

            // verify service properly consumes settings file pointed to by environment variable (supercedes configuration settings)
            using (Scope.Enter("verify service properly consumes settings file pointed to by environment variable"))
            {
                Environment.SetEnvironmentVariable("RemoteSettingsFilePath", sourceFile, EnvironmentVariableTarget.User);

                vs.CreateSubKey("RemoteControl");
                RegistryKey re = vs.OpenSubKey("RemoteControl", true);
                re.CreateSubKey("TestUrlMapping");
                RegistryKey te = re.OpenSubKey("TestUrlMapping", true);
                te.SetValue("https://az700632.vo.msecnd.net/pub/RemoteSettings", testDir);

                // Action --> start VS via the cmd-prompt and wait for idle
                var apexHostConfig = new VisualStudioHostConfiguration() { RestoreUserSettings = false };
                apexHostConfig.Environment["RemoteSettingsFilePath"] = sourceFile;
                var visualStudio = this.Operations.CreateHost<VisualStudioHost>(apexHostConfig);
                visualStudio.Start();
                visualStudio.WaitForProcessCpuIdle();

                // Validate --> RemoteSettings created based on settings in pointed-to settings file            
                IEnumerable<string> keyRemoteHive = propertyManager.GetRemoteSubCollectionNames(null);
                this.Verify.IsNotEmpty(keyRemoteHive, "{0} sub key exists", "RemoteStoreHive");
                if (keyRemoteHive != null)
                {
                    var testSettingKey = propertyManager.GetRemoteSubCollectionNames("TestEnvVarKey");
                    Verify.IsNotNull(testSettingKey, "verify TestEnvVarKey is set in the Remote store hive");

                    var testSettingValue = (string)propertyManager.GetRemoteProperty("TestEnvVarKey", "Test");
                    Verify.IsTrue(testSettingValue.Equals("EnvVarTestValue1"), "testSettingValue %1 have the expected value %2", testSettingValue, "EnvVarTestValue1");
                }

                this.Verify.IsTrue(IsTestEnvVarKeyExist("TestEnvVarKey"), "TestEnvVarKey key exist under RemoteSettings");

                RegistryKey remoteKey = vs.OpenSubKey("RemoteSettings", true);
                RegistryKey remoteSubKey = remoteKey.OpenSubKey("RemoteSettings_Common_17.0.json", true);
                string fileVersion = remoteSubKey.GetValue("FileVersion").ToString();
                this.Verify.IsTrue(fileVersion == "1.0", "The data of FileVersion is 1.0");

                visualStudio.Stop();
            }

            // verify service does not update remote store when Version is stale
            using (Scope.Enter("verify service does not update remote store when Version is stale"))
            {
                RegistryKey remoteKey = vs.OpenSubKey("RemoteSettings", true);
                RegistryKey remoteSubKey = remoteKey.OpenSubKey("RemoteSettings_Common_17.0.json", true);
                RegistryKey remoteSubKey1 = remoteSubKey.OpenSubKey("1.0", true);
                RegistryKey remoteSubKey2 = remoteSubKey1.OpenSubKey("TestEnvVarKey", true);

                // State --> update settings value in settings file to be different
                string[] fileInput =
                {
                    "{",
                    "   \"FileVersion\":\"1.0\",",
                    "   \"ChangesetId\":\"12345\",",
                    "   \"TestEnvVarKey\": {",
                    "       \"Test\": \"EnvVarTestValue2\"",    //change 1 to 2
                    "   }",
                    "}"
                };

                using (StreamWriter file = new StreamWriter(Path.Combine(testDir, "RemoteSettings_Common_17.0.json")))
                {
                    foreach (string line in fileInput)
                    {
                        file.WriteLine(line);
                    }
                }

                // Action --> start VS and wait for idle
                var apexHostConfig = new VisualStudioHostConfiguration() { RestoreUserSettings = false };
                apexHostConfig.InheritProcessEnvironment = true;
                apexHostConfig.Environment["RemoteSettingsFilePath"] = sourceFile;
                var visualStudio = this.Operations.CreateHost<VisualStudioHost>(apexHostConfig);
                visualStudio.Start();
                visualStudio.WaitForProcessCpuIdle();

                String testData = remoteSubKey2.GetValue("Test").ToString();
                this.Verify.IsTrue(testData.Equals("EnvVarTestValue1"), "testSettingValue is not changed, value is EnvVarTestValue1");

                visualStudio.Stop();
            }

            // verify service updates remote store when Version is incremented
            using (Scope.Enter("verify service updates remote store when Version is incremented"))
            {
                // State --> update Version in settings file to be different
                string[] fileInput =
                {
                    "{",
                    "   \"FileVersion\":\"2.0\",",              //change 1 to 2
                    "   \"ChangesetId\":\"12345\",",
                    "   \"TestEnvVarKey\": {",
                    "       \"Test\": \"EnvVarTestValue2\"",    //change 1 to 2
                    "   }",
                    "}"
                };

                using (StreamWriter file = new StreamWriter(Path.Combine(testDir, "RemoteSettings_Common_17.0.json")))
                {
                    foreach (string line in fileInput)
                    {
                        file.WriteLine(line);
                    }
                }

                // Action --> start VS and wait for idle 
                var apexHostConfig = new VisualStudioHostConfiguration() { RestoreUserSettings = false };
                apexHostConfig.InheritProcessEnvironment = true;
                apexHostConfig.Environment["RemoteSettingsFilePath"] = sourceFile;
                var visualStudio = this.Operations.CreateHost<VisualStudioHost>(apexHostConfig);
                visualStudio.Start();
                visualStudio.WaitForProcessCpuIdle();

                // Validate --> settings value in Remote reg hive changed as expected               
                RegistryKey remoteKey = vs.OpenSubKey("RemoteSettings", true);
                RegistryKey remoteSubKey = remoteKey.OpenSubKey("RemoteSettings_Common_17.0.json", true);
                this.Verify.IsTrue(IsTestSubKeyExist("2.0"), "Subkey 2.0 exists when change all versions");

                RegistryKey remoteSubKey1 = remoteSubKey.OpenSubKey("2.0", true);
                RegistryKey remoteSubKey2 = remoteSubKey1.OpenSubKey("TestEnvVarKey", true);
                String testData = remoteSubKey2.GetValue("Test").ToString();
                this.Verify.IsTrue(testData.Equals("EnvVarTestValue2"), "testSettingValue is changed, value is EnvVarTestValue2");

                visualStudio.Stop();
            }

            // verify service works correctly when INetCache is invalidated
            using (Scope.Enter("verify service works correctly when INetCache is invalidated"))
            {
                // State --> remove environment variable
                Environment.SetEnvironmentVariable("RemoteSettingsFilePath", "", EnvironmentVariableTarget.User);
                vs.DeleteSubKeyTree("RemoteControl", true);
                // State --> delete RemoteSettings*.json from InetCache
                DeleteInetCache("remotesettings*.cache");

                // Action --> start VS and wait for idle
                var apexHostConfig = new VisualStudioHostConfiguration() { RestoreUserSettings = false };
                var visualStudio = this.Operations.CreateHost<VisualStudioHost>(apexHostConfig);
                visualStudio.Start();
                visualStudio.WaitForProcessCpuIdle();
                visualStudio.Stop();

                // Validate --> RemoteSettings*.json is added back to InetCache
                Verify.IsTrue(IsInInetCache("remotesettings*.cache"));
            }

            // verify service selects the appropriate settings file from list of available files described in the Master.json
            //// NYA

            // verify no unhandled exception when Master.json points to files that don't exist on host
            //using (Scope.Enter("verify no unhandled exception when Master.json points to files that don't exist on host"))
            //{
            // State --> update Master.json in InetCache to contain only files that don't exist

            // Action --> start VS and wait for Idle

            // Validate --> VS did not crash

            //}

            // verify no unhandled exception when settings file contains unexpected json - validate that the previously archived settings to the Remote hive are persisted
            //using (Scope.Enter("verify no unhandled exception when Master.json points to files that don't exist on host"))
            //{
            // State --> update Master.json in InetCache to contain only files that don't exist

            // Action --> start VS and wait for Idle

            // Validate --> VS did not crash

            //}

            // verify service is properly disabled via its feature flag
            ////NYA

            if (this.VisualStudio.IsRunning)
                this.VisualStudio.Stop();
        }

        //Verify if Subkey 2.0 exist, if exist, return true.
        private bool IsTestSubKeyExist(string keyName)
        {
            RegistryKey remoteKey = vs.OpenSubKey("RemoteSettings", true);
            RegistryKey remoteSubKey = remoteKey.OpenSubKey("RemoteSettings_Common_17.0.json", true);
            string[] subkeyNames = remoteSubKey.GetSubKeyNames();

            foreach (string name in subkeyNames)
            {
                if (name == keyName)
                {
                    return true;
                }
            }

            return false;
        }

        //Verify if Subkey TestEnvVarKey exist, if exist, return true.
        private bool IsTestEnvVarKeyExist(string keyName)
        {
            RegistryKey remoteKey = vs.OpenSubKey("RemoteSettings", true);
            RegistryKey remoteSubKey = remoteKey.OpenSubKey("RemoteSettings_Common_17.0.json", true);
            RegistryKey remoteSubKey1 = remoteSubKey.OpenSubKey("1.0", true);
            string[] subkeyNames = remoteSubKey1.GetSubKeyNames();

            foreach (string name in subkeyNames)
            {
                if (name == keyName)
                {
                    return true;
                }
            }

            return false;
        }

        //Verify if key RemoteSettings exist, if exist, return true.
        private bool IsRemoteSettingsExist(string keyName)
        {
            string[] subkeyNames = vs.GetSubKeyNames();
            foreach (string name in subkeyNames)
            {
                if (name == keyName)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Deletes files matching the searchPattern from the current user's InetCache
        /// </summary>
        /// <param name="searchPattern">the files to match. Eg: *.*</param>
        private void DeleteInetCache(string searchPattern)
        {
            string localAppData = Environment.ExpandEnvironmentVariables("%LOCALAPPDATA%");
            string searchDirectory = Path.Combine(localAppData, "Microsoft", "Windows", "INetCache");

            //Directory.SetAccessControl
            foreach (var fileName in GetFiles(searchDirectory, searchPattern))
            {
                File.Delete(fileName);
            }
        }

        /// <summary>
        /// Returns True if file matching the search pattern exists in the INetCache folders
        /// </summary>
        /// <param name="searchPattern"></param>
        /// <returns></returns>
        private bool IsInInetCache(string searchPattern)
        {
            string localAppData = Environment.ExpandEnvironmentVariables("%LOCALAPPDATA%");
            string searchDirectory = Path.Combine(localAppData, "Microsoft", "Windows", "INetCache");

            //Directory.SetAccessControl
            foreach (var fileName in GetFiles(searchDirectory, searchPattern))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// traverses root (including all sub-folders) and returns a list of files matching searchPattern. Ignores exceptions (such as subfolders that are ACL protected and result in AccessDenied)
        /// </summary>
        /// <param name="root">starting folder to search from (searches all sub folders)</param>
        /// <param name="searchPattern">the files to match. Eg: *.*</param>
        /// <returns></returns>
        private List<string> GetFiles(string root, string searchPattern)
        {
            List<string> retval = new List<string>();

            Stack<string> pending = new Stack<string>();
            pending.Push(root);
            while (pending.Count != 0)
            {
                var path = pending.Pop();
                string[] next = null;
                try
                {
                    next = Directory.GetFiles(path, searchPattern);
                }
                catch { }
                if (next != null && next.Length != 0)
                    foreach (var file in next)
                    {
                        retval.Add(file);
                    }
                try
                {
                    next = Directory.GetDirectories(path);
                    foreach (var subdir in next) pending.Push(subdir);
                }
                catch { }
            }

            return retval;
        }

    }
}
