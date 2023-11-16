using System.Collections.Generic;

using Microsoft.Test.Apex.VisualStudio;
using Microsoft.Test.Apex.VisualStudio.FeatureFlags;
using Microsoft.Test.Apex.VisualStudio.Settings;
using Microsoft.VisualStudio.PlatformUI.Apex.GetToCode;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VsIde.IdeExp.GetToCodeTests
{
    public abstract class GetToCodeTestBase : VisualStudioHostTest
    {
        protected const int FiveMinutesInMs = 5 * 60 * 1000;
        protected const int FifteenMinutesInMs = 15 * 60 * 1000;
        protected const string IdeExperienceEngineersAlias = "IDE Experience";

        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();

            this.GetToCodeService = this.GetAndCheckService<GetToCodeTestService>();

            var settings = this.GetAndCheckService<SettingsService>();
            if (settings.Environment.General.OnStartup != EnvironmentSettings.StartupOption.GetToCode)
            {
                this.SetStartupOption(EnvironmentSettings.StartupOption.GetToCode);

                var featureFlag = this.GetAndCheckService<FeatureFlagService>();

                // hide start page for G2C test
                if (featureFlag.GetFeatureFlag(@"VS.Core.StartPage"))
                {
                    featureFlag.SetFeatureFlag(@"VS.Core.StartPage", false);
                }
                this.RestartVisualStudio(); // resets this.GetToCodeService
            }
        }

        private void SetStartupOption(EnvironmentSettings.StartupOption option)
        {
            var settings = this.GetAndCheckService<SettingsService>();
            settings.Environment.General.OnStartup = option;
        }

        [TestCleanup]
        public override void TestCleanup()
        {
            base.TestCleanup();

            this.VisualStudio.Stop();
        }

        /// <summary>
        /// Gets a value indicating whether VS should inherit the parent process environment. Optimization tests should
        /// set this to true so in-proc data collector can work.
        /// </summary>
        protected virtual bool InheritParentProcessEnvironment => false;

        protected abstract StartWindowStartupOption StartWindowStartupOption { get; }

        protected GetToCodeTestService GetToCodeService { get; set; }

        protected string VisualStudioCommandLine { get; set; }

        protected override VisualStudioHostConfiguration GetHostConfiguration()
        {
            return new GetToCodeHostConfiguration
            {
                CommandLineArguments = this.VisualStudioCommandLine,
                InheritProcessEnvironment = this.InheritParentProcessEnvironment,
                StartWindowStartupOption = this.StartWindowStartupOption,
            };
        }

        protected T GetAndCheckService<T>() where T : class
        {
            var testService = this.VisualStudio.Get<T>();
            this.Verify.IsNotNull(testService, $"{typeof(T).Name} not found.");
            return testService;
        }

        protected void RestartVisualStudio()
        {
            this.VisualStudio.Stop(); // the base class will automatically restart VS when the VisualStudio property getter is next invoked
            this.GetToCodeService = this.GetAndCheckService<GetToCodeTestService>();
        }

        /// <summary>
        /// Custom Visual Studio host configuration to ensure Microsoft.VisualStudio.UI.Internal.Apex is included in
        /// Apex MEF composition
        /// </summary>
        private class GetToCodeHostConfiguration : VisualStudioHostConfiguration
        {
            public override IEnumerable<string> CompositionAssemblies
            {
                get
                {
                    HashSet<string> assemblies = new HashSet<string>(base.CompositionAssemblies);
                    assemblies.Add(typeof(GetToCodeTestService).Assembly.Location);
                    return assemblies;
                }
            }
        }
    }
}
