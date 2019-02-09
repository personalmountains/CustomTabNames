using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CustomTabNames.Tests
{
	[TestClass]
	public class Global
	{
		static public Main Main { get; private set; }
		static public TestOptions Options { get; private set; }
		static public TestLogger Logger { get; private set; }
		static public TestSolution Solution { get; private set; }
		static public TestDocumentManager DocumentManager { get; private set; }

		[AssemblyInitialize]
		public static void StartTests(TestContext _)
		{
			Options = new TestOptions();
			Logger = new TestLogger();
			Solution = new TestSolution();
			DocumentManager = new TestDocumentManager();

			Main = new Main(Options, Logger, Solution, DocumentManager);
		}

		[AssemblyCleanup]
		public static void StopTests()
		{
		}
	}

	[TestClass]
	public class OptionTests
	{
		[TestMethod]
		public void Callbacks()
		{
			var o = Main.Instance.Options;

			// known values
			o.Enabled = true;
			o.Template = "";
			o.IgnoreBuiltinProjects = true;
			o.IgnoreSingleProject = true;
			o.Logging = true;
			o.LoggingLevel = 0;

			bool enabled = false, template = false, builtin = false;
			bool single = false, logging = false, level = false;

			o.EnabledChanged += () => { enabled = true; };
			o.TemplateChanged += () => { template = true; };
			o.IgnoreBuiltinProjectsChanged += () => { builtin = true; };
			o.IgnoreSingleProjectChanged += () => { single = true; };
			o.LoggingChanged += () => { logging = true; };
			o.LoggingLevelChanged += () => { level = true; };

			o.Enabled = false;
			o.Template = "1";
			o.IgnoreBuiltinProjects = false;
			o.IgnoreSingleProject = false;
			o.Logging = false;
			o.LoggingLevel = 1;

			Assert.IsTrue(enabled);
			Assert.IsTrue(template);
			Assert.IsTrue(builtin);
			Assert.IsTrue(single);
			Assert.IsTrue(logging);
			Assert.IsTrue(level);
		}
	}
}
