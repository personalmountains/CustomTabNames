using Microsoft.VisualStudio.TestTools.UnitTesting;
using EnvDTE;

namespace CustomTabNames.Tests
{
	[TestClass]
	public class Global
	{
		private const string SolutionPath =
			"c:\\dev\\projects\\CustomTabNames\\Tests\\data\\test.sln";

		public static VS VS { get; private set; }
		public static Operations Operations { get; private set; }

		[AssemblyInitialize]
		public static void StartTests(TestContext _)
		{
			VS = new VS(SolutionPath);
			Operations = VS.Operations;

			Operations.SetOption("Template",
				"$(ProjectName):$(FolderPath):$(Filename)");

			Operations.SetOption("IgnoreBuiltinProjects", true);
			Operations.SetOption("IgnoreSingleProject", true);
			Operations.SetOption("LoggingLevel", 4);
		}

		[AssemblyCleanup]
		public static void StopTests()
		{
			VS.Dispose();
		}

		public static Project CPP
		{
			get
			{
				return Operations.GetProject("cpp");
			}
		}

		public static Project CS
		{
			get
			{
				return Operations.GetProject("cs");
			}
		}
	}


}
