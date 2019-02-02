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
		public static Project CPP { get; private set; }
		public static Project CS { get; private set; }

		[AssemblyInitialize]
		public static void StartTests(TestContext _)
		{
			VS = new VS(SolutionPath);
			Operations = VS.Operations;

			Operations.SetExtensionOption("Template",
				"$(ProjectName):$(FilterPath):$(Filename)");

			Operations.SetExtensionOption("IgnoreBuiltinProjects", true);
			Operations.SetExtensionOption("IgnoreSingleProject", true);

			CPP = Operations.FindProject("cpp");
			CS = Operations.FindProject("cs");
		}

		[AssemblyCleanup]
		public static void StopTests()
		{
			VS.Dispose();
		}
	}


	[TestClass]
	public class CppTests
	{
		public static VS vs;
		public static Operations ops;
		public static Project project;

		[ClassInitialize]
		public static void Init(TestContext _)
		{
			vs = Global.VS;
			ops = Global.Operations;
			project = Global.CPP;
		}

		[TestMethod]
		public void MoveBetweenRootAndFolders()
		{
			var w = ops.OpenFile(project, "f.cpp");

			Assert.AreEqual("cpp::f.cpp", w.Caption);

			// / to /1
			ops.MoveFile(
				@"test\cpp\f.cpp",
				@"test\cpp\1");

			Assert.AreEqual("cpp:1:f.cpp", w.Caption);

			// /1 to /1/1.1
			ops.MoveFile(
				@"test\cpp\1\f.cpp",
				@"test\cpp\1\1.1");

			Assert.AreEqual("cpp:1/1.1:f.cpp", w.Caption);

			// /1/1.1 to /1/1.1/1.1.1
			ops.MoveFile(
				@"test\cpp\1\1.1\f.cpp",
				@"test\cpp\1\1.1\1.1.1");

			Assert.AreEqual("cpp:1/1.1/1.1.1:f.cpp", w.Caption);

			// /1/1.1/1.1.1 to 2/2.1/2.1.1
			ops.MoveFile(
				@"test\cpp\1\1.1\1.1.1\f.cpp",
				@"test\cpp\2\2.1\2.1.1");

			Assert.AreEqual("cpp:2/2.1/2.1.1:f.cpp", w.Caption);

			// back to root
			ops.MoveFile(
				@"test\cpp\2\2.1\2.1.1\f.cpp",
				@"test\cpp");

			Assert.AreEqual("cpp::f.cpp", w.Caption);
		}
	}
}
