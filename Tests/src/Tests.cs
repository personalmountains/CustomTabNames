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

			Operations.SetExtensionOption("Template",
				"$(ProjectName):$(FilterPath):$(Filename)");

			Operations.SetExtensionOption("IgnoreBuiltinProjects", true);
			Operations.SetExtensionOption("IgnoreSingleProject", true);
			Operations.SetExtensionOption("LoggingLevel", 4);
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


	[TestClass]
	public class CppTests
	{
		public static VS vs;
		public static Operations ops;

		[ClassInitialize]
		public static void Init(TestContext _)
		{
			vs = Global.VS;
			ops = Global.Operations;
		}

		[TestMethod]
		public void AddRemoveProjectsIgnoreSingle()
		{
			using (ops.ScopedExtensionOption("IgnoreSingleProject", true))
			{
				var w = ops.OpenFile(Global.CPP, "f.cpp");
				Assert.AreEqual("cpp::f.cpp", w.Caption);

				string csPath = Global.CS.FullName;

				ops.RemoveProject(Global.CS);
				Assert.AreEqual("::f.cpp", w.Caption);

				ops.AddProject(csPath);
				Assert.AreEqual("cpp::f.cpp", w.Caption);
			}
		}

		[TestMethod]
		public void AddRemoveProjectsDontIgnoreSingle()
		{
			using (ops.ScopedExtensionOption("IgnoreSingleProject", false))
			{
				var w = ops.OpenFile(Global.CPP, "f.cpp");
				Assert.AreEqual("cpp::f.cpp", w.Caption);

				string csPath = Global.CS.FullName;

				ops.RemoveProject(Global.CS);
				Assert.AreEqual("cpp::f.cpp", w.Caption);
				ops.AddProject(csPath);
				Assert.AreEqual("cpp::f.cpp", w.Caption);
			}
		}

		[TestMethod]
		public void MoveBetweenRootAndFolders()
		{
			var w = ops.OpenFile(Global.CPP, "f.cpp");
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
