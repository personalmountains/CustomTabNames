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


	[TestClass]
	public class CppTests
	{
		public static VS vs;
		public static Operations ops;
		public static Window file;

		[ClassInitialize]
		public static void Init(TestContext _)
		{
			vs = Global.VS;
			ops = Global.Operations;
			file = ops.OpenFile(Global.CPP, "f.cpp");
		}

		[TestMethod]
		public void AddRemoveProjectsIgnoreSingle()
		{
			Assert.AreEqual("cpp::f.cpp", file.Caption);

			using (ops.SetOptionTemp("IgnoreSingleProject", true))
			{
				using (ops.RemoveProjectTemp(Global.CS))
					Assert.AreEqual("::f.cpp", file.Caption);

				Assert.AreEqual("cpp::f.cpp", file.Caption);
			}
		}

		[TestMethod]
		public void AddRemoveProjectsDontIgnoreSingle()
		{
			Assert.AreEqual("cpp::f.cpp", file.Caption);

			using (ops.SetOptionTemp("IgnoreSingleProject", false))
			{
				using (ops.RemoveProjectTemp(Global.CS))
					Assert.AreEqual("cpp::f.cpp", file.Caption);

				Assert.AreEqual("cpp::f.cpp", file.Caption);
			}
		}

		[TestMethod]
		public void RenameProject()
		{
			Assert.AreEqual("cpp::f.cpp", file.Caption);

			using (ops.RenameProjectTemp(Global.CPP, "cpp2"))
				Assert.AreEqual("cpp2::f.cpp", file.Caption);

			Assert.AreEqual("cpp::f.cpp", file.Caption);
		}

		[TestMethod]
		public void RenameFolder()
		{
			Assert.AreEqual("cpp::f.cpp", file.Caption);

			using (ops.MoveFileTemp(@"test\cpp\f.cpp", @"test\cpp\a\b\c"))
			{
				Assert.AreEqual("cpp:a/b/c:f.cpp", file.Caption);

				using (ops.RenameFolderTemp(@"test\cpp\a", "aa"))
				{
					Assert.AreEqual("cpp:aa/b/c:f.cpp", file.Caption);

					using (ops.RenameFolderTemp(@"test\cpp\aa\b", "bb"))
					{
						Assert.AreEqual("cpp:aa/bb/c:f.cpp", file.Caption);

						using (ops.RenameFolderTemp(@"test\cpp\aa\bb\c", "cc"))
							Assert.AreEqual("cpp:aa/bb/cc:f.cpp", file.Caption);

						Assert.AreEqual("cpp:aa/bb/c:f.cpp", file.Caption);
					}

					Assert.AreEqual("cpp:aa/b/c:f.cpp", file.Caption);
				}

				Assert.AreEqual("cpp:a/b/c:f.cpp", file.Caption);
			}
		}

		[TestMethod]
		public void MoveBetweenRootAndFolders()
		{
			Assert.AreEqual("cpp::f.cpp", file.Caption);

			// / to /a
			using (ops.MoveFileTemp(@"test\cpp\f.cpp", @"test\cpp\a"))
			{
				Assert.AreEqual("cpp:a:f.cpp", file.Caption);

				// /a to /a/b
				using (ops.MoveFileTemp(@"test\cpp\a\f.cpp", @"test\cpp\a\b"))
				{
					Assert.AreEqual("cpp:a/b:f.cpp", file.Caption);

					// /a/b to /a/b/c
					using (ops.MoveFileTemp(@"test\cpp\a\b\f.cpp", @"test\cpp\a\b\c"))
					{
						Assert.AreEqual("cpp:a/b/c:f.cpp", file.Caption);

						// /a/b/c to d/e/f
						using (ops.MoveFileTemp(@"test\cpp\a\b\c\f.cpp", @"test\cpp\d\e\f"))
						{
							Assert.AreEqual("cpp:d/e/f:f.cpp", file.Caption);
						}

						Assert.AreEqual("cpp:a/b/c:f.cpp", file.Caption);
					}

					Assert.AreEqual("cpp:a/b:f.cpp", file.Caption);
				}

				Assert.AreEqual("cpp:a:f.cpp", file.Caption);
			}

			Assert.AreEqual("cpp::f.cpp", file.Caption);
		}
	}
}
