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
			AssertCaption("cpp::f.cpp");

			using (ops.SetOptionTemp("IgnoreSingleProject", true))
			{
				using (ops.RemoveProjectTemp(Global.CS))
					AssertCaption("::f.cpp");

				AssertCaption("cpp::f.cpp");
			}
		}

		[TestMethod]
		public void AddRemoveProjectsDontIgnoreSingle()
		{
			AssertCaption("cpp::f.cpp");

			using (ops.SetOptionTemp("IgnoreSingleProject", false))
			{
				using (ops.RemoveProjectTemp(Global.CS))
					AssertCaption("cpp::f.cpp");

				AssertCaption("cpp::f.cpp");
			}
		}

		[TestMethod]
		public void RenameProject()
		{
			AssertCaption("cpp::f.cpp");

			using (ops.RenameProjectTemp(Global.CPP, "cpp2"))
				AssertCaption("cpp2::f.cpp");

			AssertCaption("cpp::f.cpp");
		}

		[TestMethod]
		public void RenameFolder()
		{
			AssertCaption("cpp::f.cpp");

			using (ops.MoveFileTemp(@"test\cpp\f.cpp", @"test\cpp\a\b\c"))
			{
				AssertCaption("cpp:a/b/c:f.cpp");

				using (ops.RenameFolderTemp(@"test\cpp\a", "1"))
				{
					AssertCaption("cpp:1/b/c:f.cpp");

					using (ops.RenameFolderTemp(@"test\cpp\1\b", "2"))
					{
						AssertCaption("cpp:1/2/c:f.cpp");

						using (ops.RenameFolderTemp(@"test\cpp\1\2\c", "3"))
							AssertCaption("cpp:1/2/3:f.cpp");

						AssertCaption("cpp:1/2/c:f.cpp");
					}

					AssertCaption("cpp:1/b/c:f.cpp");
				}

				AssertCaption("cpp:a/b/c:f.cpp");
			}
		}

		[TestMethod]
		public void MoveFolder()
		{
			AssertCaption("cpp::f.cpp");

			using (ops.MoveFileTemp(@"test\cpp\f.cpp", @"test\cpp\a\b\c"))
			{
				AssertCaption("cpp:a/b/c:f.cpp");

				using (ops.MoveFolderTemp(@"test\cpp\a\b\c", @"test\cpp"))
				{
					AssertCaption("cpp:c:f.cpp");

					using (ops.MoveFolderTemp(@"test\cpp\c", @"test\cpp\a"))
					{
						AssertCaption("cpp:a/c:f.cpp");

						using (ops.MoveFolderTemp(@"test\cpp\a\c", @"test\cpp\a\b"))
							AssertCaption("cpp:a/b/c:f.cpp");

						AssertCaption("cpp:a/c:f.cpp");
					}

					AssertCaption("cpp:c:f.cpp");
				}

				AssertCaption("cpp:a/b/c:f.cpp");
			}

			AssertCaption("cpp::f.cpp");
		}

		[TestMethod]
		public void RenameFile()
		{
			AssertCaption("cpp::f.cpp");

			using (ops.RenameFileTemp(@"test\cpp\f.cpp", "f1.cpp"))
				AssertCaption("cpp::f1.cpp");

			AssertCaption("cpp::f.cpp");
		}

		[TestMethod]
		public void MoveFile()
		{
			AssertCaption("cpp::f.cpp");

			using (ops.MoveFileTemp(@"test\cpp\f.cpp", @"test\cpp\a"))
			{
				AssertCaption("cpp:a:f.cpp");

				using (ops.MoveFileTemp(@"test\cpp\a\f.cpp", @"test\cpp\a\b"))
				{
					AssertCaption("cpp:a/b:f.cpp");

					using (ops.MoveFileTemp(@"test\cpp\a\b\f.cpp", @"test\cpp\a\b\c"))
					{
						AssertCaption("cpp:a/b/c:f.cpp");

						using (ops.MoveFileTemp(@"test\cpp\a\b\c\f.cpp", @"test\cpp\d\e\f"))
						{
							AssertCaption("cpp:d/e/f:f.cpp");
						}

						AssertCaption("cpp:a/b/c:f.cpp");
					}

					AssertCaption("cpp:a/b:f.cpp");
				}

				AssertCaption("cpp:a:f.cpp");
			}

			AssertCaption("cpp::f.cpp");
		}

		[TestMethod]
		public void OpenFile()
		{
			file.Close();

			file = ops.OpenFile(Global.CPP, "f.cpp");
			AssertCaption("cpp::f.cpp");
		}

		private void AssertCaption(string s)
		{
			Assert.AreEqual(s, file.Caption);
		}
	}
}
