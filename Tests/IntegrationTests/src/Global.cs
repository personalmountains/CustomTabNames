using Microsoft.VisualStudio.TestTools.UnitTesting;
using EnvDTE;

namespace CustomTabNames.Tests
{
	[TestClass]
	public class Global
	{
		private const string SolutionPath =
			"..\\..\\..\\data\\test.sln";

		public static VS VS { get; private set; }
		public static Operations Operations { get; private set; }

		[AssemblyInitialize]
		public static void StartTests(TestContext _)
		{
			VS = new VS(System.IO.Path.GetFullPath(SolutionPath));
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


	public class Tests
	{
		private readonly Operations ops;
		private readonly string project, filename, p;
		private Window file;

		public Tests(string projectName, string filename)
		{
			this.ops = Global.Operations;
			this.project = projectName;
			this.filename = filename;
			this.file = ops.OpenFile(ops.GetProject(projectName), filename);

			// prefix
			this.p = @"test\" + project;
		}

		public void AddRemoveProjectsIgnoreSingle(Project otherProject)
		{
			AssertCaption("");

			using (ops.SetOptionTemp("IgnoreSingleProject", true))
			{
				AssertCaption("");

				using (ops.RemoveProjectTemp(otherProject))
					AssertCaption("", "", filename);

				AssertCaption("");
			}
		}

		public void AddRemoveProjectsDontIgnoreSingle(Project otherProject)
		{
			AssertCaption("");

			using (ops.SetOptionTemp("IgnoreSingleProject", false))
			{
				AssertCaption("");

				using (ops.RemoveProjectTemp(otherProject))
					AssertCaption("");

				AssertCaption("");
			}
		}

		public void RenameProject()
		{
			var to = project + "2";

			AssertCaption("");

			using (ops.RenameProjectTemp(ops.GetProject(project), to))
				AssertCaption(to, "", filename);

			AssertCaption("");
		}

		public void RenameFolder()
		{
			AssertCaption("");

			using (ops.MoveFileTemp(p + @"\" + filename, p + @"\a\b\c"))
			{
				AssertCaption("a/b/c");

				using (ops.RenameFolderTemp(p + @"\a", "1"))
				{
					AssertCaption("1/b/c");

					using (ops.RenameFolderTemp(p + @"\1\b", "2"))
					{
						AssertCaption("1/2/c");

						using (ops.RenameFolderTemp(p + @"\1\2\c", "3"))
							AssertCaption("1/2/3");

						AssertCaption("1/2/c");
					}

					AssertCaption("1/b/c");
				}

				AssertCaption("a/b/c");
			}
		}

		public void MoveFolder()
		{
			AssertCaption("");

			using (ops.MoveFileTemp(p + @"\" + filename, p + @"\a\b\c"))
			{
				AssertCaption("a/b/c");

				using (ops.MoveFolderTemp(p + @"\a\b\c", p))
				{
					AssertCaption("c");

					using (ops.MoveFolderTemp(p + @"\c", p + @"\a"))
					{
						AssertCaption("a/c");

						using (ops.MoveFolderTemp(p + @"\a\c", p + @"\a\b"))
							AssertCaption("a/b/c");

						AssertCaption("a/c");
					}

					AssertCaption("c");
				}

				AssertCaption("a/b/c");
			}

			AssertCaption("");
		}

		public void RenameFile()
		{
			AssertCaption("");

			var to = "2" + filename;

			using (ops.RenameFileTemp(p + @"\" + filename, to))
				AssertCaption(project, "", to);

			AssertCaption("");
		}

		public void MoveFile()
		{
			AssertCaption("");

			using (ops.MoveFileTemp(p + @"\" + filename, p + @"\a"))
			{
				AssertCaption("a");

				using (ops.MoveFileTemp(p + @"\a\" + filename, p + @"\a\b"))
				{
					AssertCaption("a/b");

					using (ops.MoveFileTemp(p + @"\a\b\" + filename, p + @"\a\b\c"))
					{
						AssertCaption("a/b/c");

						using (ops.MoveFileTemp(p + @"\a\b\c\" + filename, p + @"\d\e\f"))
						{
							AssertCaption("d/e/f");
						}

						AssertCaption("a/b/c");
					}

					AssertCaption("a/b");
				}

				AssertCaption("a");
			}

			AssertCaption("");
		}

		public void OpenFile(Project p)
		{
			file.Close();

			file = ops.OpenFile(p, filename);
			AssertCaption("");
		}

		public void ToggleExtension()
		{
			AssertCaption("");

			using (ops.SetOptionTemp("Enabled", false))
				Assert.AreEqual(filename, file.Caption);

			AssertCaption("");
		}

		private void AssertCaption(string s)
		{
			AssertCaption(project, s, filename);
		}

		private void AssertCaption(
			string project, string folders, string filename)
		{
			string expected = project + ":" + folders + ":" + filename;
			Assert.AreEqual(expected, file.Caption);
		}
	}
}
