using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CustomTabNames.Tests
{
	[TestClass]
	public class CsTests
	{
		public static Tests tests;

		[ClassInitialize]
		public static void Init(TestContext _)
		{
			tests = new Tests("cs", "f.cs");
		}

		[TestMethod]
		public void AddRemoveProjectsIgnoreSingle()
		{
			tests.AddRemoveProjectsIgnoreSingle(Global.CPP);
		}

		[TestMethod]
		public void AddRemoveProjectsDontIgnoreSingle()
		{
			tests.AddRemoveProjectsDontIgnoreSingle(Global.CPP);
		}

		[TestMethod]
		public void RenameProject()
		{
			tests.RenameProject();
		}

		[TestMethod]
		public void RenameFolder()
		{
			tests.RenameFolder();
		}

		[TestMethod]
		public void MoveFolder()
		{
			tests.MoveFolder();
		}

		[TestMethod]
		public void RenameFile()
		{
			tests.RenameFile();
		}

		[TestMethod]
		public void MoveFile()
		{
			tests.MoveFile();
		}

		[TestMethod]
		public void OpenFile()
		{
			tests.OpenFile(Global.CS);
		}

		[TestMethod]
		public void ToggleExtension()
		{
			tests.ToggleExtension();
		}
	}
}
