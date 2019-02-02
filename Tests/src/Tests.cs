using Microsoft.VisualStudio.TestTools.UnitTesting;
using EnvDTE;

namespace CustomTabNames.Tests
{
	[TestClass]
	public class UnitTest1
	{
		private const string SolutionPath =
			"c:\\dev\\projects\\CustomTabNames\\Tests\\data\\test.sln";

		private static VS vs;
		private static Operations ops;
		private static Project cpp, cs;

		[ClassInitialize]
		public static void StartTests(TestContext _)
		{
			vs = new VS(SolutionPath);
			ops = vs.Operations;

			cpp = ops.FindProject("cpp");
			cs = ops.FindProject("cs");
		}

		[TestCleanup]
		public void StopTests()
		{
			vs.Dispose();
		}

		[TestMethod]
		public void TestMethod1()
		{
			var w = ops.OpenFile(cpp, "f.cpp");

			ops.MoveFile(
				@"test\cpp\f.cpp",
				@"test\cpp\1\1.1\1.1.1");

			Assert.AreEqual("cpp:1/1.1/1.1.1/f.cpp", w.Caption);
		}
	}
}
