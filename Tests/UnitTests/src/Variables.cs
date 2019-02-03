using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using EnvDTE;
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

			Options.Enabled = false;
			Main = new Main(Options, Logger, Solution, DocumentManager);
		}

		[AssemblyCleanup]
		public static void StopTests()
		{
		}
	}


	[TestClass]
	public class VariableTests
	{
		[TestMethod]
		public void ExpandWithoutVariables()
		{
			CheckSame("");
			CheckSame("1");
			CheckSame("123 word");

			// broken variables
			CheckSame("$word");
			CheckSame("$(word");
			CheckSame("$word)");
			CheckSame("(word)");
			CheckSame("$(word 't'");
		}

		[TestMethod]
		public void ExpandBadNames()
		{
			// bad names
			CheckSame("$(_word)");
			CheckSame("$(1word)");
			CheckSame("$(-word)");

			CheckSame("$(wo_rd)");
			CheckSame("$(wo1rd)");
			CheckSame("$(woo-rd)");

			CheckSame("$(word_)");
			CheckSame("$(word1)");
			CheckSame("$(word-)");

			CheckSame("$(_word 'text')");
			CheckSame("$(1word 'text')");
			CheckSame("$(-word 'text')");

			CheckSame("$($word))");
		}

		[TestMethod]
		public void ExpandUnknownVariable()
		{
			Check("$(bad)", "bad");
			Check("$(bad 'text')", "bad");

			Check("$($(bad))", "bad");
			Check("$($(bad 'text'))", "bad");
			Check("$($(bad) 'text')", "bad");

			Check("$($(bad)", "$(bad");
			Check("$$(bad))", "$bad)");
		}

		[TestMethod]
		public void ProjectNameTests()
		{
			var p = new TestProject("project", false);
			var d = new TestDocument("path", "name", p, null);
			var dNp = new TestDocument("path", "name", null, null);

			Action<IDocument, string> Check = CheckVariable<ProjectName>;

			// IgnoreSingleProject true
			Global.Options.IgnoreSingleProject = true;
			Global.Solution.HasSingleProject = true;
			Check(d, "");
			Check(dNp, "");

			Global.Solution.HasSingleProject = false;
			Check(d, "project");
			Check(dNp, "");

			// IgnoreSingleProject false
			Global.Options.IgnoreSingleProject = false;
			Global.Solution.HasSingleProject = true;
			Check(d, "project");
			Check(dNp, "");

			Global.Solution.HasSingleProject = false;
			Check(d, "project");
			Check(dNp, "");
		}

		private void CheckSame(string s)
		{
			TestDocument d = new TestDocument();
			Assert.AreEqual(s, Variables.Expand(d, s));
		}

		private void Check(string s, string expected)
		{
			TestDocument d = new TestDocument();
			Assert.AreEqual(expected, Variables.Expand(d, s));
		}

		private void CheckVariable<Variable>(IDocument d, string expected)
			where Variable : IVariable, new()
		{
			Assert.AreEqual(expected, new Variable().Expand(d));
		}
	}
}
