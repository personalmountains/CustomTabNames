using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CustomTabNames.Tests
{
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
			void Check(IDocument dd, string name)
			{
				CheckVariable<ProjectName>(dd, name);
			}

			var p = new TestProject("project", false);
			var d = new TestDocument("path", "name", p, null);
			var dNp = new TestDocument("path", "name", null, null);

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

		[TestMethod]
		public void ParentDirTests()
		{
			void Check(string path, string expected)
			{
				CheckVariable<ParentDir>(new TestDocument(path), expected);
			}

			Check("", "");

			Check("C:\\f.cpp", "");
			Check("C:\\a\\f.cpp", "a");
			Check("C:\\a\\b\\f.cpp", "b");

			Check("f.cpp", "");
			Check("a\\f.cpp", "a");
			Check("a\\b\\f.cpp", "b");

			Check("/f.cpp", "");
			Check("/a/f.cpp", "a");
			Check("/a/b/f.cpp", "b");

			Check("f.cpp", "");
			Check("a/f.cpp", "a");
			Check("a/b/f.cpp", "b");
		}

		[TestMethod]
		public void FilenameTests()
		{
			void Check(string path, string expected)
			{
				CheckVariable<Filename>(new TestDocument(path), expected);
			};

			Check("", "");

			Check("C:\\f.cpp", "f.cpp");
			Check("C:\\a\\f.cpp", "f.cpp");
			Check("C:\\a\\b\\f.cpp", "f.cpp");

			Check("C:\\", "");
			Check("C:\\a\\", "a");
			Check("C:\\a\\b\\", "b");

			Check("f.cpp", "f.cpp");
			Check("a\\f.cpp", "f.cpp");
			Check("a\\b\\f.cpp", "f.cpp");

			Check("a\\", "a");
			Check("a\\b\\", "b");

			Check("/f.cpp", "f.cpp");
			Check("/a/f.cpp", "f.cpp");
			Check("/a/b/f.cpp", "f.cpp");

			Check("f.cpp", "f.cpp");
			Check("a/f.cpp", "f.cpp");
			Check("a/b/f.cpp", "f.cpp");

			Check("/a/", "a");
			Check("/a/b/", "b");
			Check("a/", "a");
			Check("a/b/", "b");
		}

		[TestMethod]
		public void FullPathTests()
		{
			void Check(string path, string expected)
			{
				CheckVariable<FullPath>(new TestDocument(path), expected);
			};

			Check("", "");

			Check("C:\\", "C:\\");
			Check("C:\\a", "C:\\a");
			Check("C:\\a\\", "C:\\a\\");
			Check("C:\\a\\f.cpp", "C:\\a\\f.cpp");
		}

		[TestMethod]
		public void FolderTests()
		{
			void CheckFolderPath(ITreeItem item, string expected)
			{
				CheckVariable<FolderPath>(
					new TestDocument(treeItem: item), expected);
			};

			void CheckParentFolder(ITreeItem item, string expected)
			{
				CheckVariable<ParentFolder>(
					new TestDocument(treeItem: item), expected);
			};

			var root = new TestTreeItem("root");
			var root_f = new TestTreeItem("f", root);
			var root_a = new TestTreeItem("a", root, true);
			var root_a_f = new TestTreeItem("f", root_a);
			var root_a_b = new TestTreeItem("b", root_a, true);
			var root_a_b_f = new TestTreeItem("f", root_a_b);
			var root_c = new TestTreeItem("c", root, true);
			var root_c_f = new TestTreeItem("f", root_c);
			var root_c_d = new TestTreeItem("d", root_c, true);
			var root_c_d_f = new TestTreeItem("f", root_c_d);


			CheckFolderPath(null, "");
			CheckFolderPath(root, "");
			CheckFolderPath(root_f, "");
			CheckFolderPath(root_a_f, "a");
			CheckFolderPath(root_a_b_f, "a/b");
			CheckFolderPath(root_c_f, "c");
			CheckFolderPath(root_c_d_f, "c/d");

			CheckParentFolder(null, "");
			CheckParentFolder(root, "");
			CheckParentFolder(root_f, "");
			CheckParentFolder(root_a_f, "a");
			CheckParentFolder(root_a_b_f, "b");
			CheckParentFolder(root_c_f, "c");
			CheckParentFolder(root_c_d_f, "d");
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
