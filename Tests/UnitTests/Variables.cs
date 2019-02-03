using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using EnvDTE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CustomTabNames.Tests
{
	public class TestLogger : ILoggerBackend
	{
		public void Output(string s)
		{
		}
	}

	public class TestOptions : IOptionsBackend
	{
		public bool Enabled { get; set; }
		public string Template { get; set; }
		public bool IgnoreBuiltinProjects { get; set; }
		public bool IgnoreSingleProject { get; set; }
		public bool Logging { get; set; }
		public int LoggingLevel { get; set; }

		public void RegisterCallback(Action<string> a)
		{
		}
	}

	public class TestDocumentManager : DocumentManager
	{
		public override void Start()
		{
		}

		public override void Stop()
		{
		}
	}

	public class TestSolution : ISolution
	{
		public List<ITreeItem> ProjectItems
		{
			get
			{
				return null;
			}
		}

		public List<IDocument> Documents
		{
			get
			{
				return null;
			}
		}

		public bool HasSingleProject
		{
			get
			{
				return false;
			}
		}
	}

	public class TestDocument : IDocument
	{
		private string caption;

		public string Path
		{
			get
			{
				return "";
			}
		}

		public string Name
		{
			get
			{
				return "";
			}
		}

		public IProject Project
		{
			get
			{
				return null;
			}
		}

		public ITreeItem TreeItem
		{
			get
			{
				return null;
			}
		}

		public void ResetCaption()
		{
			caption = "(initial)";
		}

		public void SetCaption(string s)
		{
			caption = s;
		}
	}


	[TestClass]
	public class Global
	{
		static Main main;

		[AssemblyInitialize]
		public static void StartTests(TestContext _)
		{
			main = new Main(
				new TestOptions(), new TestLogger(),
				new TestSolution(), new TestDocumentManager());
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
			var d = new TestDocument();

			Assert.AreEqual("", Variables.Expand(d, ""));
			Assert.AreEqual("1", Variables.Expand(d, "1"));
			Assert.AreEqual("123 word", Variables.Expand(d, "123 word"));

			// broken variables
			Assert.AreEqual("123 $word", Variables.Expand(d, "123 $word"));
			Assert.AreEqual("123 $(word", Variables.Expand(d, "123 $(word"));
			Assert.AreEqual("123 $word)", Variables.Expand(d, "123 $word)"));
			Assert.AreEqual("123 (word)", Variables.Expand(d, "123 (word)"));
			Assert.AreEqual("123 $(word 't'", Variables.Expand(d, "123 $(word 't'"));
		}
	}
}
