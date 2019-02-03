using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
		public bool Enabled { get; set; } = Options.Defaults.Enabled;
		public string Template { get; set; } = Options.Defaults.Template;
		public bool IgnoreBuiltinProjects { get; set; } = Options.Defaults.IgnoreBuiltinProjects;
		public bool IgnoreSingleProject { get; set; } = Options.Defaults.IgnoreSingleProject;
		public bool Logging { get; set; } = Options.Defaults.Logging;
		public int LoggingLevel { get; set; } = Options.Defaults.LoggingLevel;

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

		public bool HasSingleProject { get; set; }
	}

	public class TestProject : IProject
	{
		private readonly string name;
		private readonly bool isBuiltIn;

		public TestProject(string name, bool isBuiltIn)
		{
			this.name = name;
			this.isBuiltIn = isBuiltIn;
		}

		public string Name
		{
			get
			{
				return name;
			}
		}

		public bool IsBuiltIn
		{
			get
			{
				return isBuiltIn;
			}
		}
	}

	public class TestDocument : IDocument
	{
		private readonly string path, name;
		private string caption;
		private readonly IProject project;
		private readonly ITreeItem treeItem;

		public TestDocument()
		{
		}

		public TestDocument(
			string path, string name, IProject project, ITreeItem treeItem)
		{
			this.path = path;
			this.name = name;
			this.project = project;
			this.treeItem = treeItem;

			ResetCaption();
		}

		public string Path
		{
			get
			{
				return path;
			}
		}

		public string Name
		{
			get
			{
				return name;
			}
		}

		public IProject Project
		{
			get
			{
				return project;
			}
		}

		public ITreeItem TreeItem
		{
			get
			{
				return treeItem;
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
}
