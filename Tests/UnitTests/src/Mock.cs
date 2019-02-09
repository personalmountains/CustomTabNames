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
		private Action<string> callback;

		private bool enabled = Options.Defaults.Enabled;
		private string template = Options.Defaults.Template;
		private bool ignoreBuiltinProjects =
			Options.Defaults.IgnoreBuiltinProjects;
		private bool ignoreSingleProject =
			Options.Defaults.IgnoreSingleProject;
		private bool logging = Options.Defaults.Logging;
		private int loggingLevel = Options.Defaults.LoggingLevel;

		public void RegisterCallback(Action<string> a)
		{
			callback = a;
		}

		public bool Enabled
		{
			get
			{
				return enabled;
			}

			set
			{
				if (enabled != value)
				{
					enabled = value;
					callback?.Invoke("Enabled");
				}
			}
		}

		public string Template
		{
			get
			{
				return template;
			}

			set
			{
				var v = (value.Length == 0 ? Options.Defaults.Template : value);

				if (template != v)
				{
					template = v;
					callback?.Invoke("Template");
				}
			}
		}

		public bool IgnoreBuiltinProjects
		{
			get
			{
				return ignoreBuiltinProjects;
			}

			set
			{
				if (ignoreBuiltinProjects != value)
				{
					ignoreBuiltinProjects = value;
					callback?.Invoke("IgnoreBuiltinProjects");
				}
			}
		}

		public bool IgnoreSingleProject
		{
			get
			{
				return ignoreSingleProject;
			}

			set
			{
				if (ignoreSingleProject != value)
				{
					ignoreSingleProject = value;
					callback?.Invoke("IgnoreSingleProject");
				}
			}
		}

		public bool Logging
		{
			get
			{
				return logging;
			}

			set
			{
				if (logging != value)
				{
					logging = value;
					callback?.Invoke("Logging");
				}
			}
		}

		public int LoggingLevel
		{
			get
			{
				return loggingLevel;
			}

			set
			{
				if (value < 0)
					value = 0;
				else if (value > 4)
					value = 4;

				if (loggingLevel != value)
				{
					loggingLevel = value;
					callback?.Invoke("LoggingLevel");
				}
			}
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
				return new List<ITreeItem>();
			}
		}

		public List<IDocument> Documents
		{
			get
			{
				return new List<IDocument>();
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
			string path = "", string name = "",
			IProject project = null, ITreeItem treeItem = null)
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

	public class TestTreeItem : ITreeItem
	{
		public TestTreeItem(
			string name = "", ITreeItem parent = null,
			bool isFolder = false)
		{
			Name = name;
			Parent = parent;
			IsFolder = isFolder;
		}

		public string Name { get; }
		public ITreeItem Parent { get; }
		public bool IsFolder { get; }
		public string DebugName { get; }
	}
}
