using System;
using System.Diagnostics;

namespace CustomTabNames
{
	public interface IOptionsBackend
	{
		bool Enabled { get; set; }
		string Template { get; set; }
		bool IgnoreBuiltinProjects { get; set; }
		bool IgnoreSingleProject { get; set; }
		bool Logging { get; set; }
		int LoggingLevel { get; set; }
		void RegisterCallback(Action<string> a);
	}


	public class Options
	{
		public struct Defaults
		{
			public const bool Enabled = true;
			public const string Template =
				"$(ProjectName ':')$(FolderPath '/')$(Filename)";
			public const bool IgnoreBuiltinProjects = true;
			public const bool IgnoreSingleProject = true;
			public const bool Logging = false;
			public const int LoggingLevel = 2;
		}

		public delegate void Handler();

		// fired when the various options change
		public event Handler EnabledChanged, TemplateChanged;
		public event Handler IgnoreBuiltinProjectsChanged;
		public event Handler IgnoreSingleProjectChanged, LoggingChanged;
		public event Handler LoggingLevelChanged;

		private readonly IOptionsBackend backend;


		public Options(IOptionsBackend b)
		{
			backend = b;
			backend.RegisterCallback(OnOptionChanged);
		}

		private void OnOptionChanged(string name)
		{
			switch (name)
			{
				case "Enabled":
					EnabledChanged?.Invoke();
					break;

				case "Template":
					TemplateChanged?.Invoke();
					break;

				case "IgnoreBuiltinProjects":
					IgnoreBuiltinProjectsChanged?.Invoke();
					break;

				case "IgnoreSingleProject":
					IgnoreSingleProjectChanged?.Invoke();
					break;

				case "Logging":
					LoggingChanged?.Invoke();
					break;

				case "LoggingLevel":
					LoggingLevelChanged?.Invoke();
					break;

				default:
					Debug.Fail("unknown option '" + name + "'");
					break;
			}
		}

		public bool Enabled
		{
			get
			{
				return backend.Enabled;
			}

			set
			{
				backend.Enabled = value;
			}
		}

		public string Template
		{
			get
			{
				return backend.Template;
			}

			set
			{
				backend.Template = value;
			}
		}

		public bool IgnoreBuiltinProjects
		{
			get
			{
				return backend.IgnoreBuiltinProjects;
			}

			set
			{
				backend.IgnoreBuiltinProjects = value;
			}
		}

		public bool IgnoreSingleProject
		{
			get
			{
				return backend.IgnoreSingleProject;
			}

			set
			{
				backend.IgnoreSingleProject = value;
			}
		}

		public bool Logging
		{
			get
			{
				return backend.Logging;
			}

			set
			{
				backend.Logging = value;
			}
		}

		public int LoggingLevel
		{
			get
			{
				return backend.LoggingLevel;
			}

			set
			{
				backend.LoggingLevel = value;
			}
		}
	}
}
