using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace CustomTabNames
{
	// stores all options and provides a DialogPage that can be displayed in
	// the options dialog
	//
	// DesignerCategory("") is because Visual Studio keeps trying to show a
	// designer window for this on double-click, but it gives errors and is
	// useless in any case; this forces a code view
	//
	// ComVisible(true) allows these options to be accessed by automation
	// through DTE.Properties, which is used by tests
	//
	[System.ComponentModel.DesignerCategory("")]
	[ComVisible(true)]
	public class Options : DialogPage
	{
		struct Defaults
		{
			public const bool Enabled = true;
			public const string Template =
				"$(ProjectName ':')$(FilterPath)$(Filename)";
			public const bool IgnoreBuiltinProjects = true;
			public const bool IgnoreSingleProject = true;
			public const bool Logging = false;
			public const int LoggingLevel = 2;
		}

		private bool enabled = Defaults.Enabled;
		private string template = Defaults.Template;
		private bool ignoreBuiltinProjects = Defaults.IgnoreBuiltinProjects;
		private bool ignoreSingleProject = Defaults.IgnoreSingleProject;
		private bool logging = Defaults.Logging;
		private int loggingLevel = Defaults.LoggingLevel;


		public delegate void Handler();

		// fired when the various options change
		public event Handler EnabledChanged, TemplateChanged;
		public event Handler IgnoreBuiltinProjectsChanged;
		public event Handler IgnoreSingleProjectChanged, LoggingChanged;
		public event Handler LoggingLevelChanged;


		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionEnabled)]
		[Description(Strings.OptionEnabledDescription)]
		[DefaultValue(Defaults.Enabled)]
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
					EnabledChanged?.Invoke();
				}
			}
		}

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionTemplate)]
		[Description(Strings.OptionTemplateDescription)]
		[DefaultValue(Defaults.Template)]
		public string Template
		{
			get
			{
				return template;
			}

			set
			{
				var v = (value.Length == 0 ? Defaults.Template : value);

				if (template != v)
				{
					template = v;
					TemplateChanged?.Invoke();
				}
			}
		}

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionIgnoreBuiltinProjects)]
		[Description(Strings.OptionIgnoreBuiltinProjectsDescription)]
		[DefaultValue(Defaults.IgnoreBuiltinProjects)]
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
					ignoreBuiltinProjects  = value;
					IgnoreBuiltinProjectsChanged?.Invoke();
				}
			}
		}

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionIgnoreSingleProject)]
		[Description(Strings.OptionIgnoreSingleProjectDescription)]
		[DefaultValue(Defaults.IgnoreSingleProject)]
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
					IgnoreSingleProjectChanged?.Invoke();
				}
			}
		}

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionLogging)]
		[Description(Strings.OptionLoggingDescription)]
		[DefaultValue(Defaults.Logging)]
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
					LoggingChanged?.Invoke();
				}
			}
		}

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionLoggingLevel)]
		[Description(Strings.OptionLoggingLevelDescription)]
		[DefaultValue(Defaults.LoggingLevel)]
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
					LoggingLevelChanged?.Invoke();
				}
			}
		}
	}
}
