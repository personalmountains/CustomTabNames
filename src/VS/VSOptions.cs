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
	[DesignerCategory("")]
	[ComVisible(true)]
	public class VSOptions : DialogPage, IOptionsBackend
	{
		private Action<string> callback;
		private bool enabled = Options.Defaults.Enabled;
		private string template = Options.Defaults.Template;
		private bool ignoreBuiltinProjects = Options.Defaults.IgnoreBuiltinProjects;
		private bool ignoreSingleProject = Options.Defaults.IgnoreSingleProject;
		private bool logging = Options.Defaults.Logging;
		private int loggingLevel = Options.Defaults.LoggingLevel;

		public void RegisterCallback(Action<string> a)
		{
			callback = a;
		}

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionEnabled)]
		[Description(Strings.OptionEnabledDescription)]
		[DefaultValue(Options.Defaults.Enabled)]
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

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionTemplate)]
		[Description(Strings.OptionTemplateDescription)]
		[DefaultValue(Options.Defaults.Template)]
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

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionIgnoreBuiltinProjects)]
		[Description(Strings.OptionIgnoreBuiltinProjectsDescription)]
		[DefaultValue(Options.Defaults.IgnoreBuiltinProjects)]
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

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionIgnoreSingleProject)]
		[Description(Strings.OptionIgnoreSingleProjectDescription)]
		[DefaultValue(Options.Defaults.IgnoreSingleProject)]
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


		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionLogging)]
		[Description(Strings.OptionLoggingDescription)]
		[DefaultValue(Options.Defaults.Logging)]
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

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionLoggingLevel)]
		[Description(Strings.OptionLoggingLevelDescription)]
		[DefaultValue(Options.Defaults.LoggingLevel)]
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
}
