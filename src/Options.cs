using System.ComponentModel;
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
	[System.ComponentModel.DesignerCategory("")]
	public class Options : DialogPage
	{
		struct Defaults
		{
			public const bool Enabled = true;
			public const string Template =
				"$(ProjectName ':')$(ParentDir)$(Filename)";
			public const bool Logging = false;
		}

		private bool enabled    = Defaults.Enabled;
		private string template = Defaults.Template;
		private bool logging    = Defaults.Logging;


		public delegate void Handler();

		// fired when the various options change
		public event Handler EnabledChanged, TemplateChanged, LoggingChanged;


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
				var v = (value == "" ? Defaults.Template : value);

				if (template != v)
				{
					template = value;
					TemplateChanged?.Invoke();
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

	}
}
