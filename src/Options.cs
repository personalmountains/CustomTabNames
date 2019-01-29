﻿using System;
using System.ComponentModel;
using Microsoft.VisualStudio.Shell;

namespace CustomTabNames
{
	[System.ComponentModel.DesignerCategory("")]
	public class Options : DialogPage
	{
		private const bool defaultEnabled = true;

		private const string defaultTemplate =
			"$(ProjectName ':')$(ParentDir)$(Filename)";

		private const bool defaultLogging = false;

		private bool enabled = defaultEnabled;
		private string template = defaultTemplate;
		private bool logging = defaultLogging;

		public event EventHandler TemplateChanged, EnabledChanged;

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionEnabled)]
		[Description(Strings.OptionEnabledDescription)]
		[DefaultValue(defaultEnabled)]
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
					EnabledChanged?.Invoke(this, EventArgs.Empty);
				}
			}
		}

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionTemplate)]
		[Description(Strings.OptionTemplateDescription)]
		[DefaultValue(defaultTemplate)]
		public string Template
		{
			get
			{
				return template;
			}

			set
			{
				var v = (value == "" ? defaultTemplate : value);

				if (template != v)
				{
					template = value;
					TemplateChanged?.Invoke(this, EventArgs.Empty);
				}
			}
		}

		[Category(Strings.OptionsCategory)]
		[DisplayName(Strings.OptionLogging)]
		[Description(Strings.OptionLoggingDescription)]
		[DefaultValue(defaultLogging)]
		public bool Logging
		{
			get
			{
				return logging;
			}

			set
			{
				logging = value;
			}
		}

	}
}
