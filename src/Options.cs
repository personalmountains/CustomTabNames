using System;
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

		public event EventHandler TemplatesChanged, EnabledChanged;

		[Category("General")]
		[DisplayName("Enabled")]
		[Description("Enabled")]
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

		[Category("General")]
		[DisplayName("Template")]
		[Description("Template")]
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
					TemplatesChanged?.Invoke(this, EventArgs.Empty);
				}
			}
		}

		[Category("General")]
		[DisplayName("Logging")]
		[Description("Logging")]
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
