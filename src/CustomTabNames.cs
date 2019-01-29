using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using EnvDTE80;
using System.Collections.Generic;
using Task = System.Threading.Tasks.Task;
using OLE = Microsoft.VisualStudio.OLE;

namespace CustomTabNames
{
	using VariablesDictionary = Dictionary<string, Func<Document, string>>;

	[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
	[InstalledProductRegistration(Strings.ExtensionName, Strings.ExtensionDescription, Strings.ExtensionVersion)]
	[ProvideService(typeof(CustomTabNames), IsAsyncQueryable = true)]
	[ProvideAutoLoad(UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideOptionPage(typeof(Options), Strings.ExtensionName, Strings.OptionsCategory, 0, 0, true)]
	[ProvideProfileAttribute(typeof(Options), Strings.ExtensionName, Strings.OptionsCategory, 0, 0, isToolsOptionPage: true)]
	[Guid(Strings.ExtensionGuid)]
	public sealed class CustomTabNames : AsyncPackage
	{
		private DocumentManager manager;
		private VariablesDictionary variables;
		private int tries = 0;
		private Timer timer = null;
		private bool started = false;

		public CustomTabNames()
		{
			Instance = this;
		}

		public static CustomTabNames Instance { get; private set; }
		public ServiceProvider ServiceProvider { get; private set; }

		public Options Options
		{
			get
			{
				return (Options)GetDialogPage(typeof(Options));
			}
		}

		protected override async Task InitializeAsync(
			CancellationToken ct, IProgress<ServiceProgressData> p)
		{
			await this.JoinableTaskFactory.SwitchToMainThreadAsync(ct);

			var dte = Package.GetGlobalService(typeof(DTE)) as DTE2;

			this.ServiceProvider = new ServiceProvider(
				(OLE.Interop.IServiceProvider)dte);

			this.manager = new DocumentManager(dte);
			this.variables = Variables.MakeDictionary();

			this.manager.DocumentChanged += OnDocumentChanged;
			Options.EnabledChanged += OnEnabledChanged;
			Options.TemplatesChanged += OnTemplatesChanged;

			if (Options.Enabled)
			{
				Logger.Log("initialized");
				Start();
			}
			else
			{
				Logger.Log("initialized but disabled in the options");
			}
		}

		public void Start()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("starting");

			if (started)
			{
				Logger.Log("already started");
				return;
			}

			started = true;
			manager.Start();
			FixAllDocuments();
		}

		public void Stop()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("stopping");

			if (!started)
			{
				Logger.Log("already stopped");
				return;
			}

			started = false;
			manager.Stop();
			ResetAllDocuments();
		}

		private void OnTemplatesChanged(object s, EventArgs a)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (!Options.Enabled)
				return;

			Logger.Log("template options changed");
			FixAllDocuments();
		}

		private void OnEnabledChanged(object s, EventArgs a)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("enabled option changed");

			if (Options.Enabled)
				Start();
			else
				Stop();
		}

		private void OnDocumentChanged(DocumentWrapper d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (!FixCaption(d))
				Defer();
		}

		private async Task FixAllDocumentsAsync()
		{
			await JoinableTaskFactory.SwitchToMainThreadAsync();

			tries = 0;
			FixAllDocuments();
		}

		private void FixAllDocuments()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("fixing all documents");

			bool failed = false;

			manager.ForEachDocument((d) =>
			{
				if (!FixCaption(d))
					failed = true;
			});

			if (failed)
			{
				++tries;
				Logger.Log("fixing all documents failed, try {0}", tries);

				if (tries == 10)
					return;

				Defer();
			}
		}

		private void ResetAllDocuments()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("reseting all documents");

			manager.ForEachDocument((d) =>
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				d.SetCaption(d.document.Name);
			});
		}

		private void Defer()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (timer == null)
			{
				Logger.Log("deferring");
				timer = new Timer(OnTimer, null, 2000, Timeout.Infinite);
			}
			else
			{
				Logger.Log("deferring, timer already started");
				timer.Change(2000, Timeout.Infinite);
			}
		}

		private void OnTimer(object o)
		{
			// not on main thread

			timer = null;
			_ = FixAllDocumentsAsync();
		}

		private bool FixCaption(DocumentWrapper d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var caption = MakeCaption(d.document);
			if (caption != null)
				d.SetCaption(caption);

			return true;
		}

		private string MakeCaption(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			string s = Options.Template;
			var re = new Regex(@"\$\((.*?)\s*(?:'(.*?)')?\)");

			Logger.Log(
				"making caption for {0} using template {1}",
				d.FullName, s);

			while (true)
			{
				var m = re.Match(s);
				if (!m.Success)
					break;

				var name = m.Groups[1].Value;
				var text = m.Groups[2].Value;

				string replacement = "";

				if (variables.TryGetValue(name, out var v))
				{
					replacement = v(d);

					// don't append the text if the result was empty
					if (replacement != "")
						replacement += text;

					Logger.Log(
						"  . variable {0} replaced by '{1}'",
						name, replacement);
				}
				else
				{
					// not found, put the variable name to notify the user
					Logger.Log("  . variable {0} not found", name);
					replacement = name;
				}

				s = Replace(s, m, replacement);
			}

			Logger.Log("  . caption is now {0}", s);

			return s;
		}

		private string Replace(string String, Match m, string Replacement)
		{
			var ns = String.Substring(0, m.Index);
			ns += Replacement;
			ns += String.Substring(m.Index + m.Length);
			return ns;
		}
	}
}
