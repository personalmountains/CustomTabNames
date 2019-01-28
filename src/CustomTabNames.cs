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
	[InstalledProductRegistration("#100", "#101", "1.0")]
	[ProvideService(typeof(CustomTabNames), IsAsyncQueryable = true)]
	[ProvideAutoLoad(UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideOptionPage(typeof(Options), "CustomTabNames", "General", 0, 0, true)]
	[ProvideProfileAttribute(typeof(Options), "CustomTabNames", "General", 200, 201, isToolsOptionPage: true, DescriptionResourceID = 202)]
	[Guid(CustomTabNames.Guid)]
	public sealed class CustomTabNames : AsyncPackage
	{
		public const string Guid = "BEE6C21E-FBF8-49B1-A0F8-89D7DFA732EE";

		private DTE2 dte;
		private ServiceProvider sp;
		private DocumentEvents docEvents;
		private WindowEvents winEvents;

		private VariablesDictionary variables;
		private int tries = 0;
		private Timer timer = null;
		private bool started = false;

		private static CustomTabNames instance = null;

		public CustomTabNames()
		{
			instance = this;
		}

		public static CustomTabNames Instance
		{
			get
			{
				return instance;
			}
		}

		public ServiceProvider ServiceProvider
		{
			get
			{
				return sp;
			}
		}

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

			this.dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
			this.sp = new ServiceProvider((OLE.Interop.IServiceProvider)dte);
			this.docEvents = dte.Events.DocumentEvents;
			this.winEvents = dte.Events.WindowEvents;
			this.variables = Variables.MakeDictionary();

			// this never gets removed
			Options.EnabledChanged += OnEnabledChanged;

			Logger.Log("initialized");
			Start();
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
			SetEvents(true);
			FixAllDocumentsAsync();
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
			SetEvents(false);

			FixAllDocuments();
		}

		void SetEvents(bool add)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (add)
			{
				Logger.Log("adding events");
				docEvents.DocumentOpened += OnDocumentOpened;
				winEvents.WindowCreated += OnWindowCreated;
				winEvents.WindowActivated += OnWindowActivated;
				Options.TemplatesChanged += OnTemplatesChanged;
			}
			else
			{
				Logger.Log("removing events");
				docEvents.DocumentOpened -= OnDocumentOpened;
				winEvents.WindowCreated -= OnWindowCreated;
				winEvents.WindowActivated -= OnWindowActivated;
				Options.TemplatesChanged -= OnTemplatesChanged;
			}
		}

		private void OnDocumentOpened(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("document opened: {0}", d.FullName);

			if (!FixCaption(d))
				Defer();
		}

		private void OnWindowCreated(Window w)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("window created: {0}", w.Caption);

			if (w.Document == null)
			{
				Logger.Log("not a document");
				return;
			}

			if (!FixCaption(w.Document))
				Defer();
		}

		private void OnWindowActivated(Window w, Window lost)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("window activated: {0}", w.Caption);

			if (w.Document == null)
			{
				Logger.Log("not a document");
				return;
			}

			if (!FixCaption(w.Document))
				Defer();
		}

		private void OnTemplatesChanged(object s, EventArgs a)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("template options changed");

			FixAllDocumentsAsync();
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

		private void FixAllDocumentsAsync()
		{
			_ = FixAllDocumentsImplAsync();
		}

		private async Task FixAllDocumentsImplAsync()
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

			foreach (Document d in dte.Documents)
			{
				if (!FixCaption(d))
					failed = true;
			}

			if (failed)
			{
				++tries;
				Logger.Log("fixing all documents failed, try {0}", tries);

				if (tries == 10)
					return;

				Defer();
			}
		}

		private bool FixCaption(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var f = DocumentFrame(d);
			if (f == null)
			{
				Logger.Log("document {0} has no frame", d.FullName);
				return false;
			}

			if (Options.Enabled)
			{
				var caption = MakeCaption(d);
				if (caption != null)
					SetCaption(f, caption);
			}
			else
			{
				SetCaption(f, d.Name);
			}

			return true;
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
			FixAllDocumentsAsync();
		}

		private IVsWindowFrame DocumentFrame(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			VsShellUtilities.IsDocumentOpen(
				sp, d.FullName, VSConstants.LOGVIEWID.Primary_guid,
				out IVsUIHierarchy h, out uint id,
				out IVsWindowFrame frame);

			return frame;
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

		private void SetCaption(IVsWindowFrame f, string s)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			f.SetProperty((int)VsFramePropID.EditorCaption, null);
			f.SetProperty((int)VsFramePropID.OwnerCaption, s);
		}
	}
}
