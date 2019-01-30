using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using EnvDTE80;
using Task = System.Threading.Tasks.Task;
using OLE = Microsoft.VisualStudio.OLE;

// the CustomTabNames class is the package; it starts the DocumentManager,
// waits for events and calls FixCaption() on documents
//
// the DocumentManager class registers events like opening documents and
// notifies CustomTabNames that a document caption needs fixing
//
// Logger has simple static functions to log strings to the output window and
// Strings has most of the localizable strings
//
// Options has all the options and acts as a DialogPage that can be shown in the
// options dialog
//
// Variables has all available variables and can expand them based on a template
// and a document

namespace CustomTabNames
{
	sealed class MainThreadTimer
	{
		private Timer t = null;

		public void Start(int ms, Action a)
		{
			if (t == null)
				t = new Timer(OnTimer, a, ms, Timeout.Infinite);
			else
				t.Change(ms, Timeout.Infinite);
		}

		private void OnTimer(object a)
		{
			_ = OnMainThreadAsync((Action)a);
		}

		private async Task OnMainThreadAsync(Action a)
		{
			await CustomTabNames.Instance
				.JoinableTaskFactory.SwitchToMainThreadAsync();

			a();
		}
	}

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
		// this instance
		public static CustomTabNames Instance { get; private set; }

		public DTE2 DTE { get; private set; }
		public ServiceProvider ServiceProvider { get; private set; }
		public DocumentManager DocumentManager { get; private set; }

		// options
		public Options Options { get; private set; }

		// whether Start() has already been called
		private bool started = false;

		private readonly MainThreadTimer timer = new MainThreadTimer();
		private int failures = 0;
		private const int FailureDelay = 2000;
		private const int MaxFailures = 5;

		public CustomTabNames()
		{
			Instance = this;
		}

		protected override async Task InitializeAsync(
			CancellationToken ct, IProgress<ServiceProgressData> p)
		{
			await this.JoinableTaskFactory.SwitchToMainThreadAsync(ct);

			this.DTE = Package.GetGlobalService(typeof(DTE)) as DTE2;

			this.ServiceProvider = new ServiceProvider(
				(OLE.Interop.IServiceProvider)this.DTE);

			Options = (Options)GetDialogPage(typeof(Options));
			this.DocumentManager = new DocumentManager(this.DTE);

			// fired when documents or windows are opened; fixes the caption
			// for that particular document
			this.DocumentManager.DocumentChanged += OnDocumentChanged;

			Options.EnabledChanged += OnEnabledChanged;
			Options.TemplateChanged += OnTemplateChanged;
			Options.IgnoreBuiltinProjectsChanged += OnIgnoreBuiltinProjectsChanged;
			Options.IgnoreSingleProjectChanged += OnIgnoreSingleProjectChanged;

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

		// starts the manager (which register handlers for documents) and
		// do a first pass on all opened documents
		//
		public void Start()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Trace("starting");

			if (started)
			{
				// shouldn't happen
				Logger.Error("already started");
				return;
			}

			started = true;
			DocumentManager.Start();
			FixAllDocuments();
		}

		// stops the manager (kills events for documents) and resets all the
		// captions to their original value
		//
		public void Stop()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Trace("stopping");

			if (!started)
			{
				Logger.Error("already stopped");
				return;
			}

			started = false;
			DocumentManager.Stop();
			ResetAllDocuments();
		}

		// fired when the template option changed, fixes all currently opened
		// documents
		//
		private void OnTemplateChanged()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (!Options.Enabled)
				return;

			Logger.Log("template option changed");
			FixAllDocuments();
		}

		// fired when the enabled option changed, either starts or stops
		//
		private void OnEnabledChanged()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("enabled option changed");

			if (Options.Enabled)
				Start();
			else
				Stop();
		}

		// fired when the ignore builtin projects option changed, fixes all
		// currently opened documents
		//
		private void OnIgnoreBuiltinProjectsChanged()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (!Options.Enabled)
				return;

			Logger.Log("ignore builtin projects option changed");
			FixAllDocuments();
		}

		// fired when the ignore single project option changed, fixes all
		// currently opened documents
		//
		private void OnIgnoreSingleProjectChanged()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (!Options.Enabled)
				return;

			Logger.Log("ignore single project option changed");
			FixAllDocuments();
		}

		// fired when a document or window has been opened
		//
		private void OnDocumentChanged(DocumentWrapper d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			FixCaption(d);
		}

		// walks through all opened documents and tries to set the caption for
		// each of them
		//
		private void FixAllDocuments()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("fixing all documents");

			try
			{
				DocumentManager.ForEachDocument((d) =>
				{
					FixCaption(d);
				});
			}
			catch (COMException e)
			{
				// not sure why this happens, but dte.Documents will sometimes
				// throw a COMException with E_FAIL
				//
				// this starts a timer and tries again after a few seconds, but
				// only MaxFailures times
				//
				// if it's still broken after that, tabs that are already
				// opened won't be processed, unless they get handled by the
				// RDT events

				++failures;

				Logger.Error("ForEachDocument failed, " + e.Message);

				if (failures < MaxFailures)
				{
					Logger.Error("trying again in {0} ms", FailureDelay);
					timer.Start(FailureDelay, FixAllDocuments);
				}
				else
				{
					Logger.Error("bailing out");
				}
			}
		}

		// walks through all opened documents and resets the caption for each
		// of them; failure is ignored
		//
		private void ResetAllDocuments()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("reseting all documents");

			DocumentManager.ForEachDocument((d) =>
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				d.ResetCaption();
			});
		}

		// called on each document by FixAllDocuments() and on documents given
		// by the DocumentChanged event from the DocumentManager
		//
		// creates a caption and tries to set it on the document
		//
		// may fail if the document doesn't have a frame yet
		//
		private void FixCaption(DocumentWrapper d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var c = Variables.Expand(d.Document, Options.Template);
			Logger.Log("caption for {0} is {1}", d.Document.FullName, c);

			d.SetCaption(c);
		}
	}
}
