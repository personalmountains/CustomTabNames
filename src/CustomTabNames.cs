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
		// number of times the defer timer was fired, see Defer()
		private int tries = 0;
		private const int MaxTries = 5;
		private const int TryInterval = 2000;

		// started by Defer() when at least when document didn't have a frame,
		// see Defer()
		private Timer timer = null;
		private readonly object timerLock = new object();

		// whether Start() has already been called
		private bool started = false;


		// this instance
		public static CustomTabNames Instance { get; private set; }

		public DTE2 DTE { get; private set; }
		public ServiceProvider ServiceProvider { get; private set; }
		public DocumentManager DocumentManager { get; private set; }

		// options
		public Options Options { get; private set; }


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
			Logger.Log("starting");

			if (started)
			{
				// shouldn't happen
				Logger.Log("already started");
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
			Logger.Log("stopping");

			if (!started)
			{
				Logger.Log("already stopped");
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

			if (!FixCaption(d))
				Defer();
		}

		// called from OnTimer() below because it's not on the main thread
		//
		private async Task FixAllDocumentsAsync()
		{
			await JoinableTaskFactory.SwitchToMainThreadAsync();
			FixAllDocuments();
		}

		// walks through all opened documents and tries to set the caption for
		// each of them
		//
		// this may fail for certain documents if they're open but don't have a
		// frame yet, which seemingly happens when loading a project; in that
		// case, Defer() is called, which starts a timer to try again later
		//
		private void FixAllDocuments()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("fixing all documents");

			bool failed = false;

			DocumentManager.ForEachDocument((d) =>
			{
				if (!FixCaption(d))
					failed = true;
			});

			if (failed)
			{
				// at least one document failed

				// don't try again indefinitely
				++tries;
				Logger.Log("fixing all documents failed, try {0}", tries);

				if (tries >= MaxTries)
				{
					// tried too many times
					Logger.Log("exceeded {0} tries, bailing out", MaxTries);
					tries = 0;
					return;
				}

				// try again later
				Defer();
			}
			else
			{
				// succeeded, reset the tries for next time, although this
				// doesn't seem to happen again once all the documents have
				// been loaded
				tries = 0;
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

		// called by FixAllDocuments() when at least one document failed; starts
		// or updates a timer so they can be tried again
		//
		private void Defer()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			lock (timerLock)
			{
				if (timer == null)
				{
					Logger.Log("deferring");

					timer = new Timer(
						OnTimer, null, TryInterval, Timeout.Infinite);
				}
				else
				{
					// this shouldn't happen, Defer() is only called once all
					// documents have been processed
					Logger.Log("deferring, timer already started");
					timer.Change(TryInterval, Timeout.Infinite);
				}
			}
		}

		// fired by the timer set in Defer(), tries to fix all documents again
		//
		private void OnTimer(object o)
		{
			// careful: not on main thread

			lock (timerLock)
			{
				timer = null;
			}

			_ = FixAllDocumentsAsync();
		}

		// called on each document by FixAllDocuments() and on documents given
		// by the DocumentChanged event from the DocumentManager
		//
		// creates a caption and tries to set it on the document
		//
		// may fail if the document doesn't have a frame yet
		//
		private bool FixCaption(DocumentWrapper d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var caption = Variables.Expand(d.Document, Options.Template);
			return d.SetCaption(caption);
		}
	}
}
