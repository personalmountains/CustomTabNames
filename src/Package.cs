using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

// the Package class is the package; it starts the DocumentManager,
// waits for events and calls FixCaption() on documents
//
// the DocumentManager and class registers events like opening documents and
// notifies Package that a document caption needs fixing
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
	[ProvideService(typeof(Package), IsAsyncQueryable = true)]
	[ProvideAutoLoad(UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideOptionPage(typeof(Options), Strings.ExtensionName, Strings.OptionsCategory, 0, 0, true)]
	[ProvideProfile(typeof(Options), Strings.ExtensionName, Strings.OptionsCategory, 0, 0, isToolsOptionPage: true)]
	[Guid(Strings.ExtensionGuid)]
	public sealed class Package : AsyncPackage, IDisposable
	{
		// this instance
		public static Package Instance { get; private set; }

		public DocumentManager DocumentManager { get; private set; }

		// options
		public Options Options { get; private set; }

		// whether Start() has already been called
		private bool started = false;

		// services
		IVsSolution solution = null;
		IVsRunningDocumentTable rdt = null;
		IVsRunningDocumentTable4 rdt4 = null;
		IVsOutputWindow outputWindow = null;

		// this timer is used when fixing documents fails because frames aren't
		// ready yet; this happens when the projects are in the process of
		// loading
		private readonly MainThreadTimer timer = new MainThreadTimer();
		private int failures = 0;
		private const int FailureDelay = 2000;
		private const int MaxFailures = 5;

		public Package()
		{
			Instance = this;
		}

		public void Dispose()
		{
			timer.Dispose();
		}

		protected override async Task InitializeAsync(
			CancellationToken ct, IProgress<ServiceProgressData> p)
		{
			await JoinableTaskFactory.SwitchToMainThreadAsync(ct);

			if (Solution == null || RDT == null)
			{
				Logger.Error("bailing out");
				return;
			}

			Options = (Options)GetDialogPage(typeof(Options));
			DocumentManager = new DocumentManager();

			DocumentManager.DocumentChanged += OnDocumentChanged;
			DocumentManager.ContainersChanged += OnContainersChanged;

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

		public IVsSolution Solution
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (solution == null)
				{
					solution = GetService(typeof(SVsSolution)) as IVsSolution;
					if (solution == null)
						Logger.Error("failed to get IVsSolution");
				}

				return solution;
			}
		}

		public IVsRunningDocumentTable RDT
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (rdt == null)
				{
					rdt = GetService(typeof(SVsRunningDocumentTable))
						as IVsRunningDocumentTable;

					if (rdt == null)
						Logger.Error("can't get IVsRunningDocumentTable");
				}

				return rdt;
			}
		}

		public IVsRunningDocumentTable4 RDT4
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (rdt4 == null)
				{
					rdt4 = GetService(typeof(SVsRunningDocumentTable))
						as IVsRunningDocumentTable4;

					if (rdt4 == null)
						Logger.Error("can't get IVsRunningDocumentTable4");
				}

				return rdt4;
			}
		}

		public IVsOutputWindow OutputWindow
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (outputWindow == null)
				{
					outputWindow = GetService(typeof(SVsOutputWindow))
						as IVsOutputWindow;

					if (outputWindow == null)
						Logger.Error("can't get IVsOutputWindow");
				}

				return outputWindow;
			}
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

		private void OnContainersChanged()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("containers changed");
			FixAllDocuments();
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
				Utilities.ForEachDocument((d) =>
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

			Utilities.ForEachDocument((d) =>
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
			d.SetCaption(Variables.Expand(d.Document, Options.Template));
		}
	}
}
