﻿using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using EnvDTE80;
using System.Collections.Generic;
using Task = System.Threading.Tasks.Task;
using OLE = Microsoft.VisualStudio.OLE;

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
		// handles document events and allows iterating all opened documents
		private DocumentManager manager;

		// number of times the defer timer was fired, see Defer()
		private int tries = 0;
		private const int MaxTries = 5;
		private const int TryInterval = 2000;

		// started by Defer() when at least when document didn't have a frame,
		// see Defer()
		private Timer timer = null;

		// whether Start() has already been called
		private bool started = false;


		// this instance
		public static CustomTabNames Instance { get; private set; }

		// used by both logging and document manager
		public ServiceProvider ServiceProvider { get; private set; }

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

			var dte = Package.GetGlobalService(typeof(DTE)) as DTE2;

			this.ServiceProvider = new ServiceProvider(
				(OLE.Interop.IServiceProvider)dte);

			this.manager = new DocumentManager(dte);

			// fired when documents or windows are opened
			this.manager.DocumentChanged += OnDocumentChanged;

			Options = (Options)GetDialogPage(typeof(Options));
			Options.EnabledChanged += OnEnabledChanged;
			Options.TemplateChanged += OnTemplateChanged;

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
			manager.Start();
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
			manager.Stop();
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

			Logger.Log("template options changed");
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

			manager.ForEachDocument((d) =>
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

			manager.ForEachDocument((d) =>
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

			lock (timer)
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

			lock (timer)
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

			var caption = Variables.MakeCaption(d.Document, Options.Template);
			return d.SetCaption(caption);
		}
	}
}
