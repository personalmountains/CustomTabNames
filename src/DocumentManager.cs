using System;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using EnvDTE80;

namespace CustomTabNames
{
	// wraps a Document and a IVsWindowFrame
	//
	// the Document is used to feed information to variables, like path and
	// filters; the IVsWindowFrame is required to set the actual caption
	//
	// the frame might be null in cases when the document is opened, but no
	// window has been created yet
	//
	public sealed class DocumentWrapper
	{
		public Document Document { get; private set; }
		private IVsWindowFrame Frame { get; set; }

		public DocumentWrapper(Document d, IVsWindowFrame f)
		{
			Document = d;
			Frame = f;
		}

		// sets the caption of this document to the given string
		//
		public bool SetCaption(string s)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// fail without frame
			if (Frame == null)
				return false;

			// the visible caption is made of a combination of the EditorCaption
			// and OwnerCaption; setting the EditorCaption to null makes sure
			// the caption can be controlled uniquely by OwnerCaption
			Frame.SetProperty((int)VsFramePropID.EditorCaption, null);
			Frame.SetProperty((int)VsFramePropID.OwnerCaption, s);

			return true;
		}

		// resets the caption of this document to the default value
		//
		public void ResetCaption()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// todo: it'd be nice to set the caption back to a default value
			// instead of hardcoding the name, but there doesn't seem to be a
			// way to do that
			SetCaption(Document.Name);
		}
	}


	// manages the various events for opening documents and windows, and fires
	// DocumentChanged when they do
	//
	public sealed class DocumentManager
	{
		private readonly DTE2 dte;

		// these two need to be referenced here so they don't get collected, in
		// which case none of the events registered work
		// see https://stackoverflow.com/a/3899794/4885801
		private readonly DocumentEvents docEvents;
		private readonly WindowEvents winEvents;

		// fired every time a document changes in a way that may require
		// fixing the caption
		public delegate void DocumentChangedHandler(DocumentWrapper d);
		public event DocumentChangedHandler DocumentChanged;

		public DocumentManager(DTE2 dte)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			this.dte = dte;
			this.docEvents = dte.Events.DocumentEvents;
			this.winEvents = dte.Events.WindowEvents;
		}

		// starts the manager
		//
		public void Start()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			SetEvents(true);
		}

		// stops the manager
		public void Stop()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			SetEvents(false);
		}

		// calls f() for each opened document
		//
		public void ForEachDocument(Action<DocumentWrapper> f)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			foreach (Document d in dte.Documents)
				f(MakeDocumentWrapper(d));
		}

		// either registers or unregisters the events
		//
		void SetEvents(bool add)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// todo: registering for all three events may not be necessary, as
			// all of them happen every time a document is opened, but the
			// frame might not be available in all of them
			//
			// in particular, DocumentOpened sometimes gives documents without
			// frames
			//
			// so this will unfortunately change captions much more often than
			// necessary

			if (add)
			{
				Logger.Log("adding events");
				docEvents.DocumentOpened += OnDocumentOpened;
				winEvents.WindowCreated += OnWindowCreated;
				winEvents.WindowActivated += OnWindowActivated;
			}
			else
			{
				Logger.Log("removing events");
				docEvents.DocumentOpened -= OnDocumentOpened;
				winEvents.WindowCreated -= OnWindowCreated;
				winEvents.WindowActivated -= OnWindowActivated;
			}
		}

		// fired when a document has been opened
		//
		private void OnDocumentOpened(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("document opened: {0}", d.FullName);

			DocumentChanged?.Invoke(MakeDocumentWrapper(d));
		}

		// fired when a window has been created
		//
		private void OnWindowCreated(Window w)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("window created: {0}", w.Caption);

			if (w.Document == null)
			{
				// this happens for windows like the solution explorer
				Logger.Log("not a document");
				return;
			}

			DocumentChanged?.Invoke(MakeDocumentWrapper(w.Document));
		}

		// fired when a window has been activated
		//
		private void OnWindowActivated(Window w, Window lost)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("window activated: {0}", w.Caption);

			if (w.Document == null)
			{
				// this happens for windows like the solution explorer
				Logger.Log("not a document");
				return;
			}

			DocumentChanged?.Invoke(MakeDocumentWrapper(w.Document));
		}

		private DocumentWrapper MakeDocumentWrapper(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// this is messy
			//
			// the events above either give a Document or a Window, but neither
			// can be used to change the caption, this can only be done on a
			// IVsWindowFrame
			//
			// there doesn't seem to be any good way of getting a IVsWindowFrame
			// from a Document except for IsDocumentOpen()
			//
			// it checks if a document is open by matching full paths, which
			// isn't great, but seems to be enough; a side-effect is that it
			// also provides the associated IVsWindowFrame if the document is
			// opened
			//
			// note that a document might be open, but without a frame, which
			// seems to happen mostly while a project is being loaded

			VsShellUtilities.IsDocumentOpen(
				CustomTabNames.Instance.ServiceProvider,
				d.FullName, VSConstants.LOGVIEWID.Primary_guid,
				out _, out _, out var f);

			return new DocumentWrapper(d, f);
		}
	}
}
