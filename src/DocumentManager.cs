using System;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using EnvDTE80;

namespace CustomTabNames
{
	public sealed class DocumentWrapper
	{
		public Document document;
		public IVsWindowFrame frame;

		public bool SetCaption(string s)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (frame == null)
				return false;

			frame.SetProperty((int)VsFramePropID.EditorCaption, null);
			frame.SetProperty((int)VsFramePropID.OwnerCaption, s);

			return true;
		}

		public void ResetCaption()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			SetCaption(document.Name);
		}
	}


	public sealed class DocumentManager
	{
		private readonly DTE2 dte;
		private readonly DocumentEvents docEvents;
		private readonly WindowEvents winEvents;

		public delegate void DocumentChangedHandler(DocumentWrapper d);
		public event DocumentChangedHandler DocumentChanged;

		public DocumentManager(DTE2 dte)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			this.dte = dte;
			this.docEvents = dte.Events.DocumentEvents;
			this.winEvents = dte.Events.WindowEvents;
		}

		public void Start()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			SetEvents(true);
		}

		public void Stop()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			SetEvents(false);
		}

		public void ForEachDocument(Action<DocumentWrapper> f)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			foreach (Document d in dte.Documents)
				f(MakeDocumentWrapper(d));
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
			}
			else
			{
				Logger.Log("removing events");
				docEvents.DocumentOpened -= OnDocumentOpened;
				winEvents.WindowCreated -= OnWindowCreated;
				winEvents.WindowActivated -= OnWindowActivated;
			}
		}

		private void OnDocumentOpened(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Logger.Log("document opened: {0}", d.FullName);
			DocumentChanged?.Invoke(MakeDocumentWrapper(d));
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

			DocumentChanged?.Invoke(MakeDocumentWrapper(w.Document));
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

			DocumentChanged?.Invoke(MakeDocumentWrapper(w.Document));
		}

		private DocumentWrapper MakeDocumentWrapper(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var dw = new DocumentWrapper();
			dw.document = d;

			VsShellUtilities.IsDocumentOpen(
				CustomTabNames.Instance.ServiceProvider,
				d.FullName, VSConstants.LOGVIEWID.Primary_guid,
				out IVsUIHierarchy h, out uint id, out dw.frame);

			return dw;
		}
	}
}
