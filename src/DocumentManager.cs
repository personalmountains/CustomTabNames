using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;

namespace CustomTabNames
{
	// manages the various events for opening documents and windows, and fires
	// DocumentChanged when they do
	//
	public sealed class DocumentManager : LoggingContext
	{
		private readonly DocumentEventHandlers docHandlers
			 = new DocumentEventHandlers();

		private readonly SolutionEventHandlers solHandlers
			= new SolutionEventHandlers();

		// fired every time a document changes in a way that may require
		// fixing the caption
		public delegate void DocumentChangedHandler(IDocument d);
		public event DocumentChangedHandler DocumentChanged;

		// fired when projects or folders are added, removed or renamed
		//
		public delegate void ContainersChangedHandler();
		public event ContainersChangedHandler ContainersChanged;


		protected override string LogPrefix()
		{
			return "DocumentManager";
		}


		public DocumentManager()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// project events
			solHandlers.ProjectAdded += OnProjectAdded;
			solHandlers.ProjectRemoved += OnProjectRemoved;
			solHandlers.ProjectRenamed += OnProjectRenamed;

			// folder events
			solHandlers.FolderRenamed += OnFolderRenamed;

			// document events
			solHandlers.DocumentRenamed += OnDocumentRenamed;
			docHandlers.DocumentRenamed += OnDocumentRenamed;
			docHandlers.DocumentOpened += OnDocumentOpened;
		}

		// starts the manager
		//
		public void Start()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			SetEvents(true);
		}

		// stops the manager
		//
		public void Stop()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			SetEvents(false);
		}

		// either registers or unregisters the events
		//
		private void SetEvents(bool add)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (add)
			{
				docHandlers.Register();
				solHandlers.Register();
			}
			else
			{
				solHandlers.Unregister();
				docHandlers.Unregister();
			}
		}


		private void OnProjectAdded(ITreeItem i)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("project {0} added", i.DebugName);
			ContainersChanged?.Invoke();
		}

		private void OnProjectRemoved(ITreeItem i)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("project {0} removed", i.DebugName);
			ContainersChanged?.Invoke();
		}

		private void OnProjectRenamed(ITreeItem i)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("project {0} renamed", i.DebugName);
			ContainersChanged?.Invoke();
		}

		private void OnFolderRenamed(ITreeItem i)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("folder {0} renamed", i.DebugName);
			ContainersChanged?.Invoke();
		}

		private void OnDocumentRenamed(IDocument d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("document {0} renamed", d.Path);
			DocumentChanged?.Invoke(d);
		}

		private void OnDocumentOpened(IDocument d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("document {0} opened", d.Path);
			DocumentChanged?.Invoke(d);
		}
	}
}
