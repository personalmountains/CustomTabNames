using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;

namespace CustomTabNames
{
	// manages the various events for opening documents and windows, and fires
	// DocumentChanged when they do
	//
	public abstract class DocumentManager : LoggingContext
	{
		// fired every time a document changes in a way that may require
		// fixing the caption
		public delegate void DocumentChangedHandler(IDocument d);
		public event DocumentChangedHandler DocumentChanged;

		// fired when projects or folders are added, removed or renamed
		//
		public delegate void ContainersChangedHandler();
		public event ContainersChangedHandler ContainersChanged;


		public DocumentManager()
		{
		}

		protected override string LogPrefix()
		{
			return "DocumentManager";
		}

		// starts the manager
		//
		public abstract void Start();

		// stops the manager
		//
		public abstract void Stop();


		protected void OnProjectAdded(ITreeItem i)
		{
			Trace("project {0} added", i.DebugName);
			ContainersChanged?.Invoke();
		}

		protected void OnProjectRemoved(ITreeItem i)
		{
			Trace("project {0} removed", i.DebugName);
			ContainersChanged?.Invoke();
		}

		protected void OnProjectRenamed(ITreeItem i)
		{
			Trace("project {0} renamed", i.DebugName);
			ContainersChanged?.Invoke();
		}

		protected void OnFolderRenamed(ITreeItem i)
		{
			Trace("folder {0} renamed", i.DebugName);
			ContainersChanged?.Invoke();
		}

		protected void OnDocumentRenamed(IDocument d)
		{
			Trace("document {0} renamed", d.Path);
			DocumentChanged?.Invoke(d);
		}

		protected void OnDocumentOpened(IDocument d)
		{
			Trace("document {0} opened", d.Path);
			DocumentChanged?.Invoke(d);
		}
	}
}
