using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;

namespace CustomTabNames
{
	// wraps a Document and a IVsWindowFrame
	//
	// the Document is used to feed information to variables, like path and
	// filters; the IVsWindowFrame is required to set the actual caption
	//
	public sealed class DocumentWrapper : LoggingContext
	{
		public Document Document { get; private set; }
		private IVsWindowFrame Frame { get; set; }

		protected override string LogPrefix()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			return "DocumentWrapper " + Document.FullName;
		}

		public DocumentWrapper(Document d, IVsWindowFrame f)
		{
			Document = d;
			Frame = f;
		}

		// sets the caption of this document to the given string
		//
		public void SetCaption(string s)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Log("setting to {0}", s);

			// the visible caption is made of a combination of the EditorCaption
			// and OwnerCaption; setting the EditorCaption to null makes sure
			// the caption can be controlled uniquely by OwnerCaption
			Frame.SetProperty((int)VsFramePropID.EditorCaption, null);
			Frame.SetProperty((int)VsFramePropID.OwnerCaption, s);
		}

		// resets the caption of this document to the default value
		//
		public void ResetCaption()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Log("resetting to {0}", Document.Name);

			SetCaption(Document.Name);
		}
	}


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
		public delegate void DocumentChangedHandler(DocumentWrapper d);
		public event DocumentChangedHandler DocumentChanged;

		// fired when projects or filters are added, removed or renamed
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


		private void OnProjectAdded(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("project {0} added", Utilities.DebugHierarchyName(h));
			ContainersChanged?.Invoke();
		}

		private void OnProjectRemoved(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("project {0} removed", Utilities.DebugHierarchyName(h));
			ContainersChanged?.Invoke();
		}

		private void OnProjectRenamed(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("project {0} renamed", Utilities.DebugHierarchyName(h));
			ContainersChanged?.Invoke();
		}

		private void OnFolderRenamed(IVsHierarchy h, uint item)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("folder {0} renamed", Utilities.DebugHierarchyName(h, item));
			ContainersChanged?.Invoke();
		}

		private void OnDocumentRenamed(Document d, IVsWindowFrame wf)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("document {0} renamed", d.FullName);
			DocumentChanged?.Invoke(new DocumentWrapper(d, wf));
		}

		private void OnDocumentOpened(Document d, IVsWindowFrame wf)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("document {0} opened", d.FullName);
			DocumentChanged?.Invoke(new DocumentWrapper(d, wf));
		}
	}
}
