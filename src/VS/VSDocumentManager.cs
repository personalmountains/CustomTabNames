using Microsoft.VisualStudio.Shell;

namespace CustomTabNames
{
	// manages the various events for opening documents and windows, and fires
	// DocumentChanged when they do
	//
	public sealed class VSDocumentManager : DocumentManager
	{
		private readonly DocumentEventHandlers docHandlers
			 = new DocumentEventHandlers();

		private readonly SolutionEventHandlers solHandlers
			= new SolutionEventHandlers();

		public VSDocumentManager()
		{
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
		public override void Start()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			SetEvents(true);
		}

		// stops the manager
		//
		public override void Stop()
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
	}
}
