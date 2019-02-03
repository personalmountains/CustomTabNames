using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CustomTabNames
{
	// hierarchy events are fired when items in the solution explorer change;
	// this is used to update captions when folders are renamed, items are moved
	// between folders, projects are renamed, etc.
	//
	// while the solution and document events are global, the hierarchy events
	// must be registered per project; this is done in two places in
	// SolutionEventHandlers: when projects are opened and closed, and in
	// Register()
	//
	class HierarchyEventHandlers : HierarchyEventHandlersBase
	{
		public event ProjectHandler ProjectRenamed;
		public event FolderHandler FolderRenamed;
		public event DocumentHandler DocumentRenamed;

		// registration cookie, used in Unregister()
		private uint cookie = VSConstants.VSCOOKIE_NIL;

		// the project this handler has been registered for
		public IVsHierarchy Hierarchy { get; private set; }


		public HierarchyEventHandlers(IVsHierarchy h)
		{
			Hierarchy = h;
		}

		protected override string LogPrefix()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			return
				"HierarchyEventHandlers " +
				VSTreeItem.MakeDebugName(Hierarchy);
		}

		// registers for events on the project
		//
		public void Register()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("advising");

			var e = Hierarchy.AdviseHierarchyEvents(this, out cookie);

			if (e != VSConstants.S_OK)
			{
				// just in case
				cookie = VSConstants.VSCOOKIE_NIL;
				ErrorCode(e, "AdviseHierarchyEvents() failed");
			}
		}

		// unregisters for events on the project
		//
		public void Unregister()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("unadvising");

			if (cookie == VSConstants.VSCOOKIE_NIL)
			{
				Error("cookie is nil");
				return;
			}

			var e = Hierarchy.UnadviseHierarchyEvents(cookie);

			if (e != VSConstants.S_OK)
				ErrorCode(e, "failed to unadvise");
		}


		// fired when an item is moved or added
		//
		public override int OnItemAdded(
			uint parent, uint prevSibling, uint item)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Trace(
				"OnItemAdded: parent={0} prevSibling={1} item={2}",
				parent, prevSibling, item);

			// it's generally impossible to differentiate between moves and
			// renames for either files or folders, so they're merged into one
			// rename event

			// OnItemAdded for folders and files is fired for C# projects on
			// move and rename, but only on move for C++ projects for move
			//
			// renaming a file or a folder for C++ projects fires
			// OnPropertyChanged, which is handled below

			if (VSTreeItem.GetIsFolder(Hierarchy, item))
			{
				FolderRenamed?.Invoke(new VSTreeItem(Hierarchy, item));
			}
			else
			{
				var d = VSDocument.DocumentFromItemID(Hierarchy, item);
				if (d == null)
				{
					// this happens when renaming C# files for whatever reason,
					// but it's fine because it's already handled in
					// OnAfterAttributeChangeEx
					return VSConstants.S_OK;
				}

				DocumentRenamed?.Invoke(new VSDocument(d));
			}

			return VSConstants.S_OK;
		}

		// fired when various properties are changed on items, like the caption,
		// but also expanded state, etc.
		//
		public override int OnPropertyChanged(uint item, int prop, uint flags)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (prop != (int)__VSHPROPID.VSHPROPID_Caption)
			{
				// ignore everything except renaming
				return VSConstants.S_OK;
			}

			Trace(
				"OnPropertyChanged caption: item={0} prop={1} flags={2}",
				item, prop, flags);

			// this is fired when renaming:
			//
			// 1) projects
			// this is handled here for all projects and ignored in
			// DocumentEvents.OnAfterAttributeChangeEx
			//
			// 2) C++ folders
			// this is handled here
			//
			// 3) C++ files
			// this is ignored because DocumentEvents.OnAfterAttributeChangeEx
			// is also fired; for whatever reason, OnPropertyChanged is fired
			// *3*  times with the exact same arguments, and so cannot be
			// differentiated, which would force an update three times every
			// time a file is renamed

			if (item == (uint)VSConstants.VSITEMID.Root)
			{
				// itemid of root means this is a project
				ProjectRenamed?.Invoke(new VSTreeItem(Hierarchy, item));
			}
			else if (VSTreeItem.GetIsFolder(Hierarchy, item))
			{
				// this is a C++ folder
				FolderRenamed?.Invoke(new VSTreeItem(Hierarchy, item));
			}

			return VSConstants.S_OK;
		}
	}
}
