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
				Utilities.DebugHierarchyName(Hierarchy);
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


		public override int OnItemAdded(
			uint parent, uint prevSibling, uint item)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Trace(
				"OnItemAdded: parent={0} prevSibling={1} item={2}",
				parent, prevSibling, item);

			// it's generally impossible to differentiate between moves and
			// renames for either files or folders

			if (Utilities.ItemIsFolder(Hierarchy, item))
			{
				FolderRenamed?.Invoke(Hierarchy, item);
			}
			else
			{
				var d = Utilities.DocumentFromItemID(Hierarchy, item);

				if (d == null)
				{
					// this happens when renaming C# files, but it's fine
					// because it's already handled in OnAfterAttributeChangeEx
					return VSConstants.S_OK;
				}

				var project = d.ProjectItem?.ContainingProject?.Kind;
				const string csProject = "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}";

				if (project == csProject)
				{
					// don't fire for C# projects; the reverse is tempting, but it's
					// better to have multiple events than none
					return VSConstants.S_OK;
				}


				DocumentRenamed?.Invoke(Hierarchy, item);
			}

			return VSConstants.S_OK;
		}

		public override int OnPropertyChanged(uint item, int prop, uint flags)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (prop != (int)__VSHPROPID.VSHPROPID_Caption)
			{
				// ignore everything but changing the caption
				return VSConstants.S_OK;
			}

			Trace(
				"OnPropertyChanged caption: item={0} prop={1} flags={2}",
				item, prop, flags);

			if (item == (uint)VSConstants.VSITEMID.Root)
			{
				ProjectRenamed?.Invoke(Hierarchy);
			}
			else if (Utilities.ItemIsFolder(Hierarchy, item))
			{
				FolderRenamed?.Invoke(Hierarchy, item);
			}

			return VSConstants.S_OK;
		}
	}
}
