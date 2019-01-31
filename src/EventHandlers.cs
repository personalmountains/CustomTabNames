using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;

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
	class HierarchyEventHandlers : LoggingContext, IVsHierarchyEvents
	{
		// registration cookie, used in Unregister()
		private uint cookie = VSConstants.VSCOOKIE_NIL;

		// fired when documents are moved between folders
		public delegate void DocumentMovedHandler(DocumentWrapper d);
		public event DocumentMovedHandler DocumentMoved;

		// fired when folders or projects are renamed
		public delegate void ContainerNameChangedHandler();
		public event ContainerNameChangedHandler ContainerNameChanged;

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


		public int OnItemAdded(uint parent, uint prevSibling, uint item)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Trace(
				"OnItemAdded: parent={0} prevSibling={1} item={2}",
				parent, prevSibling, item);

			// there's no real way of knowing whether this is called because
			// an item is moved (delete+add) or a new item is added to the
			// project
			//
			// it might be possible to handle OnItemDeleted and remember the
			// itemid to see if it's the same here (it seems to be), but that
			// sounds error-prone (what if it's actually a new item that reused
			// an old id at the same time?) and a PITA (I need to keep a list
			// of items in sync, purge it regularly for items that were actually
			// deleted, etc.)
			//
			// however, prevSibling looks promising: it's usually Nil for items
			// that were just added, and not Nil for items that were moved, but
			// there seems to be a few cases where items are moved and
			// prevSibling is still Nil, so it's not full-proof
			//
			// therefore, adding a file might fire twice: once here, and once
			// in the document events below

			var d = Utilities.DocumentFromItemID(Hierarchy, item);
			if (d == null)
			{
				// this may be a folder or an unopened document

				if (Utilities.ItemIsFolder(Hierarchy, item))
				{
					if (ShouldFireForFolder(item))
					{
						// this is a non-empty folder that was moved
						//
						// todo: this could perhaps be more efficient by only
						// fixing documents that are under this folder

						ContainerNameChanged?.Invoke();
					}

					// this isn't a document, so don't bother with the rest

					return VSConstants.S_OK;
				}

				// this isn't a folder and it has no associated document, this
				// happens for unopened documents, also when renaming c# files
				// even if they're opened, who knows why
				Trace("OnItemAdded: no document");
				return VSConstants.S_OK;
			}

			var wf = Utilities.WindowFrameFromDocument(d);
			if (wf == null)
			{
				// this sometimes happens when an item was moved in the
				// hierarchy without being opened
				Error("OnItemAdded: document {0} has no frame", d.FullName);
				return VSConstants.S_OK;
			}

			DocumentMoved?.Invoke(new DocumentWrapper(d, wf));

			return VSConstants.S_OK;
		}

		// fired when properties on an item change, like renaming
		//
		public int OnPropertyChanged(uint item, int prop, uint flags)
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

			if (Utilities.ItemIsFolder(Hierarchy, item))
			{
				// don't trigger a full update for empty folders

				if (!ShouldFireForFolder(item))
					return VSConstants.S_OK;
			}

			// this catches both renaming projects and files
			// todo: this could perhaps be more efficient by only fixing
			// documents that are under this folder (if it's a folder)

			ContainerNameChanged?.Invoke();

			return VSConstants.S_OK;
		}


		public int OnInvalidateIcon(IntPtr hicon)
		{
			return VSConstants.S_OK;
		}

		public int OnInvalidateItems(uint itemidParent)
		{
			return VSConstants.S_OK;
		}

		public int OnItemDeleted(uint itemid)
		{
			return VSConstants.S_OK;
		}

		public int OnItemsAppended(uint itemidParent)
		{
			return VSConstants.S_OK;
		}

		private bool ShouldFireForFolder(uint item)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// todo: this doesn't handle a folder that contains other empty
			// folders

			var e = Hierarchy.GetProperty(
				item, (int)__VSHPROPID.VSHPROPID_FirstChild,
				out var childObject);

			// careful: for whatever reason, FirstChild is an int instead
			// of a uint

			if (e != VSConstants.S_OK || !(childObject is int))
			{
				ErrorCode(
					e, "OnPropertyChanged failed to get folder " +
					"first child");
			}
			else
			{
				var child = (uint)(int)childObject;

				if (child == (uint)VSConstants.VSITEMID.Nil)
				{
					Trace("this is a folder without children, ignoring");
					return false;
				}
			}

			return true;
		}
	}


	// handles solution events, like adding, renaming and removing projects
	//
	// note most of these events are only fired when a corresponding change is
	// made to files *on disk* (rename, move, etc.), which happens for C#
	// projects, but not C++ projects
	//
	// for example, when a C# project is renamed, its .csproj is also renamed
	// automatically and this fires the handler
	//
	// when a C++ project is renamed, the content of its .vcxproj is changed,
	// but it is _not_ renamed, and so these handlers don't fire; these events
	// are instead handled by the hierarchy handlers above, which fire on
	// renaming tree item captions
	//
	// todo: are both necessary? isn't the hierarchy event enough for both?
	//
	public sealed class SolutionEventHandlers :
		LoggingContext,
		IVsSolutionEvents, IVsSolutionEvents2,
		IVsSolutionEvents3, IVsSolutionEvents4
	{
		// fired when projects were added or removed
		public delegate void ProjectCountChangedHandler();
		public event ProjectCountChangedHandler ProjectCountChanged;

		// fired when a document was moved in the tree
		public delegate void DocumentMovedHandler(DocumentWrapper d);
		public event DocumentMovedHandler DocumentMoved;

		// fired when a project was renamed
		public delegate void ContainerNameChangedHandler();
		public event ContainerNameChangedHandler ContainerNameChanged;

		// registration cookie, used in Unregister()
		private uint cookie = VSConstants.VSCOOKIE_NIL;

		// hierarchy handlers are per-project, this remembers them
		private readonly Dictionary<string, HierarchyEventHandlers>
			hierarchyHandlers =
				new Dictionary<string, HierarchyEventHandlers>();


		protected override string LogPrefix()
		{
			return "SolutionEventHandlers";
		}

		// register for solution events and walks all current projects to
		// register for their own hierarchy events: this is necessary because
		// the extension might initialize _after_ projects are loaded, and
		// therefore won't receive the corresponding events
		//
		public void Register()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("registering for events");

			var e = Package.Instance.Solution
				.AdviseSolutionEvents(this, out cookie);

			if (e != VSConstants.S_OK)
			{
				ErrorCode(e, "AdviseSolutionEvents() failed");

				// don't bother with the rest, there's a major problem
				return;
			}

			// go through every loaded project and register for hierarchy
			// events; this handles cases where the extension was loaded after
			// the current solution
			DocumentManager.ForEachProjectHierarchy((h) =>
			{
				AddProjectHierarchy(h);
			});
		}

		// unregisters from solution events
		//
		public void Unregister()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// unregister all per-project handlers
			foreach (var hh in hierarchyHandlers)
				hh.Value.Unregister();

			hierarchyHandlers.Clear();

			Trace("unregistering from events");

			if (cookie == VSConstants.VSCOOKIE_NIL)
			{
				Error("cookie is nil");
				return;
			}

			var e = Package.Instance.Solution.UnadviseSolutionEvents(cookie);
			if (e != VSConstants.S_OK)
				ErrorCode(e, "UnadviseSolutionEvents() failed");
		}


		// fired once per project in the current solution
		//
		public int OnAfterOpenProject(IVsHierarchy hierarchy, int added)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("OnAfterOpenProject");

			// register for hierarchy events for this project and add it to the
			// list
			AddProjectHierarchy(hierarchy);

			ProjectCountChanged?.Invoke();

			return VSConstants.S_OK;
		}

		// fired just _before_ a project is removed from the solution; there
		// is no corresponding OnAfterCloseProject, which is unfortunate, see
		// DocumentManager.OnProjectCountChanged()
		//
		public int OnBeforeCloseProject(IVsHierarchy hierarchy, int removed)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("OnBeforeCloseProject");

			// unregister for hierarchy events and remove it from the list
			RemoveProjectHierarchy(hierarchy);

			ProjectCountChanged?.Invoke();
			return VSConstants.S_OK;
		}

		// fired after a project is renamed on disk
		//
		public int OnAfterRenameProject(IVsHierarchy hierarchy)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("OnAfterRenameProject");

			ContainerNameChanged?.Invoke();
			return VSConstants.S_OK;
		}

		// fired after a document is moved on disk
		//
		private void OnDocumentMoved(DocumentWrapper d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			DocumentMoved?.Invoke(d);
		}

		// fired by the HierarchyEventHandlers (per-project) when a folder or
		// the project itself was renamed
		//
		private void OnContainerNameChanged()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			ContainerNameChanged?.Invoke();
		}

		// called once per project, registers for project hierarchy events and
		// adds the handler to the list
		//
		private void AddProjectHierarchy(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// the canonical name of a project hierarchy object is unique, so
			// it is used as a key in the dictionary
			var e = h.GetCanonicalName(
				(uint)VSConstants.VSITEMID.Root, out var cn);

			if (e != VSConstants.S_OK)
			{
				ErrorCode(e, "AddProjectHierarchy: GetCanonicalName failed");
				return;
			}

			Log("AddProjectHierarchy: adding {0} to list", cn);

			if (hierarchyHandlers.ContainsKey(cn))
			{
				Error("AddProjectHierarchy: list already contains {0}", cn);
				return;
			}

			var hh = new HierarchyEventHandlers(h);

			// registers for hierarchy events on this project
			hh.Register();

			hh.DocumentMoved += OnDocumentMoved;
			hh.ContainerNameChanged += OnContainerNameChanged;

			hierarchyHandlers.Add(cn, hh);
		}

		private void RemoveProjectHierarchy(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// the canonical name of a project hierarchy object is unique, so
			// it is used as a key in the dictionary
			var e = h.GetCanonicalName(
				(uint)VSConstants.VSITEMID.Root, out var cn);

			if (e != VSConstants.S_OK)
			{
				ErrorCode(e, "RemoveProjectHierarchy: GetCanonicalName failed");
				return;
			}

			Log("RemoveProjectHierarchy: removing {0} from list", cn);

			hierarchyHandlers.TryGetValue(cn, out var hh);
			if (hh == null)
			{
				Error("RemoveProjectHierarchy: {0} not in list", cn);
				return;
			}

			// unregistering from events
			hh.Unregister();

			hierarchyHandlers.Remove(cn);
		}

		public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterCloseSolution(object pUnkReserved)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterClosingChildren(IVsHierarchy pHierarchy)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterLoadProject(
			IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterMergeSolution(object pUnkReserved)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterOpeningChildren(IVsHierarchy hierarchy)
		{
			return VSConstants.S_OK;
		}

		public int OnBeforeCloseSolution(object pUnkReserved)
		{
			return VSConstants.S_OK;
		}

		public int OnBeforeClosingChildren(IVsHierarchy hierarchy)
		{
			return VSConstants.S_OK;
		}

		public int OnBeforeOpeningChildren(IVsHierarchy hierarchy)
		{
			return VSConstants.S_OK;
		}

		public int OnBeforeUnloadProject(
			IVsHierarchy realHierarchy, IVsHierarchy rtubHierarchy)
		{
			return VSConstants.S_OK;
		}

		public int OnQueryCloseProject(
			IVsHierarchy hierarchy, int removing, ref int cancel)
		{
			return VSConstants.S_OK;
		}

		public int OnQueryCloseSolution(object pUnkReserved, ref int cancel)
		{
			return VSConstants.S_OK;
		}

		public int OnQueryUnloadProject(
			IVsHierarchy pRealHierarchy, ref int cancel)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterAsynchOpenProject(IVsHierarchy hierarchy, int added)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterChangeProjectParent(IVsHierarchy hierarchy)
		{
			return VSConstants.S_OK;
		}

		public int OnQueryChangeProjectParent(
			IVsHierarchy hierarchy, IVsHierarchy newParentHier, ref int cancel)
		{
			return VSConstants.S_OK;
		}
	}


	// handles document events, like opening and renaming; these are also called
	// for project events, but they're ignored because they're ignored in the
	// solution or hierarchy events above
	//
	// some of these events only fire when a change is made in filenames on
	// disk, see SolutionEventHandlers class above
	//
	public sealed class DocumentEventHandlers :
		LoggingContext,
		IVsRunningDocTableEvents, IVsRunningDocTableEvents2,
		IVsRunningDocTableEvents3, IVsRunningDocTableEvents4
	{
		public delegate void DocumentHandler(DocumentWrapper d);

		// fired when a document is opened
		public event DocumentHandler DocumentOpened;

		// fired when a document is renamed
		public event DocumentHandler DocumentRenamed;

		// registration cookie, used in Unregister()
		private uint cookie = VSConstants.VSCOOKIE_NIL;


		protected override string LogPrefix()
		{
			return "DocumentEventHandlers";
		}


		// registers for events
		//
		public void Register()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Trace("registering for events");

			var e = Package.Instance.RDT
				.AdviseRunningDocTableEvents(this, out cookie);

			if (e != VSConstants.S_OK)
				ErrorCode(e, "AdviseRunningDocTableEvents failed");
		}

		// unregisters from events
		//
		public void Unregister()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Trace("unregistering from document events");

			if (cookie == VSConstants.VSCOOKIE_NIL)
			{
				Error("cookie is nil");
				return;
			}

			var e = Package.Instance.RDT.UnadviseRunningDocTableEvents(cookie);

			if (e != VSConstants.S_OK)
				ErrorCode(e, "UnadviseRunningDocTableEvents failed");
		}


		// fired when a document is about to be shown on screen for various
		// reasons
		//
		public int OnBeforeDocumentWindowShow(
			uint itemCookie, int first, IVsWindowFrame wf)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (first == 0)
			{
				// only handle the first time a document is shown
				return VSConstants.S_OK;
			}

			if (wf == null)
			{
				// doesn't seem to happen, but handlers assume it's not null
				return VSConstants.S_OK;
			}

			var d = Utilities.DocumentFromWindowFrame(wf);
			if (d == null)
			{
				// this shouldn't happen
				Error("OnBeforeDocumentWindowShow: frame has no document");
				return VSConstants.S_OK;
			}

			Trace("OnBeforeDocumentWindowShow for {0}", d.Name);

			DocumentOpened?.Invoke(new DocumentWrapper(d, wf));
			return VSConstants.S_OK;
		}

		// fired when document attributes change, such as renaming, but also
		// dirty state, etc.
		//
		public int OnAfterAttributeChangeEx(
				uint cookie, uint atts,
				IVsHierarchy oldHier, uint oldId, string oldPath,
				IVsHierarchy newHier, uint newId, string newPath)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// note that this is also called when renaming projects

			const uint RenameBit = (uint)__VSRDTATTRIB.RDTA_MkDocument;

			if ((atts & RenameBit) == 0)
			{
				// don't bother with anything else than rename
				return VSConstants.S_OK;
			}

			Trace(
				"OnAfterAttributeChangeEx rename: cookie={0} atts={1} " +
				"oldId={2} oldPath={3} newId={4} newPath={5}",
				cookie, atts, oldId, oldPath, newId, newPath);

			var d = Utilities.DocumentFromCookie(cookie);
			if (d == null)
			{
				// this happens when a project was renamed, ignore it because
				// it will also fire a hierarchy event
				return VSConstants.S_OK;
			}

			var wf = Utilities.WindowFrameFromDocument(d);
			if (wf == null)
			{
				Error(
					"OnAfterAttributeChangeEx: frame not found for {0}",
					d.FullName);

				return VSConstants.S_OK;
			}

			DocumentRenamed?.Invoke(new DocumentWrapper(d, wf));

			return VSConstants.S_OK;
		}

		public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterDocumentWindowHide(
			uint docCookie, IVsWindowFrame pFrame)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterFirstDocumentLock(
			uint docCookie, uint dwRDTLockType,
			uint dwReadLocksRemaining, uint dwEditLocksRemaining)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterSave(uint docCookie)
		{
			return VSConstants.S_OK;
		}

		public int OnBeforeLastDocumentUnlock(
			uint docCookie, uint dwRDTLockType,
			uint dwReadLocksRemaining, uint dwEditLocksRemaining)
		{
			return VSConstants.S_OK;
		}

		public int OnBeforeSave(uint docCookie)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterLastDocumentUnlock(
			IVsHierarchy pHier, uint itemid,
			string pszMkDocument, int fClosedWithoutSaving)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterSaveAll()
		{
			return VSConstants.S_OK;
		}

		public int OnBeforeFirstDocumentLock(
			IVsHierarchy pHier, uint itemid, string pszMkDocument)
		{
			return VSConstants.S_OK;
		}
	}
}
