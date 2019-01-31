using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;

// these are the events that seem to be fired for various components:
//
//                  C++                            C#
// add project      Sol.OnAfterOpenProject         same
// remove project   Sol.OnBeforeCloseProject       same
// rename project   Hier.OnPropertyChanged         Hier.OnPropertyChanged
//                  --                             Doc.OnAfterAttributeChangeEx
//                  --                             Sol.OnAfterRenameProject
//
// rename folder    Hier.OnPropertyChanged         Hier.OnItemAdded
// move folder      Hier.OnItemAdded               same
//
// rename file      Hier.OnPropertyChanged x3      Hier.OnItemAdded
//                  Doc.OnAfterAttributeChangeEx   same
//                    (only when opened)
// move file        Hier.OnItemAdded               same
//                  --                             Doc.OnAfterAttributeChangeEx
//                                                   (only when opened)
// open file        OnBeforeDocumentWindowShow     same
//
//
// there seems to be a bug where the first folder move in a C# project doesn't
// trigger _any_ handlers at all, no idea how to fix that
//
// todo: renaming a C# folder triggers both a full update _and_ an
// OnAfterAttributeChangeEx per file in that folder (and it's recursive!)


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


	public sealed class SolutionEventHandlers : SolutionEventHandlersBase
	{
		public event ProjectHandler ProjectAdded, ProjectRemoved, ProjectRenamed;
		public event FolderHandler FolderRenamed;
		public event DocumentHandler DocumentRenamed;

		// registration cookie, used in Unregister()
		private uint cookie = VSConstants.VSCOOKIE_NIL;

		// hierarchy handlers are per-project, this remembers them
		private readonly Dictionary<string, HierarchyEventHandlers>
			hierarchyHandlers =
				new Dictionary<string, HierarchyEventHandlers>();

		// see OnBeforeCloseProject()
		private readonly MainThreadTimer projectCloseTimer
			= new MainThreadTimer();


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
			Utilities.ForEachProjectHierarchy((h) =>
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
		public override int OnAfterOpenProject(
			IVsHierarchy hierarchy, int added)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("OnAfterOpenProject");

			// register for hierarchy events for this project and add it to the
			// list
			AddProjectHierarchy(hierarchy);

			ProjectAdded?.Invoke(hierarchy);

			return VSConstants.S_OK;
		}

		// fired just _before_ a project is removed from the solution; there
		// is no corresponding OnAfterCloseProject, which is unfortunate, see
		// DocumentManager.OnProjectCountChanged()
		//
		public override int OnBeforeCloseProject(
			IVsHierarchy hierarchy, int removed)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("OnBeforeCloseProject");

			// unregister for hierarchy events and remove it from the list
			RemoveProjectHierarchy(hierarchy);

			// this is an "on before" handler, and so the project count hasn't
			// been updated yet
			//
			// there is no On*After*CloseProject
			//
			// it's unclear what kind of delay there is between
			// OnBeforeCloseProject and the actual update of the project count,
			// but one second works for now
			//
			// in any case, the close can probably still be canceled by the
			// user if there are unsaved changes, so this may be a false
			// positive
			//
			// todo: try to find a way to get a callback to fire when the count
			// actually changes instead of using a stupid timer

			projectCloseTimer.Start(1000, () =>
			{
				ProjectRemoved?.Invoke(hierarchy);
			});

			return VSConstants.S_OK;
		}

		// fired after a project is renamed on disk
		//

		// fired after a document is moved on disk
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

			hh.ProjectRenamed += ProjectRenamed;
			hh.FolderRenamed += FolderRenamed;
			hh.DocumentRenamed += DocumentRenamed;

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
	}


	// handles document events, like opening and renaming; these are also called
	// for project events, but they're ignored because they're ignored in the
	// solution or hierarchy events above
	//
	public sealed class DocumentEventHandlers : DocumentEventHandlersBase
	{
		public event DocumentHandler DocumentRenamed, DocumentOpened;

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
		public override int OnBeforeDocumentWindowShow(
			uint itemCookie, int first, IVsWindowFrame wf)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (first == 0)
			{
				// only handle the first time a document is shown
				return VSConstants.S_OK;
			}

			Trace(
				"OnBeforeDocumentWindowShow: {0}",
				Utilities.DebugWindowFrameName(wf));

			var d = Utilities.DocumentFromWindowFrame(wf);
			if (d == null)
			{
				Error("OnBeforeDocumentWindowShow: frame has no document");
				return VSConstants.S_OK;
			}

			if (!Utilities.ItemIDFromDocument(d, out var h, out var id))
			{
				Error("OnBeforeDocumentWindowShow: can't get item id");
				return VSConstants.S_OK;
			}

			DocumentOpened?.Invoke(h, id);
			return VSConstants.S_OK;
		}

		// fired when document attributes change, such as renaming, but also
		// dirty state, etc.
		//
		public override int OnAfterAttributeChangeEx(
				uint cookie, uint atts,
				IVsHierarchy oldHier, uint oldId, string oldPath,
				IVsHierarchy newHier, uint newId, string newPath)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

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

			if (!Utilities.ItemIDFromCookie(cookie, out var h, out var id))
			{
				Error("can't get hierarchy for cookie {0}", cookie);
				return VSConstants.S_OK;
			}

			if (Utilities.DocumentFromCookie(cookie) == null)
			{
				// this happens when renaming a C# project, which is handled
				// elsewhere, so that's fine
				return VSConstants.S_OK;
			}

			DocumentRenamed?.Invoke(h, id);

			return VSConstants.S_OK;
		}
	}
}
