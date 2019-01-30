using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using System;
using System.Collections.Generic;

namespace CustomTabNames
{
	class HierarchyEventHandlers : IVsHierarchyEvents
	{
		private uint cookie = VSConstants.VSCOOKIE_NIL;

		public delegate void DocumentMovedHandler(DocumentWrapper d);
		public event DocumentMovedHandler DocumentMoved;

		public IVsHierarchy Hierarchy { get; private set; }

		public HierarchyEventHandlers(IVsHierarchy h)
		{
			Hierarchy = h;
		}

		public void Register()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Logger.Trace("registering for hierarchy events");
			Hierarchy.AdviseHierarchyEvents(this, out cookie);
		}

		public void Unregister()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Trace("unregistering from hierarchy events");

			if (cookie == VSConstants.VSCOOKIE_NIL)
				Logger.Error("cookie is nil");
			else
				Hierarchy.UnadviseHierarchyEvents(cookie);
		}


		public int OnItemAdded(uint itemidParent, uint itemidSiblingPrev, uint itemidAdded)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// when items are moved between filters, they are first deleted
			// and added again, and itemidSiblingPrev contains the sibling
			// before the deletion
			//
			// when new items are added to the project, OnItemDeleted() is
			// never called and itemidSiblingPrev is Nil

			if (itemidSiblingPrev == (uint)VSConstants.VSITEMID.Nil)
				return VSConstants.S_OK;

			// item was moved

			var e = Hierarchy.GetProperty(
				itemidAdded, (int)__VSHPROPID.VSHPROPID_ExtObject,
				out var extObject);

			if (e != VSConstants.S_OK || extObject == null)
			{
				Logger.Error(
					"OnItemAdded: GetProperty for ExtObject failed, {0}", e);

				return VSConstants.S_OK;
			}

			var pi = extObject as ProjectItem;

			var d = pi.Document;
			if (d == null)
			{
				// the item was moved, but the document probably isn't opened
				return VSConstants.S_OK;
			}

			var wf = DocumentManager.WindowFrameFromDocument(d);
			if (wf == null)
			{
				Logger.Error(
					"OnItemAdded: can't get window frame from document {0}",
					d.FullName);

				return VSConstants.S_OK;
			}

			DocumentMoved?.Invoke(new DocumentWrapper(d, wf));

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

		public int OnPropertyChanged(uint itemid, int propid, uint flags)
		{
			return VSConstants.S_OK;
		}
	}


	public sealed class SolutionEventHandlers :
		IVsSolutionEvents, IVsSolutionEvents2,
		IVsSolutionEvents3, IVsSolutionEvents4
	{
		public delegate void ProjectCountChangedHandler();
		public event ProjectCountChangedHandler ProjectCountChanged;

		public delegate void DocumentMovedHandler(DocumentWrapper d);
		public event DocumentMovedHandler DocumentMoved;

		private uint cookie = VSConstants.VSCOOKIE_NIL;

		private readonly Dictionary<string, HierarchyEventHandlers>
			hierarchyHandlers =
				new Dictionary<string, HierarchyEventHandlers>();

		public void Register()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var s = CustomTabNames.Instance.ServiceProvider.GetService(
				typeof(SVsSolution)) as IVsSolution;

			if (s == null)
			{
				Logger.Error("can't get SVsSolution");
				return;
			}

			Logger.Trace("registering for solution events");
			s.AdviseSolutionEvents(this, out cookie);

			CustomTabNames.Instance.DocumentManager.ForEachProjectHierarchy((h) =>
			{
				AddProjectHierarchy(h);
			});
		}

		public void Unregister()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var s = CustomTabNames.Instance.ServiceProvider.GetService(
				typeof(SVsSolution)) as IVsSolution;

			if (s == null)
			{
				Logger.Error("can't get SVsSolution");
				return;
			}

			foreach (var hh in hierarchyHandlers)
				RemoveProjectHierarchy(hh.Value.Hierarchy);

			hierarchyHandlers.Clear();

			Logger.Trace("unregistering from solution events");

			if (cookie == VSConstants.VSCOOKIE_NIL)
				Logger.Error("cookie is nil");
			else
				s.UnadviseSolutionEvents(cookie);
		}


		public int OnAfterOpenProject(IVsHierarchy hierarchy, int added)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Trace("OnAfterOpenProject");

			AddProjectHierarchy(hierarchy);
			ProjectCountChanged?.Invoke();

			return VSConstants.S_OK;
		}

		public int OnBeforeCloseProject(IVsHierarchy hierarchy, int removed)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Trace("OnBeforeCloseProject");

			RemoveProjectHierarchy(hierarchy);
			ProjectCountChanged?.Invoke();

			return VSConstants.S_OK;
		}

		public int OnAfterRenameProject(IVsHierarchy hierarchy)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Trace("OnAfterRenameProject");
			ProjectCountChanged?.Invoke();
			return VSConstants.S_OK;
		}

		private void OnDocumentMoved(DocumentWrapper d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			DocumentMoved?.Invoke(d);
		}

		private void AddProjectHierarchy(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var e = h.GetCanonicalName(
				(uint)VSConstants.VSITEMID.Root, out var cn);

			if (e != VSConstants.S_OK)
			{
				Logger.Error("GetCanonicalName() on hierarchy failed, {0}", e);
				return;
			}

			Logger.Log("adding hierarchy {0} to list", cn);

			if (hierarchyHandlers.ContainsKey(cn))
			{
				Logger.Warn("hierarchy list already contains {0}", cn);
			}
			else
			{
				var hh = new HierarchyEventHandlers(h);
				hh.Register();
				hh.DocumentMoved += OnDocumentMoved;

				hierarchyHandlers.Add(cn, hh);
			}
		}

		private void RemoveProjectHierarchy(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var e = h.GetCanonicalName(
				(uint)VSConstants.VSITEMID.Root, out var cn);

			if (e != VSConstants.S_OK)
			{
				Logger.Error("GetCanonicalName() on hierarchy failed, {0}", e);
				return;
			}

			hierarchyHandlers.TryGetValue(cn, out var hh);

			if (hh == null)
			{
				Logger.Log("hierarchy {0} not found in handlers", cn);
			}
			else
			{
				Logger.Log("removing hierarchy {0} from list", cn);
				hh.Unregister();
				hierarchyHandlers.Remove(cn);
			}
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


	public sealed class DocumentEventHandlers :
		IVsRunningDocTableEvents, IVsRunningDocTableEvents2,
		IVsRunningDocTableEvents3, IVsRunningDocTableEvents4
	{
		public delegate void DocumentOpenedHandler(DocumentWrapper d);
		public event DocumentOpenedHandler DocumentOpened, DocumentRenamed;

		private uint cookie = VSConstants.VSCOOKIE_NIL;

		public void Register()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var rdt = CustomTabNames.Instance.ServiceProvider.GetService(
				typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;

			if (rdt == null)
			{
				Logger.Error("can't get SVsRunningDocumentTable");
				return;
			}

			Logger.Trace("registering for document events");

			rdt.AdviseRunningDocTableEvents(this, out cookie);
		}

		public void Unregister()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var rdt = CustomTabNames.Instance.ServiceProvider.GetService(
				typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;

			if (rdt == null)
			{
				Logger.Error("can't get SVsRunningDocumentTable");
				return;
			}

			Logger.Trace("unregistering from document events");

			if (cookie == VSConstants.VSCOOKIE_NIL)
				Logger.Error("cookie is nil");
			else
				rdt.UnadviseRunningDocTableEvents(cookie);
		}


		public int OnBeforeDocumentWindowShow(
			uint itemCookie, int first, IVsWindowFrame wf)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (first == 0)
			{
				// only handle the first time a document is shown
				return VSConstants.S_OK;
			}

			Window w = VsShellUtilities.GetWindowObject(wf);
			if (w == null)
			{
				Logger.Error(
					"OnBeforeDocumentWindowShow, " +
					"couldn't get window object");

				return VSConstants.S_OK;
			}

			Document d = w.Document;
			if (d == null)
			{
				Logger.Error(
					"OnBeforeDocumentWindowShow, " +
					"couldn't get document from window object");

				return VSConstants.S_OK;
			}

			Logger.Trace("OnBeforeDocumentWindowShow {0}", d.Name);
			DocumentOpened?.Invoke(new DocumentWrapper(d, wf));

			return VSConstants.S_OK;
		}

		public int OnAfterAttributeChangeEx(
				uint itemCookie, uint atts,
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

			Logger.Trace(
				"OnAfterAttributeChangeEx renamed {0} to {1} ({2})",
				oldPath, newPath, atts);

			var f = DocumentManager.WindowFrameFromPath(newPath);
			if (f == null)
			{
				// this seems to happen when renaming projects, not sure why
				return VSConstants.S_OK;
			}

			var w = VsShellUtilities.GetWindowObject(f);
			if (w == null)
			{
				Logger.Error(
					"OnAfterAttributeChangeEx can't get " +
					"window object from frame");

				return VSConstants.S_OK;
			}

			var d = w.Document;
			if (d == null)
			{
				Logger.Error(
					"OnAfterAttributeChangeEx can't get document " +
					"from window object");

				return VSConstants.S_OK;
			}

			DocumentRenamed?.Invoke(new DocumentWrapper(d, f));

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
