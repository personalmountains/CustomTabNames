using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;

namespace CustomTabNames
{
	public sealed class SolutionEventHandlers :
		IVsSolutionEvents, IVsSolutionEvents2,
		IVsSolutionEvents3, IVsSolutionEvents4
	{
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

		public int OnAfterOpenProject(IVsHierarchy hierarchy, int added)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
		{
			return VSConstants.S_OK;
		}

		public int OnAfterOpeningChildren(IVsHierarchy hierarchy)
		{
			return VSConstants.S_OK;
		}

		public int OnBeforeCloseProject(IVsHierarchy hierarchy, int removed)
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

		public int OnAfterRenameProject(IVsHierarchy hierarchy)
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

		public int OnBeforeDocumentWindowShow(
			uint cookie, int first, IVsWindowFrame wf)
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

			Logger.Trace(
				"OnAfterAttributeChangeEx renamed {0} to {1} ({2})",
				oldPath, newPath, atts);

			var f = DocumentManager.WindowFrameFromPath(newPath);
			if (f == null)
			{
				Logger.Error("OnAfterAttributeChangeEx can't get frame");
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
