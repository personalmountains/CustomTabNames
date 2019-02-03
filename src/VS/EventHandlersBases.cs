using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace CustomTabNames
{

// these are used as base classes in EventHandlers
//
// they define delegates and also implement the methods that are not
// used so the various concrete handlers can look a bit cleaner
//
// these are the events that seem to be fired for various components:
//
//   S=SolutionEvents  H=HierarchyEvents  D=DocumentEvents
//
//      when an action fires more than one event, the one in brackets
//      is ignored
//
// +----------+------------------------------+--------------------------------+
// |          | C++                          | C#                             |
// +----------+------------------------------+--------------------------------+
// | add      | S.OnAfterOpenProject         | S.OnAfterOpenProject           |
// | project  |                              |                                |
// +----------+------------------------------+--------------------------------+
// | remove   | S.OnBeforeCloseProject       | S.OnBeforeCloseProject         |
// | project  |                              |                                |
// +----------+------------------------------+--------------------------------+
// | rename   | H.OnPropertyChanged          | H.OnPropertyChanged            |
// | project  |                              | (D.OnAfterAttributeChangeEx)   |
// +----------+------------------------------+--------------------------------+
// | rename   | H.OnPropertyChanged          | H.OnItemAdded                  |
// | folder   |                              |                                |
// +----------+------------------------------+--------------------------------+
// | move     | H.OnItemAdded                | H.OnItemAdded                  |
// | folder   |                              |                                |
// +----------+------------------------------+--------------------------------+
// | rename   | D.OnAfterAttributeChangeEx   | D.OnAfterAttributeChangeEx     |
// | file     | (H.OnPropertyChanged x3)     | (H.OnItemAdded)                |
// +----------+------------------------------+--------------------------------+
// | move     | H.OnItemAdded                | D.OnAfterAttributeChangeEx     |
// | file     |                              | (H.OnItemAdded)                |
// +----------+------------------------------+--------------------------------+
// | open     | D.OnBeforeDocumentWindowShow | D.OnBeforeDocumentWindowShow   |
// | file     |                              |                                |
// +----------+------------------------------+--------------------------------+
//
//
//


	public abstract class EventHandlersBase : LoggingContext
	{
		public delegate void ProjectHandler(ITreeItem h);
		public delegate void FolderHandler(ITreeItem h);
		public delegate void DocumentHandler(IDocument d);
	}


	public abstract class HierarchyEventHandlersBase :
		EventHandlersBase, IVsHierarchyEvents
	{
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

		public abstract int OnItemAdded(
			uint itemidParent, uint itemidSiblingPrev, uint itemidAdded);

		public abstract int OnPropertyChanged(
			uint itemid, int propid, uint flags);
	}


	public abstract class SolutionEventHandlersBase :
		EventHandlersBase,
		IVsSolutionEvents, IVsSolutionEvents2,
		IVsSolutionEvents3, IVsSolutionEvents4
	{
		public int OnAfterRenameProject(IVsHierarchy hierarchy)
		{
			return VSConstants.S_OK;
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

		public abstract int OnAfterOpenProject(IVsHierarchy h, int fAdded);
		public abstract int OnBeforeCloseProject(IVsHierarchy h, int fRemoved);
	}


	public abstract class DocumentEventHandlersBase :
		EventHandlersBase,
		IVsRunningDocTableEvents, IVsRunningDocTableEvents2,
		IVsRunningDocTableEvents3, IVsRunningDocTableEvents4
	{
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

		public abstract int OnBeforeDocumentWindowShow(
			uint docCookie, int fFirstShow, IVsWindowFrame pFrame);

		public abstract int OnAfterAttributeChangeEx(
			uint docCookie, uint grfAttribs,
			IVsHierarchy pHierOld, uint itemidOld, string pszMkDocumentOld,
			IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew);
	}
}
