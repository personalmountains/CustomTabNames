﻿using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CustomTabNames
{
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
