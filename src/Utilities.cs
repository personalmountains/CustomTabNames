﻿using System;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using Task = System.Threading.Tasks.Task;

namespace CustomTabNames
{
	public sealed class MainThreadTimer : IDisposable
	{
		private Timer t = null;

		public void Dispose()
		{
			t?.Dispose();
		}

		public void Start(int ms, Action a)
		{
			if (t == null)
				t = new Timer(OnTimer, a, ms, Timeout.Infinite);
			else
				t.Change(ms, Timeout.Infinite);
		}

		private void OnTimer(object a)
		{
			_ = OnMainThreadAsync((Action)a);
		}

		private async Task OnMainThreadAsync(Action a)
		{
			await Package.Instance
				.JoinableTaskFactory.SwitchToMainThreadAsync();

			a();
		}
	}


	public sealed class Utilities
	{
		// returns an IVsWindowFrame associated with the given path
		//
		// there doesn't seem to be any good way of getting a IVsWindowFrame
		// from a Document except for IsDocumentOpen()
		//
		// it checks if a document is open by matching full paths, which
		// isn't great, but seems to be enough; a side-effect is that it
		// also provides the associated IVsWindowFrame if the document is
		// opened
		//
		// note that a document might be open, but without a frame, which
		// seems to happen mostly while a project is being loaded, so this
		// may return null
		//
		public static IVsWindowFrame WindowFrameFromPath(string path)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			VsShellUtilities.IsDocumentOpen(
				Package.Instance,
				path, VSConstants.LOGVIEWID.Primary_guid,
				out _, out _, out var f);

			return f;
		}

		// returns an IVsWindowFrame associated with the given Document; see
		// WindowFrameFromPath()
		//
		public static IVsWindowFrame WindowFrameFromDocument(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			return WindowFrameFromPath(d.FullName);
		}

		// returns the Document associated with the given IVsWindowFrame
		//
		public static Document DocumentFromWindowFrame(IVsWindowFrame wf)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var w = VsShellUtilities.GetWindowObject(wf);
			if (w == null)
				return null;

			return w.Document;
		}

		// returns a Document associated with an itemid in a hierarchy
		//
		public static Document DocumentFromItemID(IVsHierarchy h, uint itemid)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var e = h.GetProperty(
				itemid, (int)__VSHPROPID.VSHPROPID_ExtObject, out var o);

			if (e != VSConstants.S_OK || o == null)
			{
				Logger.ErrorCode(e,
					"DocumentFromID: GetProperty for extObject failed");

				return null;
			}

			if (o is ProjectItem pi)
				return pi.Document;

			// not all items are project items, this happens particularly
			// with ForEachDocument, because GetRunningDocumentsEnum()
			// seems to return projects as well as documents
			//
			// therefore, don't warn, just ignore
			return null;
		}

		// returns a Document from the given cookie
		//
		public static Document DocumentFromCookie(uint cookie)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var mk = Package.Instance.RDT4.GetDocumentMoniker(cookie);

			if (mk == null)
			{
				Logger.Error(
					"GetDocumentMoniker failed for cookie {0} failed",
					cookie);

				return null;
			}

			var wf = WindowFrameFromPath(mk);
			if (wf == null)
				return null;

			return DocumentFromWindowFrame(wf);
		}

		// returns a hierarchy and itemid for the given document
		//
		public static bool ItemIDFromDocument(
			Document d, out IVsHierarchy h, out uint itemid)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			h = null;
			itemid = (uint)VSConstants.VSITEMID.Nil;

			var cookie = Package.Instance.RDT4.GetDocumentCookie(d.FullName);
			if (cookie == VSConstants.VSCOOKIE_NIL)
			{
				Logger.Error("cookie not found for {0}", d.FullName);
				return false;
			}

			Package.Instance.RDT4.GetDocumentHierarchyItem(
				cookie, out h, out itemid);

			if (h == null || itemid == (uint)VSConstants.VSITEMID.Nil)
			{
				Logger.Error(
					"can't get hierarchy item for {0} (cookie {1})",
					d.FullName, cookie);

				return false;
			}

			return true;
		}

		// returns an itemid's parent, can be Nil and still return true
		//
		public static bool ParentItemID(
			IVsHierarchy h, uint itemid, out uint parentItemid)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			parentItemid = (uint)VSConstants.VSITEMID.Nil;

			var e = h.GetProperty(
				itemid, (int)__VSHPROPID.VSHPROPID_Parent,
				out var pid);

			// for whatever reason, VSHPROPID_Parent returns an int instead of
			// a uint

			if (e != VSConstants.S_OK || !(pid is int))
			{
				Logger.ErrorCode(e, "can't get parent item", itemid);
				return false;
			}

			parentItemid = (uint)(int)pid;
			return true;
		}

		// returns the item's name, or null
		//
		public static string ItemName(IVsHierarchy h, uint itemid)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var e = h.GetProperty(
				itemid, (int)__VSHPROPID.VSHPROPID_Name,
				out var name);

			if (e != VSConstants.S_OK || !(name is string))
			{
				Logger.ErrorCode(e, "can't get itemid name");
				return null;
			}

			return (string)name;
		}

		// returns whether the given item is any type of folder; note that
		// some items may return true even if they're not actually folders,
		// see Variables.FilterPath()
		//
		public static bool ItemIsFolder(IVsHierarchy h, uint itemid)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var e = h.GetGuidProperty(
				itemid, (int)__VSHPROPID.VSHPROPID_TypeGuid,
				out var type);

			if (e != VSConstants.S_OK || type == null)
			{
				Logger.ErrorCode(e, "can't get typeguid");
				return false;
			}

			// ignore anything but folders
			if (type == VSConstants.ItemTypeGuid.PhysicalFolder_guid)
				return true;

			if (type == VSConstants.ItemTypeGuid.VirtualFolder_guid)
				return true;

			return false;
		}
	}
}
