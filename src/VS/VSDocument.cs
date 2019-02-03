using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Runtime.InteropServices;

namespace CustomTabNames
{
	public class VSDocument : LoggingContext, IDocument
	{
		private readonly EnvDTE.Document d;
		private IVsWindowFrame f;

		public VSDocument(EnvDTE.Document d, IVsWindowFrame f = null)
		{
			this.d = d;
			this.f = f;
		}

		protected override string LogPrefix()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (d == null)
				return "Document (null)";
			else
				return "Document " + d.FullName;
		}

		public string Path
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				return d.FullName;
			}
		}

		public string Name
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				return d.Name;
			}
		}

		public IProject Project
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				return new VSProject(d.ProjectItem.ContainingProject);
			}
		}

		public ITreeItem TreeItem
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (!VSDocument.ItemIDFromDocument(d, out var h, out var id))
					return null;

				return new VSTreeItem(h, id);
			}
		}

		public void SetCaption(string s)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Log("setting to {0}", s);

			// the visible caption is made of a combination of the
			// EditorCaption and OwnerCaption; setting the EditorCaption to
			// null makes sure the caption can be controlled uniquely by
			//OwnerCaption
			WindowFrame.SetProperty((int)VsFramePropID.EditorCaption, null);
			WindowFrame.SetProperty((int)VsFramePropID.OwnerCaption, s);
		}

		public void ResetCaption()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Log("resetting to {0}", d.Name);
			SetCaption(d.Name);
		}

		public IVsWindowFrame WindowFrame
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (f == null)
					f = WindowFrameFromDocument(d);

				return f;
			}
		}

		// returns a Document from the given cookie
		//
		public static Document DocumentFromCookie(uint cookie)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var mk = Package.Instance.RDT4.GetDocumentMoniker(cookie);

			if (mk == null)
			{
				Main.Instance.Logger.Error(
					"GetDocumentMoniker failed for cookie {0} failed",
					cookie);

				return null;
			}

			var wf = WindowFrameFromPath(mk);
			if (wf == null)
				return null;

			return DocumentFromWindowFrame(wf);
		}

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
				Main.Instance.Logger.Error(
					"cookie not found for {0}", d.FullName);
				return false;
			}

			Package.Instance.RDT4.GetDocumentHierarchyItem(
				cookie, out h, out itemid);

			if (h == null || itemid == (uint)VSConstants.VSITEMID.Nil)
			{
				Main.Instance.Logger.Error(
					"can't get hierarchy item for {0} (cookie {1})",
					d.FullName, cookie);

				return false;
			}

			return true;
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
				Main.Instance.Logger.ErrorCode(e,
					"DocumentFromID: GetProperty for extObject failed");

				return null;
			}

			if (o is ProjectItem pi)
			{
				// pi.Document sometimes throws instead of just returning null,
				// like for folders in a C# project; the goal is just to return
				// null if this isn't a document, so eat the exception and do
				// that

				try
				{
					return pi.Document;
				}
				catch (COMException)
				{
					return null;
				}
			}

			// not all items are project items, this happens particularly
			// with ForEachDocument, because GetRunningDocumentsEnum()
			// seems to return projects as well as documents
			//
			// therefore, don't warn, just ignore
			return null;
		}

		// returns a hierarchy and item for the given cookie
		//
		public static bool ItemIDFromCookie(
			uint cookie, out IVsHierarchy h, out uint itemid)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Package.Instance.RDT4.GetDocumentHierarchyItem(
				cookie, out h, out itemid);

			if (h == null || itemid == (uint)VSConstants.VSITEMID.Nil)
			{
				Main.Instance.Logger.Error(
					"can't get hierarchy item for cookie {0}", cookie);

				return false;
			}

			return true;
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

		// used for logging
		//
		public static string DebugWindowFrameName(IVsWindowFrame wf)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var d = DocumentFromWindowFrame(wf);
			if (d != null)
			{
				if (d.FullName.Length > 0)
					return d.FullName;
			}

			return "?";
		}
	}
}
