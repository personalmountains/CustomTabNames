using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CustomTabNames
{
	public class VSSolution : LoggingContext, ISolution
	{
		protected override string LogPrefix()
		{
			return "Solution";
		}

		public bool HasSingleProject
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				try
				{
					var e = Package.Instance.SolutionService.GetProperty(
						(int)__VSPROPID.VSPROPID_ProjectCount, out var o);

					if (e != VSConstants.S_OK || !(o is int))
					{
						ErrorCode(
							e, "HasSingleProject: failed to get project count");

						return false;
					}

					int i = (int)o;
					return (i == 1);
				}
				catch (Exception e)
				{
					Error(
						"HasSingleProject: failed to get project count, {0}",
						e.Message);

					return false;
				}
			}
		}

		public List<ITreeItem> ProjectItems
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				var list = new List<ITreeItem>();

				Guid guid = Guid.Empty;

				// getting enumerator
				var e = Package.Instance.SolutionService.GetProjectEnum(
					(uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION,
					ref guid, out var enumerator);

				if (e != VSConstants.S_OK)
				{
					ErrorCode(
						e, "ForEachProjectHierarchy: GetProjectEnum failed");

					return list;
				}

				// will store one hierarchy at a time, but Next() still requires an
				// array
				IVsHierarchy[] hierarchies = new IVsHierarchy[1] { null };

				enumerator.Reset();

				while (true)
				{
					e = enumerator.Next(1, hierarchies, out var fetched);

					if (e == VSConstants.S_FALSE || fetched != 1)
					{
						// done
						break;
					}

					if (e != VSConstants.S_OK)
					{
						ErrorCode(
							e, "ForEachProjectHierarchy: enum next failed");

						break;
					}


					var h = hierarchies[0];

					if (h == null)
					{
						// shouldn't happen
						continue;
					}

					list.Add(new VSTreeItem(h));
				}

				return list;
			}
		}

		public List<IDocument> Documents
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				var list = new List<IDocument>();

				// getting enumerator
				var e = Package.Instance.RDT.GetRunningDocumentsEnum(
					out var enumerator);

				if (e != VSConstants.S_OK)
				{
					ErrorCode(
						e, "ForEachDocument: GetRunningDocumentsEnum failed");

					return list;
				}

				// will store one cookie at a time, but Next() still requires an
				// array
				uint[] cookies = new uint[1] { VSConstants.VSCOOKIE_NIL };

				enumerator.Reset();

				while (true)
				{
					e = enumerator.Next(1, cookies, out var fetched);

					if (e == VSConstants.S_FALSE || fetched != 1)
					{
						// done
						break;
					}

					if (e != VSConstants.S_OK)
					{
						ErrorCode(e, "ForEachDocument: enum next failed");
						break;
					}


					var cookie = cookies[0];

					if (cookie == VSConstants.VSCOOKIE_NIL)
					{
						// shouldn't happen
						Trace("  . nil cookie");
						continue;
					}

					var flags = Package.Instance.RDT4.GetDocumentFlags(cookie);
					const uint Pending = (uint)_VSRDTFLAGS4.RDT_PendingInitialization;

					if ((flags & Pending) != 0)
					{
						// document not initialized yet, skip it
						Trace("  . {0} pending", cookie);
						continue;
					}

					var d = VSDocument.DocumentFromCookie(cookie);
					if (d == null)
					{
						var mk = Package.Instance.RDT4.GetDocumentMoniker(cookie);
						Trace("  . {0} no document ({1})", cookie, mk);

						// GetRunningDocumentsEnum() enumerates all sorts of stuff
						// that are not documents, like the project files, even the
						// .sln file; all of those return null here, so they can
						// be safely ignored
						continue;
					}

					var wf = VSDocument.WindowFrameFromDocument(d);
					if (wf == null)
					{
						// this seems to happen for documents that haven't loaded
						// yet, they should get picked up by
						// DocumentEventHandlers.OnBeforeDocumentWindowShow later
						Trace("  . {0} no frame ({1})", cookie, d.FullName);

						continue;
					}

					Trace("  . {0} ok ({1})", cookie, d.FullName);

					list.Add(new VSDocument(d));
				}

				return list;
			}
		}
	}

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


	public class VSProject : IProject
	{
		// built-in projects that can be ignored
		//
		private static readonly List<string> BuiltinProjects = new List<string>()
		{
			EnvDTE.Constants.vsProjectKindMisc,
			EnvDTE.Constants.vsProjectKindSolutionItems,
			EnvDTE.Constants.vsProjectKindUnmodeled
		};

		private readonly EnvDTE.Project p;

		public VSProject(EnvDTE.Project p)
		{
			this.p = p;
		}

		public string Name
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				return p.Name;
			}
		}

		public bool IsBuiltIn
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				var k = p?.Kind;
				if (k == null)
					return false;

				return BuiltinProjects.Contains(k);
			}
		}
	}


	public class VSTreeItem : LoggingContext, ITreeItem
	{
		private readonly IVsHierarchy h;
		private readonly uint id;

		public VSTreeItem(
			IVsHierarchy h, uint id = (uint)VSConstants.VSITEMID.Root)
		{
			this.h = h;
			this.id = id;
		}

		protected override string LogPrefix()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			return "TreeItem " + DebugName;
		}

		public string Name
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				var e = h.GetProperty(
					id, (int)__VSHPROPID.VSHPROPID_Name,
					out var name);

				if (e != VSConstants.S_OK || !(name is string))
				{
					ErrorCode(e, "can't get itemid name");
					return null;
				}

				return (string)name;
			}
		}

		public ITreeItem Parent
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				var e = h.GetProperty(
					id, (int)__VSHPROPID.VSHPROPID_Parent,
					out var pidObject);

				// for whatever reason, VSHPROPID_Parent returns an int instead of
				// a uint

				if (e != VSConstants.S_OK || !(pidObject is int))
				{
					ErrorCode(e, "can't get parent item");
					return null;
				}

				var pid = (uint)(int)pidObject;
				if (pid == (uint)VSConstants.VSITEMID.Nil)
				{
					// no parent
					return null;
				}

				return new VSTreeItem(h, pid);
			}
		}

		// returns whether the given item is any type of folder; note that
		// some items may return true even if they're not actually folders,
		// see Variables.FolderPath()
		//
		public bool IsFolder
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				return GetIsFolder(h, id);
			}
		}

		// used for logging
		//
		public string DebugName
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				return MakeDebugName(Hierarchy, id);
			}
		}

		public static string MakeDebugName(
			IVsHierarchy h, uint id=(uint)VSConstants.VSITEMID.Root)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (h == null)
				return "(null hierarchy)";

			var e = h.GetCanonicalName(id, out var cn);

			if (e == VSConstants.S_OK)
			{
				if (cn is string s)
				{
					if (s.Length > 0)
						return s;
				}
			}

			// failed, try the name property

			e = h.GetProperty(
				id, (int)__VSHPROPID.VSHPROPID_Name, out var no);

			if (e == VSConstants.S_OK)
			{
				if (no is string s)
				{
					if (s.Length > 0)
						return s;
				}
			}

			// whatever
			return "?";
		}

		public static bool GetIsFolder(
			IVsHierarchy h, uint id=(uint)VSConstants.VSITEMID.Root)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var e = h.GetGuidProperty(
				id, (int)__VSHPROPID.VSHPROPID_TypeGuid,
				out var type);

			if (e != VSConstants.S_OK || type == null)
			{
				Main.Instance.Logger.ErrorCode(e, "can't get TypeGuid");
				return false;
			}

			// ignore anything but folders
			if (type == VSConstants.ItemTypeGuid.PhysicalFolder_guid)
				return true;

			if (type == VSConstants.ItemTypeGuid.VirtualFolder_guid)
				return true;

			return false;
		}

		public IVsHierarchy Hierarchy
		{
			get
			{
				return h;
			}
		}
	}


	public class VSLogger : ILogger
	{
		// output pane
		private IVsOutputWindowPane pane = null;

		public void Output(string s)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// make sure the pane exists
			if (!CheckPane())
				return;

			pane.OutputString(s + "\n");
		}

		// creates the pane in the output window if necessary, returns whether
		// the pane is available
		//
		private bool CheckPane()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// create only once
			if (pane != null)
				return true;

			// try getting the output window
			var w = Package.Instance.OutputWindow;
			if (w == null)
				return false;

			// create a new pane for this extension; this adds an entry in the
			// "show output from" combo box
			var guid = new System.Guid(Strings.ExtensionGuid);
			w.CreatePane(
				ref guid, Strings.ExtensionName,
				Convert.ToInt32(true), Convert.ToInt32(false));

			// try to get the pane that was just created
			w.GetPane(guid, out pane);
			if (pane == null)
				return false;

			return true;
		}
	}
}
