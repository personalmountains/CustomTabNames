using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;

namespace CustomTabNames
{
	public interface ISolution
	{
		List<ITreeItem> ProjectItems
		{
			get;
		}

		List<IDocument> Documents
		{
			get;
		}

		bool HasSingleProject
		{
			get;
		}
	}

	public interface IDocument
	{
		string Path
		{
			get;
		}

		string Name
		{
			get;
		}

		IProject Project
		{
			get;
		}

		ITreeItem TreeItem
		{
			get;
		}

		void SetCaption(string s);
		void ResetCaption();
	}

	public interface IProject
	{
		string Name
		{
			get;
		}

		bool IsBuiltIn
		{
			get;
		}
	}

	public interface ITreeItem
	{
		string Name
		{
			get;
		}

		ITreeItem Parent
		{
			get;
		}

		bool IsFolder
		{
			get;
		}

		string DebugName
		{
			get;
		}
	}


	public class VSSolution : ISolution
	{
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
						Package.Instance.Logger.ErrorCode(
							e, "HasSingleProject: failed to get project count");

						return false;
					}

					int i = (int)o;
					return (i == 1);
				}
				catch (Exception e)
				{
					Package.Instance.Logger.Error(
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
					Package.Instance.Logger.ErrorCode(
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
						Package.Instance.Logger.ErrorCode(
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
					Package.Instance.Logger.ErrorCode(
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
						Package.Instance.Logger.ErrorCode(
							e, "ForEachDocument: enum next failed");

						break;
					}


					var cookie = cookies[0];

					if (cookie == VSConstants.VSCOOKIE_NIL)
					{
						// shouldn't happen
						Package.Instance.Logger.Trace("  . nil cookie");
						continue;
					}

					var flags = Package.Instance.RDT4.GetDocumentFlags(cookie);
					const uint Pending = (uint)_VSRDTFLAGS4.RDT_PendingInitialization;

					if ((flags & Pending) != 0)
					{
						// document not initialized yet, skip it
						Package.Instance.Logger.Trace("  . {0} pending", cookie);
						continue;
					}

					var d = Utilities.DocumentFromCookie(cookie);
					if (d == null)
					{
						var mk = Package.Instance.RDT4.GetDocumentMoniker(cookie);
						Package.Instance.Logger.Trace(
							"  . {0} no document ({1})", cookie, mk);

						// GetRunningDocumentsEnum() enumerates all sorts of stuff
						// that are not documents, like the project files, even the
						// .sln file; all of those return null here, so they can
						// be safely ignored
						continue;
					}

					var wf = Utilities.WindowFrameFromDocument(d);
					if (wf == null)
					{
						// this seems to happen for documents that haven't loaded
						// yet, they should get picked up by
						// DocumentEventHandlers.OnBeforeDocumentWindowShow later
						Package.Instance.Logger.Trace(
							"  . {0} no frame ({1})", cookie, d.FullName);

						continue;
					}

					Package.Instance.Logger.Trace(
						"  . {0} ok ({1})", cookie, d.FullName);

					list.Add(new VSDocument(d));
				}

				return list;
			}
		}
	}

	public class VSDocument : IDocument
	{
		private readonly EnvDTE.Document d;
		private IVsWindowFrame f;

		public VSDocument(EnvDTE.Document d, IVsWindowFrame f = null)
		{
			this.d = d;
			this.f = f;
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

				if (!Utilities.ItemIDFromDocument(d, out var h, out var id))
					return null;

				return new VSTreeItem(h, id);
			}
		}

		public void SetCaption(string s)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Package.Instance.Logger.Log("setting to {0}", s);

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

			Package.Instance.Logger.Log("resetting to {0}", d.Name);
			SetCaption(d.Name);
		}

		public IVsWindowFrame WindowFrame
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (f == null)
					f = Utilities.WindowFrameFromDocument(d);

				return f;
			}
		}
	}


	public class VSProject : IProject
	{
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
				return Utilities.IsBuiltInProject(p);
			}
		}
	}


	public class VSTreeItem : ITreeItem
	{
		private readonly IVsHierarchy h;
		private readonly uint id;

		public VSTreeItem(
			IVsHierarchy h, uint id=(uint)VSConstants.VSITEMID.Root)
		{
			this.h = h;
			this.id = id;
		}

		public string Name
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				return Utilities.ItemName(h, id);
			}
		}

		public ITreeItem Parent
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (!Utilities.ParentItemID(h, id, out var pid))
					return null;

				return new VSTreeItem(h, pid);
			}
		}

		public bool IsFolder
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				return Utilities.ItemIsFolder(h, id);
			}
		}

		public string DebugName
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				return Utilities.DebugHierarchyName(h, id);
			}
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
