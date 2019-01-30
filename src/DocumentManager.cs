using System;
using System.Collections.Generic;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;

namespace CustomTabNames
{
	// wraps a Document and a IVsWindowFrame
	//
	// the Document is used to feed information to variables, like path and
	// filters; the IVsWindowFrame is required to set the actual caption
	//
	// the frame might be null in cases when the document is opened, but no
	// window has been created yet
	//
	public sealed class DocumentWrapper
	{
		public Document Document { get; private set; }
		private IVsWindowFrame Frame { get; set; }

		public DocumentWrapper(Document d, IVsWindowFrame f)
		{
			Document = d;
			Frame = f;
		}

		// sets the caption of this document to the given string
		//
		public void SetCaption(string s)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Logger.Log("setting {0} to {1}", Document.FullName, s);

			// the visible caption is made of a combination of the EditorCaption
			// and OwnerCaption; setting the EditorCaption to null makes sure
			// the caption can be controlled uniquely by OwnerCaption
			Frame.SetProperty((int)VsFramePropID.EditorCaption, null);
			Frame.SetProperty((int)VsFramePropID.OwnerCaption, s);
		}

		// resets the caption of this document to the default value
		//
		public void ResetCaption()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Logger.Log(
				"resetting {0} to {1}", Document.FullName, Document.Name);

			// todo: it'd be nice to set the caption back to a default value
			// instead of hardcoding the name, but there doesn't seem to be a
			// way to do that
			SetCaption(Document.Name);
		}
	}

	// manages the various events for opening documents and windows, and fires
	// DocumentChanged when they do
	//
	public sealed class DocumentManager : IDisposable
	{
		private readonly DocumentEventHandlers docHandlers;
		private readonly SolutionEventHandlers solHandlers;

		// fired every time a document changes in a way that may require
		// fixing the caption
		public delegate void DocumentChangedHandler(DocumentWrapper d);
		public event DocumentChangedHandler DocumentChanged;

		// fired when projects are added, removed or renamed
		//
		public delegate void ProjectsChangedHandler();
		public event ProjectsChangedHandler ProjectsChanged;

		private readonly MainThreadTimer projectCountTimer
			= new MainThreadTimer();

		// built-in projects that can be ignored
		//
		private static readonly List<string> BuiltinProjects = new List<string>()
		{
			EnvDTE.Constants.vsProjectKindMisc,
			EnvDTE.Constants.vsProjectKindSolutionItems,
			EnvDTE.Constants.vsProjectKindUnmodeled
		};


		public DocumentManager()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			this.docHandlers = new DocumentEventHandlers();
			this.solHandlers = new SolutionEventHandlers();

			docHandlers.DocumentOpened += OnDocumentChanged;
			docHandlers.DocumentRenamed += OnDocumentChanged;
			solHandlers.ProjectCountChanged += OnProjectCountChanged;
			solHandlers.DocumentMoved += OnDocumentChanged;
		}

		public void Dispose()
		{
			projectCountTimer.Dispose();
		}

		// starts the manager
		//
		public void Start()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			SetEvents(true);
		}

		// stops the manager
		public void Stop()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			SetEvents(false);
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
				CustomTabNames.Instance,
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
				Logger.Error(
					"DocumentFromID: GetProperty for extObject failed, {0}", e);

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

			var mk = CustomTabNames.Instance.RDT4.GetDocumentMoniker(cookie);

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

		// calls f() for each opened document
		//
		public static void ForEachDocument(Action<DocumentWrapper> f)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var e = CustomTabNames.Instance.RDT.GetRunningDocumentsEnum(
				out var enumerator);

			if (e != VSConstants.S_OK)
			{
				Logger.Error("GetRunningDocumentsEnum failed, {0}", e);
				return;
			}

			uint[] cookie = new uint[1] { VSConstants.VSCOOKIE_NIL };
			enumerator.Reset();

			while (true)
			{
				e = enumerator.Next(1, cookie, out var fetched);
				if (e != VSConstants.S_OK)
					break;

				if (fetched != 1)
					break;

				if (cookie[0] == VSConstants.VSCOOKIE_NIL)
					continue;

				var d = DocumentFromCookie(cookie[0]);
				if (d == null)
					continue;

				var wf = WindowFrameFromDocument(d);
				if (wf == null)
				{
					// this seems to happen for documents that haven't loaded
					// yet, they should get picked up by
					// DocumentEventHandlers.OnBeforeDocumentWindowShow later
					Logger.Log("skipping {0}, no frame", d.FullName);
					continue;
				}

				f(new DocumentWrapper(d, wf));
			}
		}

		// calls f() for each loaded project
		//
		public static void ForEachProjectHierarchy(Action<IVsHierarchy> f)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			Guid guid = Guid.Empty;

			var e = CustomTabNames.Instance.Solution.GetProjectEnum(
				(uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION,
				ref guid, out var enumerator);

			if (e != VSConstants.S_OK)
			{
				Logger.Error("GetProjectEnum failed, {0}", e);
				return;
			}

			IVsHierarchy[] hierarchy = new IVsHierarchy[1] { null };
			enumerator.Reset();

			while (true)
			{
				e = enumerator.Next(1, hierarchy, out var fetched);
				if (e != VSConstants.S_OK)
					break;

				if (fetched != 1)
					break;

				f(hierarchy[0]);
			}
		}

		public static bool HasSingleProject()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			try
			{
				CustomTabNames.Instance.Solution.GetProperty(
					(int)__VSPROPID.VSPROPID_ProjectCount, out var o);

				int i = (int)o;
				return (i == 1);
			}
			catch(Exception e)
			{
				Logger.Error("failed to get project count, {0}", e.Message);
				return false;
			}
		}

		public static bool IsInBuiltinProject(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var p = d?.ProjectItem?.ContainingProject;
			if (p == null)
				return false;

			return BuiltinProjects.Contains(p.Kind);
		}

		// either registers or unregisters the events
		//
		private void SetEvents(bool add)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (add)
			{
				docHandlers.Register();
				solHandlers.Register();
			}
			else
			{
				solHandlers.Unregister();
				docHandlers.Unregister();
			}
		}

		// fired when a document has been opened or renamed
		//
		private void OnDocumentChanged(DocumentWrapper d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Trace("document changed: {0}", d.Document.FullName);

			DocumentChanged?.Invoke(d);
		}

		// fired when a project was added or removed
		//
		private void OnProjectCountChanged()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Trace("project count changed");
			projectCountTimer.Start(1000, () => { ProjectsChanged?.Invoke(); });
		}
	}
}
