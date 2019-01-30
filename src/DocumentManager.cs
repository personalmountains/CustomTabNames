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

		// fired when projects or filters are added, removed or renamed
		//
		public delegate void ContainersChangedHandler();
		public event ContainersChangedHandler ContainersChanged;

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
			solHandlers.ContainerNameChanged += OnContainerNameChanged;
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

		// calls f() for each opened document
		//
		public static void ForEachDocument(Action<DocumentWrapper> f)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var e = Package.Instance.RDT.GetRunningDocumentsEnum(
				out var enumerator);

			if (e != VSConstants.S_OK)
			{
				Logger.ErrorCode(e, "GetRunningDocumentsEnum failed");
				return;
			}

			uint[] cookie = new uint[1] { VSConstants.VSCOOKIE_NIL };
			enumerator.Reset();

			while (true)
			{
				e = enumerator.Next(1, cookie, out var fetched);

				// done
				if (e == VSConstants.S_FALSE || fetched != 1)
					break;

				if (e != VSConstants.S_OK)
				{
					Logger.ErrorCode(
						e, "ForEachDocument enumerator next failed");

					break;
				}

				if (cookie[0] == VSConstants.VSCOOKIE_NIL)
					continue;

				var d = Utilities.DocumentFromCookie(cookie[0]);
				if (d == null)
					continue;

				var wf = Utilities.WindowFrameFromDocument(d);
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

			var e = Package.Instance.Solution.GetProjectEnum(
				(uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION,
				ref guid, out var enumerator);

			if (e != VSConstants.S_OK)
			{
				Logger.ErrorCode(e, "GetProjectEnum failed");
				return;
			}

			IVsHierarchy[] hierarchy = new IVsHierarchy[1] { null };
			enumerator.Reset();

			while (true)
			{
				e = enumerator.Next(1, hierarchy, out var fetched);

				// done
				if (e == VSConstants.S_FALSE || fetched != 1)
					break;

				if (e != VSConstants.S_OK)
				{
					Logger.ErrorCode(
						e, "ForEachProjectHierarchy enumerator next failed");

					break;
				}

				f(hierarchy[0]);
			}
		}

		public static bool HasSingleProject()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			try
			{
				Package.Instance.Solution.GetProperty(
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

			projectCountTimer.Start(1000, () =>
			{
				ContainersChanged?.Invoke();
			});
		}

		// fired when filters or projects get renamed
		//
		private void OnContainerNameChanged()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Trace("container name changed");
			ContainersChanged?.Invoke();
		}
	}
}
