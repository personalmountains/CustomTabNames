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
	public sealed class DocumentWrapper : LoggingContext
	{
		public Document Document { get; private set; }
		private IVsWindowFrame Frame { get; set; }

		protected override string LogPrefix()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			return "DocumentWrapper " + Document.FullName;
		}

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

			Log("setting to {0}", s);

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

			Log("resetting to {0}", Document.Name);

			// todo: it'd be nice to set the caption back to a default value
			// instead of hardcoding the name, but there doesn't seem to be a
			// way to do that
			SetCaption(Document.Name);
		}
	}


	// manages the various events for opening documents and windows, and fires
	// DocumentChanged when they do
	//
	public sealed class DocumentManager : LoggingContext
	{
		private readonly DocumentEventHandlers docHandlers
			 = new DocumentEventHandlers();

		private readonly SolutionEventHandlers solHandlers
			= new SolutionEventHandlers();

		// fired every time a document changes in a way that may require
		// fixing the caption
		public delegate void DocumentChangedHandler(DocumentWrapper d);
		public event DocumentChangedHandler DocumentChanged;

		// fired when projects or filters are added, removed or renamed
		//
		public delegate void ContainersChangedHandler();
		public event ContainersChangedHandler ContainersChanged;


		// built-in projects that can be ignored
		//
		private static readonly List<string> BuiltinProjects = new List<string>()
		{
			EnvDTE.Constants.vsProjectKindMisc,
			EnvDTE.Constants.vsProjectKindSolutionItems,
			EnvDTE.Constants.vsProjectKindUnmodeled
		};


		protected override string LogPrefix()
		{
			return "DocumentManager";
		}


		public DocumentManager()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			solHandlers.ProjectAdded += OnProjectAdded;
			solHandlers.ProjectRemoved += OnProjectRemoved;
			solHandlers.ProjectRenamed += OnProjectRenamed;

			solHandlers.FolderRenamed += OnFolderRenamed;

			solHandlers.DocumentRenamed += OnDocumentRenamed;
			docHandlers.DocumentRenamed += OnDocumentRenamed;
			docHandlers.DocumentOpened += OnDocumentOpened;
		}

		// starts the manager
		//
		public void Start()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			SetEvents(true);
		}

		// stops the manager
		//
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

			// getting enumerator
			var e = Package.Instance.RDT.GetRunningDocumentsEnum(
				out var enumerator);

			if (e != VSConstants.S_OK)
			{
				Logger.ErrorCode(
					e, "ForEachDocument: GetRunningDocumentsEnum failed");

				return;
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
					Logger.ErrorCode(e, "ForEachDocument: enum next failed");
					break;
				}


				var cookie = cookies[0];

				if (cookie == VSConstants.VSCOOKIE_NIL)
				{
					// shouldn't happen
					continue;
				}

				var d = Utilities.DocumentFromCookie(cookie);
				if (d == null)
					continue;

				var wf = Utilities.WindowFrameFromDocument(d);
				if (wf == null)
				{
					// this seems to happen for documents that haven't loaded
					// yet, they should get picked up by
					// DocumentEventHandlers.OnBeforeDocumentWindowShow later
					Logger.Log(
						"ForEachDocument: skipping {0}, no frame", d.FullName);

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

			// getting enumerator
			var e = Package.Instance.Solution.GetProjectEnum(
				(uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION,
				ref guid, out var enumerator);

			if (e != VSConstants.S_OK)
			{
				Logger.ErrorCode(
					e, "ForEachProjectHierarchy: GetProjectEnum failed");

				return;
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
					Logger.ErrorCode(
						e, "ForEachProjectHierarchy: enum next failed");

					break;
				}


				var h = hierarchies[0];

				if (h == null)
				{
					// shouldn't happen
					continue;
				}

				f(h);
			}
		}

		// returns whether the current solution only has one project in it
		//
		public static bool HasSingleProject()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			try
			{
				var e = Package.Instance.Solution.GetProperty(
					(int)__VSPROPID.VSPROPID_ProjectCount, out var o);

				if (e != VSConstants.S_OK || !(o is int))
				{
					Logger.ErrorCode(
						e, "HasSingleProject: failed to get project count");

					return false;
				}

				int i = (int)o;
				return (i == 1);
			}
			catch(Exception e)
			{
				Logger.Error(
					"HasSingleProject: failed to get project count, {0}",
					e.Message);

				return false;
			}
		}

		// returns whether the given document is in a builtin project
		//
		public static bool IsInBuiltinProject(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var k = d?.ProjectItem?.ContainingProject?.Kind;
			if (k == null)
				return false;

			return BuiltinProjects.Contains(k);
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

		private void OnProjectAdded(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("project {0} added", Utilities.DebugHierarchyName(h));

			ContainersChanged?.Invoke();
		}

		private void OnProjectRemoved(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("project {0} removed", Utilities.DebugHierarchyName(h));

			ContainersChanged?.Invoke();
		}

		private void OnProjectRenamed(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("project {0} renamed", Utilities.DebugHierarchyName(h));

			ContainersChanged?.Invoke();
		}

		private void OnFolderRenamed(IVsHierarchy h, uint item)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("folder {0} renamed", Utilities.DebugHierarchyName(h, item));

			ContainersChanged?.Invoke();
		}

		private void OnDocumentRenamed(IVsHierarchy h, uint item)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("document {0} renamed", Utilities.DebugHierarchyName(h, item));
			OnDocumentChanged(h, item);
		}

		private void OnDocumentOpened(IVsHierarchy h, uint item)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("document {0} opened", Utilities.DebugHierarchyName(h, item));
			OnDocumentChanged(h, item);
		}

		// handles both renamed and opened
		//
		private void OnDocumentChanged(IVsHierarchy h, uint item)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var d = Utilities.DocumentFromItemID(h, item);
			if (d == null)
			{
				Error(
					"document {0} has no document object",
					Utilities.DebugHierarchyName(h, item));

				return;
			}

			var wf = Utilities.WindowFrameFromDocument(d);
			if (wf == null)
			{
				Error(
					"document {0} has no frame",
					Utilities.DebugHierarchyName(h, item));

				return;
			}

			DocumentChanged?.Invoke(new DocumentWrapper(d, wf));
		}
	}
}
