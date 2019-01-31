using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Collections.Generic;

namespace CustomTabNames
{
	// solution events are fired when projects are added and removed
	//
	// for each project a HierarchyEventHandlers is created and attached,
	// because hierarchy events are per-project, not global
	//
	public sealed class SolutionEventHandlers : SolutionEventHandlersBase
	{
		public event ProjectHandler ProjectAdded, ProjectRemoved;
		public event ProjectHandler ProjectRenamed;
		public event FolderHandler FolderRenamed;
		public event DocumentHandler DocumentRenamed;

		// registration cookie, used in Unregister()
		private uint cookie = VSConstants.VSCOOKIE_NIL;

		// hierarchy handlers are per-project, this remembers them
		private readonly Dictionary<string, HierarchyEventHandlers>
			hierarchyHandlers =
				new Dictionary<string, HierarchyEventHandlers>();

		// see OnBeforeCloseProject()
		private readonly MainThreadTimer projectCloseTimer
			= new MainThreadTimer();


		protected override string LogPrefix()
		{
			return "SolutionEventHandlers";
		}

		// register for solution events and walks all current projects to
		// register for their own hierarchy events: this is necessary because
		// the extension might initialize _after_ projects are loaded, and
		// therefore won't receive the corresponding events
		//
		public void Register()
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("registering for events");

			var e = Package.Instance.Solution
				.AdviseSolutionEvents(this, out cookie);

			if (e != VSConstants.S_OK)
			{
				ErrorCode(e, "AdviseSolutionEvents() failed");

				// don't bother with the rest, there's a major problem
				return;
			}

			// go through every loaded project and register for hierarchy
			// events; this handles cases where the extension was loaded after
			// the current solution
			Utilities.ForEachProjectHierarchy((h) =>
			{
				AddProjectHierarchy(h);
			});
		}

		// unregisters from solution events
		//
		public void Unregister()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// unregister all per-project handlers
			foreach (var hh in hierarchyHandlers)
				hh.Value.Unregister();

			hierarchyHandlers.Clear();

			Trace("unregistering from events");

			if (cookie == VSConstants.VSCOOKIE_NIL)
			{
				Error("cookie is nil");
				return;
			}

			var e = Package.Instance.Solution.UnadviseSolutionEvents(cookie);
			if (e != VSConstants.S_OK)
				ErrorCode(e, "UnadviseSolutionEvents() failed");
		}


		// fired once per project in the current solution when a project is
		// added or opened
		//
		public override int OnAfterOpenProject(
			IVsHierarchy hierarchy, int added)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("OnAfterOpenProject");

			// this is fired once per project when loading a solution, but also
			// when a project is added to a solution that's already loaded
			//
			// hierarchy events need to be registered in both cases, but the
			// ProjectAdded event doesn't need to be fired while loading a
			// solution

			// register for hierarchy events for this project and add it to the
			// list
			AddProjectHierarchy(hierarchy);

			if (added != 0)
				ProjectAdded?.Invoke(hierarchy);

			return VSConstants.S_OK;
		}

		// fired just _before_ a project is removed from the solution or when
		// a solution is currently being closed; there is no corresponding
		// OnAfterCloseProject, which is unfortunate, so a timer has to be
		// used to fire the event once internal stuff has been updated
		//
		public override int OnBeforeCloseProject(
			IVsHierarchy hierarchy, int removed)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Trace("OnBeforeCloseProject");

			// this is fired once per project when closing a solution, but also
			// when a project is removed from a solution that's already loaded
			//
			// hierarchy events need to be unregistered in both cases, but the
			// ProjectRemoved event doesn't need to be fired while unloading a
			// solution

			// unregister for hierarchy events and remove it from the list
			RemoveProjectHierarchy(hierarchy);

			if (removed != 0)
				return VSConstants.S_OK;

			// this is an "on before" handler, and so the project count hasn't
			// been updated yet
			//
			// there is no On*After*CloseProject
			//
			// it's unclear what kind of delay there is between
			// OnBeforeCloseProject and the actual update of the project count,
			// but one second works for now
			//
			// in any case, the close can probably still be canceled by the
			// user if there are unsaved changes, so this may be a false
			// positive

			projectCloseTimer.Start(1000, () =>
			{
				ProjectRemoved?.Invoke(hierarchy);
			});

			return VSConstants.S_OK;
		}

		// called once per project, registers for project hierarchy events and
		// adds the handler to the list
		//
		private void AddProjectHierarchy(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// the canonical name of a project hierarchy object is unique, so
			// it is used as a key in the dictionary
			var e = h.GetCanonicalName(
				(uint)VSConstants.VSITEMID.Root, out var cn);

			if (e != VSConstants.S_OK)
			{
				ErrorCode(e, "AddProjectHierarchy: GetCanonicalName failed");
				return;
			}

			Log("AddProjectHierarchy: adding {0} to list", cn);

			if (hierarchyHandlers.ContainsKey(cn))
			{
				Error("AddProjectHierarchy: list already contains {0}", cn);
				return;
			}

			var hh = new HierarchyEventHandlers(h);

			// registering for hierarchy events on this project
			hh.Register();

			// forwarding
			hh.ProjectRenamed += ProjectRenamed;
			hh.FolderRenamed += FolderRenamed;
			hh.DocumentRenamed += DocumentRenamed;

			hierarchyHandlers.Add(cn, hh);
		}

		private void RemoveProjectHierarchy(IVsHierarchy h)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// the canonical name of a project hierarchy object is unique, so
			// it is used as a key in the dictionary
			var e = h.GetCanonicalName(
				(uint)VSConstants.VSITEMID.Root, out var cn);

			if (e != VSConstants.S_OK)
			{
				ErrorCode(e, "RemoveProjectHierarchy: GetCanonicalName failed");
				return;
			}

			Log("RemoveProjectHierarchy: removing {0} from list", cn);

			hierarchyHandlers.TryGetValue(cn, out var hh);
			if (hh == null)
			{
				Error("RemoveProjectHierarchy: {0} not in list", cn);
				return;
			}

			// unregistering from events
			hh.Unregister();

			hierarchyHandlers.Remove(cn);
		}
	}
}
