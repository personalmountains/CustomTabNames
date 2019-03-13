using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CustomTabNames
{
	public class VSTreeItem : LoggingContext, ITreeItem
	{
		private const uint RootID = (uint)VSConstants.VSITEMID.Root;
		public IVsHierarchy Hierarchy { get; private set; }
		public uint ID { get; private set; }

		public VSTreeItem(IVsHierarchy h, uint id=RootID)
		{
			Hierarchy = h;
			ID = id;
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

				var e = Hierarchy.GetProperty(
					ID, (int)__VSHPROPID.VSHPROPID_Name,
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

				var e = Hierarchy.GetProperty(
					ID, (int)__VSHPROPID.VSHPROPID_Parent,
					out var pidObject);

				// for whatever reason, VSHPROPID_Parent returns an int instead
				// of a uint

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

				return new VSTreeItem(Hierarchy, pid);
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

				if (!GetIsFolder(Hierarchy, ID))
					return false;

				// VS reports that the item is a folder, but it might not be

				// there are two important magic things:
				//
				//   1) files that are considered external dependencies
				//      (somehow included by files in the project) are put
				//      in an "External dependencies" folder, which is a
				//      direct child of the project
				//
				//      this folder doesn't seem to have any recognizable type,
				//      it's just a guid
				//
				//   2) other files that are not dependencies are put in a
				//      a separate project called "Miscellaneous files", which
				//      reports itself as being a folder instead of a project
				//
				// so both the external dependencies folder and the misc files
				// projects are magic
				//
				// however, if "hide external dependencies folder" is enabled,
				// the magic folder above is still a child of the project in the
				// project hierarchy, but checking EnvDTE.Document.ProjectItem
				// on any file returns null (see also VSDocument.Project)
				//

				// so the first thing to check is whether the ID is RootID, in
				// which case this item is actual a project, not a folder, since
				// only projects can be roots
				if (ID == RootID)
					return false;

				// even if the item is not the root and its type was a folder,
				// it could be the magic external dependencies folder
				//
				// unfortunately, there doesn't seem to be a way of figuring
				// this out, except to check if the name is a guid and interpret
				// that as the "external dependencies" folder
				//
				// this will generate false positives, hopefully not too
				// many people use guids as folder names
				if (System.Guid.TryParse(Name, out var r))
					return false;

				// VS says it's a folder, it's not a project, and its name is
				// not a guid, assume it's actually a folder
				return true;
			}
		}

		// used for logging
		//
		public string DebugName
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				return MakeDebugName(Hierarchy, ID);
			}
		}

		public static string MakeDebugName(IVsHierarchy h, uint id= RootID)
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

		public static bool GetIsFolder(IVsHierarchy h, uint id=RootID)
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
	}
}
