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

				if (GetIsFolder(Hierarchy, ID))
				{
					// sigh
					//
					// some of the built-in projects like miscellaneous items
					// seem to behave both as projects and folders
					//
					// the GetIsFolder() check above checks for physical/virtual
					// folders and works fine for regular projects, where the
					// root project item doesn't say it's a folder (cause it
					// ain't)
					//
					// but when an external file is opened, it gets put in this
					// magic miscellaneous items project, which _does_ report
					// its type as a virtual folder, event though it's the root
					// "project"
					if (ID != RootID)
						return true;
				}

				return false;
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
