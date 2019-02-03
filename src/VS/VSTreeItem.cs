using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CustomTabNames
{
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
			IVsHierarchy h, uint id = (uint)VSConstants.VSITEMID.Root)
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
			IVsHierarchy h, uint id = (uint)VSConstants.VSITEMID.Root)
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
}
