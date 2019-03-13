using System;
using System.Collections.Generic;
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

				var n = TentativeProjectCount();

				// simple case
				if (n == 0)
					return true;

				// VSPROPID_ProjectCount can include virtual projects,
				// such as misc files
				//
				// filtering those projects requires getting all the projects
				// and checking if they're built-in, but that can take a long
				// time when many projects are loaded
				//
				// assume that anything over 10 has at least 2 real projects;
				// lower values can afford to loop
				if (n >= 10)
					return false;

				int realN = 0;

				ForEachProject(__VSENUMPROJFLAGS.EPF_ALLINSOLUTION, (h) =>
				{
					ThreadHelper.ThrowIfNotOnUIThread();

					var e = h.GetProperty(
						(uint)VSConstants.VSITEMID.Root,
						(int)__VSHPROPID.VSHPROPID_ExtObject, out var po);

					if (e != VSConstants.S_OK)
					{
						ErrorCode(
							e, "HasSingleProject: failed to get " +
							"ExtObject for project hierarchy");

						return;
					}

					if (po is Project p)
					{
						var vsp = new VSProject(p);

						if (!vsp.IsBuiltIn)
							++realN;
					}
					else
					{
						Error("HasSingleProject: item is not a project");
					}
				});

				return (realN <= 1);

			}
		}

		private int TentativeProjectCount()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			try
			{
				var e = Package.Instance.Solution.GetProperty(
					(int)__VSPROPID.VSPROPID_ProjectCount, out var o);

				if (e != VSConstants.S_OK || !(o is int))
				{
					ErrorCode(
						e, "TentativeProjectCount: " +
						"failed to get project count");

					return -1;
				}

				return (int)o;
			}
			catch (Exception e)
			{
				Error(
					"HasSingleProject: failed to get project count, {0}",
					e.Message);

				return -1;
			}
		}

		public List<ITreeItem> ProjectItems
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				var list = new List<ITreeItem>();

				ForEachProject(__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION, (p) =>
				{
					list.Add(new VSTreeItem(p));
				});

				return list;
			}
		}

		public List<IDocument> Documents
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				var list = new List<IDocument>();
				var RDT = Package.Instance.RDT;
				var RDT4 = Package.Instance.RDT4;

				// getting enumerator
				var e = RDT.GetRunningDocumentsEnum(out var enumerator);
				if (e != VSConstants.S_OK)
				{
					ErrorCode(
						e, "ForEachDocument: GetRunningDocumentsEnum failed");

					return list;
				}

				// one at a time
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

					var flags = RDT4.GetDocumentFlags(cookie);
					const uint Pending = (uint)_VSRDTFLAGS4
						.RDT_PendingInitialization;

					if ((flags & Pending) != 0)
					{
						// document not initialized yet, skip it
						Trace("  . {0} pending", cookie);
						continue;
					}

					var d = VSDocument.DocumentFromCookie(cookie);
					if (d == null)
					{
						var mk = RDT4.GetDocumentMoniker(cookie);
						Trace("  . {0} no document ({1})", cookie, mk);

						// GetRunningDocumentsEnum() enumerates all sorts of
						// stuff that are not documents, like the project files,
						// even the .sln file; all of those return null here,
						// so they can be safely ignored
						continue;
					}

					var wf = VSDocument.WindowFrameFromDocument(d);
					if (wf == null)
					{
						// this seems to happen for documents that haven't
						// loaded yet, they should get picked up by
						// DocumentEventHandlers.OnBeforeDocumentWindowShow
						// later
						Trace("  . {0} no frame ({1})", cookie, d.FullName);
						continue;
					}

					Trace("  . {0} ok ({1})", cookie, d.FullName);
					list.Add(new VSDocument(d));
				}

				return list;
			}
		}

		private void ForEachProject(
			__VSENUMPROJFLAGS type, Action<IVsHierarchy> f)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var list = new List<IVsHierarchy>();

			Guid guid = Guid.Empty;

			// getting enumerator
			var e = Package.Instance.Solution.GetProjectEnum(
				(uint)type, ref guid, out var enumerator);

			if (e != VSConstants.S_OK)
			{
				ErrorCode(e, "GetProjectItems: GetProjectEnum failed");
				return;
			}

			// one at a time
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
					ErrorCode(e, "GetProjectItems: enum next failed");
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
	}
}
