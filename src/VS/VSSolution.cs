using System;
using System.Collections.Generic;
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
					var e = Package.Instance.Solution.GetProperty(
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
				var e = Package.Instance.Solution.GetProjectEnum(
					(uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION,
					ref guid, out var enumerator);

				if (e != VSConstants.S_OK)
				{
					ErrorCode(
						e, "ForEachProjectHierarchy: GetProjectEnum failed");

					return list;
				}

				// will store one hierarchy at a time, but Next() still requires
				// an array
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
	}
}
