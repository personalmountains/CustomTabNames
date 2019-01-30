using EnvDTE;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CustomTabNames
{
	using Dict = Dictionary<string, Func<Document, string>>;

	public sealed class Variables
	{
		// maps variable names to functions
		//
		private static readonly Dict Dictionary = new Dict()
		{
			// variable names must be [a-zA-Z]
			{"ProjectName",    ProjectName},
			{"ParentDir",      ParentDir},
			{"Filename",       Filename},
			{"FullPath",       FullPath},
			{"FilterPath",     FilterPath},
			{"ParentFilter",   ParentFilter}
			// update Strings.OptionTemplateDescription when changing this list
		};


		private static Options Options
		{
			get
			{
				return Package.Instance.Options;
			}
		}


		// expands all variables on the given template
		//
		public static string Expand(Document d, string template)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			string s = template;

			// matches $(VariableName 'text') where 'text' is optional
			// [1] is "VariableName"
			// [2] is "text"
			var re = new Regex(@"\$\(\s*([a-zA-Z]+)\s*(?:'(.*?)')?\)");

			Logger.Trace(
				"making caption for {0} using template {1}",
				d.FullName, s);

			while (true)
			{
				var m = re.Match(s);
				if (!m.Success)
					break;

				var name = m.Groups[1].Value;
				var text = m.Groups[2].Value;

				var replacement = ExpandOne(d, name, text);
				s = ReplaceMatch(s, m, replacement);
			}

			Logger.Trace("  . caption is now {0}", s);

			return s;
		}

		// returns the expansion of variable 'name' with option text
		//
		private static string ExpandOne(Document d, string name, string text)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (Dictionary.TryGetValue(name, out var v))
			{
				// variable is in the dictionary

				// expand it
				string s = v(d);

				// don't append the text if the result was empty
				if (s.Length != 0)
					s += text;

				Logger.Trace("  . variable {0} replaced by '{1}'", name, s);

				return s;
			}
			else
			{
				// not found in the dictionary, put the variable name verbatim
				// to notify the user
				Logger.Trace("  . variable {0} not found", name);
				return name;
			}
		}

		// replaces the given match in the string
		//
		private static string ReplaceMatch(string s, Match m, string repl)
		{
			// before the match
			var ns = s.Substring(0, m.Index);

			// the replacement
			ns += repl;

			// after the match
			ns += s.Substring(m.Index + m.Length);

			return ns;
		}


		// returns the name of the document's project, or an empty string; this
		// can happen for external files, includes, etc.
		//
		public static string ProjectName(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (Options.IgnoreSingleProject)
			{
				if (DocumentManager.HasSingleProject())
					return "";
			}

			if (Options.IgnoreBuiltinProjects)
			{
				if (DocumentManager.IsInBuiltinProject(d))
					return "";
			}

			var p = d?.ProjectItem?.ContainingProject;
			if (p == null)
				return "";

			return p.Name;
		}

		// returns the parent directory of the document's file with a slash at
		// the end, or an empty string; this happens for files in the root
		//
		public static string ParentDir(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var parts = SplitPath(d.FullName);
			if (parts.Length < 2)
				return "";

			return parts[parts.Length - 2] + "/";
		}

		// returns the last component of the full path, or an empty string
		//
		public static string Filename(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var parts = SplitPath(d.FullName);
			if (parts.Length == 0)
				return "";

			return parts[parts.Length - 1];
		}

		// returns the full path of the document
		//
		public static string FullPath(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			return d.FullName;
		}

		// walks the filter up from the document to the project root and joins
		// them with slashes and appends a slash at the end, or returns an empty
		// string if the document is directly in the project root
		//
		public static string FilterPath(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();


			var cookie = Package.Instance.RDT4.GetDocumentCookie(d.FullName);
			if (cookie == VSConstants.VSCOOKIE_NIL)
			{
				Logger.Error(
					"FilterPath: GetDocumentCookie for {0} failed",
					d.FullName);

				return "";
			}


			Package.Instance.RDT4.GetDocumentHierarchyItem(
				cookie, out var h, out var itemid);

			if (h == null || itemid == (uint)VSConstants.VSITEMID.Nil)
			{
				Logger.Error(
					"FilterPath: GetDocumentHierarchyItem for {0} failed",
					d.FullName);

				return "";
			}


			var parts = new List<string>();

			while (itemid != (uint)VSConstants.VSITEMID.Nil)
			{
				var e = h.GetProperty(
					itemid, (int)__VSHPROPID.VSHPROPID_Parent,
					out var parentidObject);

				if (e != VSConstants.S_OK || !(parentidObject is int))
				{
					Logger.Log(
						"FilterPath: GetProperty parent for {0} failed, {1}",
						d.FullName, e);

					break;
				}

				var parentid = (uint)(int)parentidObject;

				// no more parent
				if (parentid == (uint)VSConstants.VSITEMID.Nil)
					break;

				e = h.GetGuidProperty(
					parentid, (int)__VSHPROPID.VSHPROPID_TypeGuid,
					out var type);

				if (e != VSConstants.S_OK || type == null)
				{
					Logger.Error(
						"FilterPath: GetProperty typeguid for {0} failed, {1}",
						d.FullName, e);

					return "";
				}

				// ignore anything but folders
				if (type == VSConstants.ItemTypeGuid.PhysicalFolder_guid ||
					type == VSConstants.ItemTypeGuid.VirtualFolder_guid)
				{
					e = h.GetProperty(
						parentid, (int)__VSHPROPID.VSHPROPID_Name,
						out var nameObject);

					if (e != VSConstants.S_OK || !(nameObject is string))
					{
						Logger.Error(
							"FilterPath: GetProperty name for {0} failed, {1}",
							d.FullName, e);

						return "";
					}

					parts.Insert(0, (string)nameObject);
				}

				itemid = parentid;
			}


			if (DocumentManager.IsInBuiltinProject(d))
			{
				// sigh
				//
				// some of the builtin projects like miscellaneous items seem
				// to behave both as projects and folders
				//
				// the check above for physical/virtual folders works fine for
				// regular projects, where the root project item doesn't say
				// it's a folder (cause it ain't)
				//
				// but when an external file is opened, it gets put in this
				// magic miscellaneous items project, which _does_ report its
				// type as a virtual folder, event though it's the root
				// "project"
				//
				// in any case, if the document is in a builtin project, the
				// first component is removed, because there doesn't seem to be
				// any other way to verify it

				if (parts.Count > 0)
					parts.RemoveAt(0);
			}


			var s = String.Join("/", parts);

			// ending slash
			if (s.Length > 0)
				s += "/";

			return s;
		}

		// returns the name of the document's parent filter, or an empty string
		//
		public static string ParentFilter(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (d?.ProjectItem?.Collection?.Parent is ProjectItem pi)
				return pi.Name;

			return "";
		}


		// splits the given path on slash and backslash
		//
		private static string[] SplitPath(string path)
		{
			var seps = new char[] {
				Path.DirectorySeparatorChar,
				Path.AltDirectorySeparatorChar };

			return path.Split(seps, StringSplitOptions.RemoveEmptyEntries);
		}
	}
}
