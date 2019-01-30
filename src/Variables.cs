using EnvDTE;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.Shell;

namespace CustomTabNames
{
	using VariablesDictionary = Dictionary<string, Func<Document, string>>;

	public sealed class Variables
	{
		// maps variable names to functions
		//
		public static VariablesDictionary Dictionary = new VariablesDictionary()
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
				return CustomTabNames.Instance.Options;
			}
		}

		private static DocumentManager DocumentManager
		{
			get
			{
				return CustomTabNames.Instance.DocumentManager;
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

			Logger.Log(
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

			Logger.Log("  . caption is now {0}", s);

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
				if (s != "")
					s += text;

				Logger.Log("  . variable {0} replaced by '{1}'", name, s);

				return s;
			}
			else
			{
				// not found in the dictionary, put the variable name verbatim
				// to notify the user
				Logger.Log("  . variable {0} not found", name);
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

			var s = "";

			// getting the project item associated with the document, may be
			// null if the document is not in a project
			var item = d?.ProjectItem;

			while (item != null)
			{
				// getting the item's parent, which can either be a filter or
				// the project root (or something else?)
				item = item.Collection?.Parent as ProjectItem;

				if (item == null)
				{
					// no more filters
					break;
				}

				s = item.Name + "/" + s;
			}

			return s;
		}

		// returns the name of the document's parent filter, or an empty string
		//
		public static string ParentFilter(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var p = d?.ProjectItem?.Collection?.Parent as ProjectItem;
			if (p == null)
				return "";

			return p.Name;
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
