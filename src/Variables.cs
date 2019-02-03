using EnvDTE;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;

namespace CustomTabNames
{
	using Dict = Dictionary<string, Func<IDocument, string>>;

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
			{"FolderPath",     FolderPath},
			{"ParentFolder",   ParentFolder}
			// update Strings.OptionTemplateDescription when changing this list
		};


		private static Options Options
		{
			get
			{
				return Package.Instance.Options;
			}
		}

		private static Logger Logger
		{
			get
			{
				return Logger.Instance;
			}
		}


		// expands all variables on the given template
		//
		public static string Expand(IDocument d, string template)
		{
			string s = template;

			// matches $(VariableName 'text') where 'text' is optional
			// [1] is "VariableName"
			// [2] is "text"
			var re = new Regex(@"\$\(\s*([a-zA-Z]+)\s*(?:'(.*?)')?\)");

			Logger.VariableTrace(
				"making caption for {0} using template {1}",
				d.Path, s);

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

			Logger.VariableTrace("  . caption is now {0}", s);

			return s;
		}

		// returns the expansion of variable 'name' with option text
		//
		private static string ExpandOne(IDocument d, string name, string text)
		{
			if (Dictionary.TryGetValue(name, out var v))
			{
				// variable is in the dictionary

				// expand it
				string s = v(d);

				// don't append the text if the result was empty
				if (s.Length != 0)
					s += text;

				Logger.VariableTrace(
					"  . variable {0} replaced by '{1}'", name, s);

				return s;
			}
			else
			{
				// not found in the dictionary, put the variable name verbatim
				// to notify the user
				Logger.VariableTrace("  . variable {0} not found", name);
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
		public static string ProjectName(IDocument d)
		{
			if (Options.IgnoreSingleProject)
			{
				if (Package.Instance.Solution.HasSingleProject)
					return "";
			}

			if (Options.IgnoreBuiltinProjects)
			{
				if (d.Project?.IsBuiltIn ?? false)
					return "";
			}

			var p = d?.Project;
			if (p == null)
				return "";

			return p.Name;
		}

		// returns the parent directory of the document's file with a slash at
		// the end, or an empty string; this happens for files in the root
		//
		public static string ParentDir(IDocument d)
		{
			var parts = Utilities.SplitPath(d.Path);
			if (parts.Length < 2)
				return "";

			return parts[parts.Length - 2] + "/";
		}

		// returns the last component of the full path, or an empty string
		//
		public static string Filename(IDocument d)
		{
			var parts = Utilities.SplitPath(d.Path);
			if (parts.Length == 0)
				return "";

			return parts[parts.Length - 1];
		}

		// returns the full path of the document
		//
		public static string FullPath(IDocument d)
		{
			return d.Path;
		}

		// walks the folders up from the document to the project root and joins
		// them with slashes and appends a slash at the end, or returns an empty
		// string if the document is directly in the project root
		//
		public static string FolderPath(IDocument d)
		{
			return string.Join("/", FolderPathParts(d));
		}

		// returns the name of the document's parent folder, or an empty string
		//
		public static string ParentFolder(IDocument d)
		{
			var s = "";

			var parts = FolderPathParts(d);
			if (parts.Count > 0)
				s = parts[parts.Count - 1];

			return s;
		}


		private static List<string> FolderPathParts(IDocument d)
		{
			var parts = new List<string>();

			var item = d.TreeItem;

			while (item != null)
			{
				var parent = item.Parent;
				if (parent == null)
					break;

				if (parent.IsFolder)
				{
					var name = parent.Name;
					if (name != null)
						parts.Insert(0, name);
				}

				item = parent;
			}


			if (d?.Project?.IsBuiltIn ?? false)
			{
				// sigh
				//
				// some of the built-in projects like miscellaneous items seem
				// to behave both as projects and folders
				//
				// the ItemIsFolder() call above checks for physical/virtual
				// folders and works fine for regular projects, where the root
				// project item doesn't say it's a folder (cause it ain't)
				//
				// but when an external file is opened, it gets put in this
				// magic miscellaneous items project, which _does_ report its
				// type as a virtual folder, event though it's the root
				// "project"
				//
				// in any case, if the document is in a built-in project, the
				// first component is removed, because there doesn't seem to be
				// any way to figure out whether that node is a project or an
				// actual folder

				if (parts.Count > 0)
					parts.RemoveAt(0);
			}

			return parts;
		}
	}
}
