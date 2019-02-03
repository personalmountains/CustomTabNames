using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace CustomTabNames
{
	using Dict = Dictionary<string, IVariable>;

	public interface IVariable
	{
		string Expand(IDocument d);
	}

	public sealed class Variables
	{
		// maps variable names to functions
		//
		private static readonly Dict Dictionary = new Dict()
		{
			// variable names must be [a-zA-Z]
			{"ProjectName",    new ProjectName()},
			{"ParentDir",      new ParentDir()},
			{"Filename",       new Filename()},
			{"FullPath",       new FullPath()},
			{"FolderPath",     new FolderPath()},
			{"ParentFolder",   new ParentFolder()}
			// update Strings.OptionTemplateDescription when changing this list
		};


		private static Logger Logger
		{
			get
			{
				return Main.Instance.Logger;
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
				string s = v.Expand(d);

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
	}


	public class VariableUtilities
	{
		public static List<string> FolderPathParts(IDocument d)
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

			return parts;
		}

		// splits the given path on slash and backslash
		//
		public static string[] SplitPath(string path)
		{
			var seps = new char[] {
				Path.DirectorySeparatorChar,
				Path.AltDirectorySeparatorChar };

			var cs = path.Split(seps, StringSplitOptions.RemoveEmptyEntries);

			if (cs.Length > 0)
			{
				if (IsDriveLetter(cs[0]))
				{
					var list = new List<string>(cs);
					list.RemoveAt(0);
					cs = list.ToArray();
				}
			}

			return cs;
		}

		public static bool IsDriveLetter(string s)
		{
			if (s.Length != 2 || s[1] != ':')
				return false;

			var d = s[0];
			if ((d >= 'a' && d <= 'z') || (d >= 'A' && d <= 'Z'))
				return true;

			return false;
		}
	}


	// returns the name of the document's project, or an empty string; this
	// can happen for external files, includes, etc.
	//
	public class ProjectName : IVariable
	{
		public string Expand(IDocument d)
		{
			if (Main.Instance.Options.IgnoreSingleProject)
			{
				if (Main.Instance.Solution.HasSingleProject)
					return "";
			}

			if (Main.Instance.Options.IgnoreBuiltinProjects)
			{
				if (d.Project?.IsBuiltIn ?? false)
					return "";
			}

			var p = d?.Project;
			if (p == null)
				return "";

			return p.Name;
		}
	}


	// returns the parent directory of the document's file with a slash at
	// the end, or an empty string; this happens for files in the root
	//
	public class ParentDir : IVariable
	{
		public string Expand(IDocument d)
		{
			var parts = VariableUtilities.SplitPath(d.Path);
			if (parts.Length < 2)
				return "";

			return parts[parts.Length - 2];
		}
	}


	// returns the last component of the full path, or an empty string
	//
	public class Filename : IVariable
	{
		public string Expand(IDocument d)
		{
			var parts = VariableUtilities.SplitPath(d.Path);
			if (parts.Length == 0)
				return "";

			return parts[parts.Length - 1];
		}
	}


	// returns the full path of the document
	//
	public class FullPath : IVariable
	{
		public string Expand(IDocument d)
		{
			return d.Path;
		}
	}


	// walks the folders up from the document to the project root and joins
	// them with slashes and appends a slash at the end, or returns an empty
	// string if the document is directly in the project root
	//
	public class FolderPath : IVariable
	{
		public string Expand(IDocument d)
		{
			return string.Join("/", VariableUtilities.FolderPathParts(d));
		}
	}


	// returns the name of the document's parent folder, or an empty string
	//
	public class ParentFolder : IVariable
	{
		public string Expand(IDocument d)
		{
			var s = "";

			var parts = VariableUtilities.FolderPathParts(d);
			if (parts.Count > 0)
				s = parts[parts.Count - 1];

			return s;
		}
	}
}
