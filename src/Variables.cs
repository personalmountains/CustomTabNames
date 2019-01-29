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
		public static VariablesDictionary Dictionary = new VariablesDictionary()
		{
			{"ProjectName",    ProjectName},
			{"ParentDir",      ParentDir},
			{"Filename",       Filename},
			{"FullPath",       FullPath},
			{"FilterPath",     FilterPath},
			{"ParentFilter",   ParentFilter}

			// update Options.Template description when changing this list
		};


		public static string MakeCaption(Document d, string template)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			string s = template;
			var re = new Regex(@"\$\((.*?)\s*(?:'(.*?)')?\)");

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
				var replacement = Expand(d, name, text);

				s = ReplaceMatch(s, m, replacement);
			}

			Logger.Log("  . caption is now {0}", s);

			return s;
		}

		private static string Expand(Document d, string name, string text)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (Dictionary.TryGetValue(name, out var v))
			{
				string s = v(d);

				// don't append the text if the result was empty
				if (s != "")
					s += text;

				Logger.Log("  . variable {0} replaced by '{1}'", name, s);

				return s;
			}
			else
			{
				// not found, put the variable name to notify the user
				Logger.Log("  . variable {0} not found", name);
				return name;
			}
		}

		private static string ReplaceMatch(string s, Match m, string repl)
		{
			var ns = s.Substring(0, m.Index);
			ns += repl;
			ns += s.Substring(m.Index + m.Length);
			return ns;
		}


		public static string ProjectName(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var pname = d?.ProjectItem?.ContainingProject?.Name;
			if (pname == null)
				return "";

			return pname;
		}

		public static string ParentDir(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var path = d.FullName;

			var separators = new char[] {
				Path.DirectorySeparatorChar,
				Path.AltDirectorySeparatorChar };

			var parts = path.Split(
				separators, StringSplitOptions.RemoveEmptyEntries);

			if (parts.Length < 2)
				return "";

			return parts[parts.Length - 2] + "/";
		}

		public static string Filename(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var path = d.FullName;

			var separators = new char[] {
				Path.DirectorySeparatorChar,
				Path.AltDirectorySeparatorChar };

			var parts = path.Split(
				separators, StringSplitOptions.RemoveEmptyEntries);

			if (parts.Length == 0)
				return "";

			return parts[parts.Length - 1];
		}

		public static string FullPath(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			return d.FullName;
		}

		public static string FilterPath(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var pi = d?.ProjectItem;
			var s = "";

			while (pi != null)
			{
				var p = pi.Collection?.Parent;
				if (p is ProjectItem)
					pi = (ProjectItem)p;
				else
					break;

				if (s != "")
					s = "/" + s;

				s = pi.Name + s;
			}

			return s;
		}

		public static string ParentFilter(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var p = d?.ProjectItem?.Collection?.Parent;

			if (p is ProjectItem)
				return ((ProjectItem)p).Name;

			return "";
		}
	}
}
