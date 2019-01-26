using EnvDTE;
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.Shell;

namespace CustomTabNames
{
	using VariablesDictionary = Dictionary<string, Func<Document, string>>;

	public sealed class Variables
	{
		public static VariablesDictionary MakeDictionary()
		{
			return new VariablesDictionary()
			{
				{"ProjectName", ProjectName},
				{"ParentDir",   ParentDir},
				{"Filename",    Filename},
				{"FullPath",    FullPath}
			};
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
	}
}
