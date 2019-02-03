using System.Collections.Generic;
using Microsoft.VisualStudio.Shell;
using EnvDTE;

namespace CustomTabNames
{
	public class VSProject : IProject
	{
		// built-in projects that can be ignored
		//
		private static readonly List<string> BuiltinProjects = new List<string>()
		{
			Constants.vsProjectKindMisc,
			Constants.vsProjectKindSolutionItems,
			Constants.vsProjectKindUnmodeled
		};

		private readonly Project p;

		public VSProject(Project p)
		{
			this.p = p;
		}

		public string Name
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();
				return p.Name;
			}
		}

		public bool IsBuiltIn
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				var k = p?.Kind;
				if (k == null)
					return false;

				return BuiltinProjects.Contains(k);
			}
		}
	}
}
