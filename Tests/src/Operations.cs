using System;
using EnvDTE;
using Thread = System.Threading.Thread;

namespace CustomTabNames.Tests
{
	public class Operations
	{
		private readonly DTE dte;
		private readonly Window solutionExplorer;
		private readonly UIHierarchy solutionExplorerRoot;

		public Operations(DTE dte)
		{
			this.dte = dte;

			solutionExplorer =
			dte.Windows.Item(Constants.vsWindowKindSolutionExplorer);

			solutionExplorerRoot =
				(UIHierarchy)solutionExplorer.Object;

			ExpandAll(solutionExplorerRoot.UIHierarchyItems);
		}

		public static bool TryUntilTimeout(
			int timeoutMs, Func<bool> a, int sleepForMs = 1000)
		{
			DateTime startTime = DateTime.Now;

			while (true)
			{
				Thread.Sleep(sleepForMs);

				if (a())
					return true;

				var d = DateTime.Now.Subtract(startTime).TotalMilliseconds;
				if (d >= timeoutMs)
					break;
			}

			return false;
		}

		public Project FindProject(string name)
		{
			foreach (Project p in dte.Solution.Projects)
			{
				if (p.Name == name)
					return p;
			}

			throw new Failed("can't find project {0}", name);
		}

		public Window OpenFile(Project p, string name)
		{
			var pi = p.ProjectItems.Item(name);

			var w = pi.Open(Constants.vsViewKindCode);
			w.Visible = true;

			return w;
		}

		public void MoveFile(string from, string to)
		{
			var fromItem = GetItem(solutionExplorerRoot, from);
			var toItem = GetItem(solutionExplorerRoot, to);

			ActivateWindow(solutionExplorer);
			SelectItem(fromItem);
			DoCommand("Edit.Cut");
			SelectItem(toItem);
			DoCommand("Edit.Paste");
		}

		private void ExpandAll(UIHierarchyItems root)
		{
			root.Expanded = true;

			foreach (UIHierarchyItem i in root)
				ExpandAll(i.UIHierarchyItems);
		}

		public UIHierarchyItem GetItem(UIHierarchy root, string name)
		{
			try
			{
				return root.GetItem(name);
			}
			catch (ArgumentException e)
			{
				throw new Failed(
					"failed to get item {0}, {1}", name, e.Message);
			}
		}

		public void ActivateWindow(Window w)
		{
			try
			{
				w.Activate();
				Yield();
			}
			catch (Exception e)
			{
				throw new Failed(
					"can't activate window {0}, {1}", w.Caption, e.Message);
			}
		}

		public void SelectItem(UIHierarchyItem item)
		{
			try
			{
				item.Select(vsUISelectionType.vsUISelectionTypeSelect);
				Yield();
			}
			catch (Exception e)
			{
				throw new Failed(
					"can't select item {0}, {1}",
					item.Name, e.Message);
			}
		}

		public void DoCommand(string s)
		{
			try
			{
				dte.ExecuteCommand(s);
				Yield();
			}
			catch (Exception e)
			{
				throw new Failed("command '{0}' failed, {1}", s, e.Message);
			}
		}

		public void Yield()
		{
			Thread.Sleep(500);
		}
	}
}
