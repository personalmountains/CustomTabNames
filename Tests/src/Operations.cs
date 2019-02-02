using System;
using System.Collections.Generic;
using EnvDTE;
using Thread = System.Threading.Thread;

namespace CustomTabNames.Tests
{
	public class ScopedAction : IDisposable
	{
		private Action action;

		public ScopedAction(Action a)
		{
			action = a;
		}

		public void Dispose()
		{
			action?.Invoke();
		}
	}


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

			ExpandAll();
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

		public bool SetOption(string name, object value)
		{
			try
			{
				var options = dte.Properties["CustomTabNames", "General"];
				if (options == null)
					return false;

				var o = options.Item(name);
				if (o == null)
					return false;

				o.Value = value;
				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}

		public ScopedAction SetOptionTemp(string name, object value)
		{
			try
			{
				var options = dte.Properties["CustomTabNames", "General"];
				if (options == null)
					return null;

				var o = options.Item(name);
				if (o == null)
					return null;

				var old = o.Value;
				o.Value = value;

				return new ScopedAction(() =>
				{
					SetOption(name, old);
				});
			}
			catch (Exception)
			{
				return null;
			}
		}

		public Project GetProject(string name)
		{
			var p = FindProject(name);
			if (p == null)
				throw new Failed("can't find project {0}", name);

			return p;
		}

		public Project FindProject(string name)
		{
			foreach (Project p in dte.Solution.Projects)
			{
				if (p.Name == name)
					return p;
			}

			return null;
		}

		public void AddProject(string path)
		{
			dte.Solution.AddFromFile(path);
			Wait();
		}

		public void RemoveProject(Project p)
		{
			var name = p.Name;
			dte.Solution.Remove(p);

			// try to wait until it's unloaded
			Operations.TryUntilTimeout(5000, () =>
			{
				return (FindProject(name) == null);
			});

			// then wait for an update
			Wait();
		}

		public ScopedAction RemoveProjectTemp(Project p)
		{
			string path = p.FullName;
			RemoveProject(p);

			return new ScopedAction(() =>
			{
				AddProject(path);
			});
		}

		public void RenameProject(Project p, string name)
		{
			p.Name = name;
			Wait();
		}

		public ScopedAction RenameProjectTemp(Project p, string name)
		{
			string old = p.Name;
			RenameProject(p, name);

			return new ScopedAction(() =>
			{
				RenameProject(p, old);
			});
		}

		public void RenameFile(string f, string name)
		{
			var item = GetItem(solutionExplorerRoot, f);
			var pi = item.Object as ProjectItem;
			pi.Name = name;

			// some folders become closed when they're renamed
			ExpandAll();
		}

		public ScopedAction RenameFileTemp(string f, string name)
		{
			var item = GetItem(solutionExplorerRoot, f);
			var pi = item.Object as ProjectItem;

			string old = pi.Name;
			pi.Name = name;

			// some folders become closed when they're renamed
			ExpandAll();

			return new ScopedAction(() =>
			{
				pi.Name = old;

				// some folders become closed when they're renamed
				ExpandAll();
			});
		}

		public void RenameFolder(string f, string name)
		{
			RenameFile(f, name);
		}

		public ScopedAction RenameFolderTemp(string f, string name)
		{
			return RenameFileTemp(f, name);
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

			// folders become closed when they're moved around
			ExpandAll();
		}

		public ScopedAction MoveFileTemp(string from, string to)
		{
			MoveFile(from, to);

			return new ScopedAction(() =>
			{
				var fromParts = new List<string>(from.Split('\\'));
				var filename = fromParts[fromParts.Count - 1];
				fromParts.RemoveAt(fromParts.Count - 1);

				var rfrom = to + "\\" + filename;
				var rto = string.Join("\\", fromParts);

				MoveFile(rfrom, rto);
			});
		}

		public void MoveFolder(string from, string to)
		{
			MoveFile(from, to);
		}

		public ScopedAction MoveFolderTemp(string from, string to)
		{
			return MoveFileTemp(from, to);
		}

		private void ExpandAll(UIHierarchyItems root = null)
		{
			if (root == null)
				root = solutionExplorerRoot.UIHierarchyItems;

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
				Wait();
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
				Wait();
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
				Wait();
			}
			catch (Exception e)
			{
				throw new Failed("command '{0}' failed, {1}", s, e.Message);
			}
		}

		public void Wait(int ms=500)
		{
			Thread.Sleep(ms);
		}
	}
}
