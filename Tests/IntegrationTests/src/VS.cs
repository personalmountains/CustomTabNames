using System;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using EnvDTE;
using Process = System.Diagnostics.Process;

// https://www.helixoft.com/blog/creating-envdte-dte-for-vs-2017-from-outside-of-the-devenv-exe.html

namespace CustomTabNames.Tests
{
	public class Failed : Exception
	{
		public Failed(string message, params object[] args)
			: base(String.Format(message, args))
		{
		}
	}

	public class VS : IDisposable
	{
		const bool ShowWindow = true;
		const bool CloseAfter = false;

#if CUSTOM_TAB_NAMES_2019
		const string ProcessPath =
			@"C:\Program Files (x86)\Microsoft Visual Studio\2019\Preview\" +
			@"Common7\IDE\devenv.exe";
#elif CUSTOM_TAB_NAMES_2022
		const string ProcessPath =
			@"C:\Program Files\Microsoft Visual Studio\2022\Preview\" +
			@"Common7\IDE\devenv.exe";
#endif

		const string ProcessArguments = "/Embedding /rootsuffix Exp";

		private DTE dte = null;
		public Operations Operations { get; private set; }

		[DllImport("ole32.dll")]
		private static extern int CreateBindCtx(uint r, out IBindCtx ppbc);

		public VS(string solutionPath)
		{
			StartVS(solutionPath, 30 * 1000);
			Operations = new Operations(dte);
			WaitForInit(30 * 1000);
		}

		public void Dispose()
		{
#pragma warning disable 0162
			if (CloseAfter)
				StopVS();
#pragma warning restore 0162
		}

		private void StartVS(string solutionPath, int timeoutMs)
		{
			var proc = AttachVSProcess();

			Operations.TryUntilTimeout(timeoutMs, () =>
			{
				dte = GetDTE(proc.Id);
				return (dte != null);
			});

			if (dte == null)
				throw new Failed("timed out while waiting for process");

			dte.MainWindow.Visible = ShowWindow;
			dte.Solution.Close();

			try
			{
				dte.Solution.Open(solutionPath);
			}
			catch (Exception e)
			{
				throw new Failed(
					"can't open solution from '{0}', {1}",
					solutionPath, e.Message);
			}
		}

		private void StopVS()
		{
			if (dte != null)
			{
				dte.Solution.Close();
				dte.Quit();
			}
		}

		private Process AttachVSProcess()
		{
			var p = FindRunningVSProcess();
			if (p != null)
				return p;

			p = CreateVSProcess();
			if (p != null)
				return p;

			throw new Failed("failed to start VS process");
		}

		private Process FindRunningVSProcess()
		{
			var ps = Process.GetProcesses();

			foreach (var p in ps)
			{
				if (IsVSProcess(p))
					return p;
			}

			return null;
		}

		private Process CreateVSProcess()
		{
			try
			{
				ProcessStartInfo procStartInfo = new ProcessStartInfo
				{
					Arguments = ProcessArguments,
					CreateNoWindow = true,
					FileName = ProcessPath,
					WindowStyle = (ShowWindow ?
						ProcessWindowStyle.Hidden : ProcessWindowStyle.Normal),
					WorkingDirectory =
						System.IO.Path.GetDirectoryName(ProcessPath)
				};

				return Process.Start(procStartInfo);
			}
			catch (Exception w)
			{
				throw new Failed("failed to start VS process: " + w.Message);
			}
		}

		private bool IsVSProcess(Process p)
		{
			try
			{
				var path = p?.MainModule?.FileName ?? "";
				if (String.Compare(path, ProcessPath, true) != 0)
					return false;

				string query =
					"SELECT CommandLine FROM Win32_Process " +
					"WHERE ProcessId = " + p.Id;

				using (var searcher = new ManagementObjectSearcher(query))
				{
					var matchEnum = searcher.Get().GetEnumerator();

					if (matchEnum.MoveNext())
					{
						var cl = matchEnum.Current["CommandLine"]?.ToString();

						if (!cl.Contains(ProcessArguments))
							return false;
					}
				}

				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}

		private DTE GetDTE(int processId)
		{
			object runningObject = null;

			IBindCtx bindCtx = null;
			IRunningObjectTable rot = null;
			IEnumMoniker enumMonikers = null;

			try
			{
				Marshal.ThrowExceptionForHR(CreateBindCtx(0, out bindCtx));
				bindCtx.GetRunningObjectTable(out rot);
				rot.EnumRunning(out enumMonikers);

				IMoniker[] moniker = new IMoniker[1];
				IntPtr numberFetched = IntPtr.Zero;
				while (enumMonikers.Next(1, moniker, numberFetched) == 0)
				{
					IMoniker runningObjectMoniker = moniker[0];

					string name = null;

					try
					{
						if (runningObjectMoniker != null)
						{
							runningObjectMoniker.GetDisplayName(
								bindCtx, null, out name);
						}
					}
					catch (UnauthorizedAccessException)
					{
						// Do nothing, there is something in the ROT that we
						// do not have access to.
					}

					Regex monikerRegex = new Regex(
						@"!VisualStudio.DTE\.\d+\.\d+\:" + processId,
						RegexOptions.IgnoreCase);

					if (!string.IsNullOrEmpty(name))
					{
						if (monikerRegex.IsMatch(name))
						{
							Marshal.ThrowExceptionForHR(rot.GetObject(
								runningObjectMoniker, out runningObject));

							break;
						}
					}
				}
			}
			finally
			{
				if (enumMonikers != null)
				{
					Marshal.ReleaseComObject(enumMonikers);
				}

				if (rot != null)
				{
					Marshal.ReleaseComObject(rot);
				}

				if (bindCtx != null)
				{
					Marshal.ReleaseComObject(bindCtx);
				}
			}

			return runningObject as EnvDTE.DTE;
		}

		private void WaitForInit(int timeoutMs)
		{
			bool inited = false;

			Operations.TryUntilTimeout(timeoutMs, () =>
			{
			try
			{
					Operations.SetOption("Enabled", true);

					// toggle it to make sure it logs something
					Operations.SetOption("Logging", false);
					Operations.SetOption("Logging", true);

					var s = LoggingPaneText();

					if (s.Contains("logging enabled") ||
						s.Contains("initialized"))
					{
						inited = true;
						return true;
					}
				}
				catch (Exception)
				{
					// eat, try again later
				}

				return false;
			});

			if (!inited)
				throw new Failed("extension didn't initialize");
		}

		private string LoggingPaneText()
		{
			var o = dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
			if (o == null)
				return "";

			if (o.Object is OutputWindow ow)
			{
				var p = ow.OutputWindowPanes.Item("CustomTabNames");
				if (p == null)
					return "";

				var sel = p.TextDocument?.Selection;
				if (sel == null)
					return "";

				sel.StartOfDocument();
				sel.EndOfDocument(true);
				var t = sel.Text;

				// reset selection
				sel.EndOfDocument(false);

				return t;
			}

			return "";
		}
	}
}
