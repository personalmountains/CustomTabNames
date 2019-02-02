using System;
using System.Diagnostics;
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
		const string devenv =
			@"C:\Program Files (x86)\Microsoft Visual Studio\2019\Preview\" +
			@"Common7\IDE\devenv.exe";

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
			StopVS();
		}

		private void StartVS(string solutionPath, int timeoutMs)
		{
			var proc = StartVSProcess();

			Operations.TryUntilTimeout(timeoutMs, () =>
			{
				dte = GetDTE(proc.Id);
				return (dte != null);
			});

			if (dte == null)
				throw new Failed("timed out while waiting for process");

			dte.MainWindow.Visible = true;
			dte.Solution.Open(solutionPath);
		}

		private void StopVS()
		{
			if (dte != null)
			{
				dte.Solution.Close();
				dte.Quit();
			}
		}

		private Process StartVSProcess()
		{
			try
			{
				ProcessStartInfo procStartInfo = new ProcessStartInfo
				{

					//Arguments = "/Embedding /rootsuffix Exp";
					Arguments = "/rootsuffix Exp",
					CreateNoWindow = true,
					FileName = devenv,
					WindowStyle = ProcessWindowStyle.Hidden,
					WorkingDirectory =
					System.IO.Path.GetDirectoryName(devenv)
				};

				return Process.Start(procStartInfo);
			}
			catch (Exception w)
			{
				throw new Failed("failed to start VS process: " + w.Message);
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
					if (!Operations.SetExtensionOption("Enabled", true))
						throw new Failed("can't set Enabled option");

					if (!Operations.SetExtensionOption("Logging", true))
						throw new Failed("can't set Logging option");

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

				return sel.Text;
			}

			return "";
		}
	}
}
