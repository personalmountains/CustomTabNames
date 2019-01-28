using Microsoft.VisualStudio.Threading;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using System;

namespace CustomTabNames
{
	class Logger
	{
		private static IVsOutputWindowPane pane = null;

		public static void Log(string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (!CustomTabNames.Instance.Options.Logging)
				return;

			if (!CheckPane())
				return;

			Logger.pane.OutputString(String.Format(format, args) + "\n");
		}

		private static bool CheckPane()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (pane != null)
				return true;

			var sp = CustomTabNames.Instance.ServiceProvider;

			sp.QueryService(typeof(SVsOutputWindow), out var wo);
			if (wo == null)
				return false;

			var w = (IVsOutputWindow)wo;

			var guid = new System.Guid(CustomTabNames.Guid);
			w.CreatePane(
				ref guid, "CustomTabNames",
				Convert.ToInt32(true), Convert.ToInt32(false));

			w.GetPane(guid, out pane);
			if (pane == null)
				return false;

			return true;
		}
	}
}
