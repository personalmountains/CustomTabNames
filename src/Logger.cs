using Microsoft.VisualStudio.Threading;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using System;

namespace CustomTabNames
{
	// logs strings to the Output window, in a specific pane
	//
	class Logger
	{
		// output window
		private static IVsOutputWindowPane pane = null;

		// logs the given string by calling String.Format()
		//
		public static void Log(string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// don't do anything if logging is disabled
			if (!CustomTabNames.Instance.Options.Logging)
				return;

			// make sure the pane exists
			if (!CheckPane())
				return;

			pane.OutputString(String.Format(format, args) + "\n");
		}

		// creates the pane in the output window if necessary, returns whether
		// the pane is available
		//
		private static bool CheckPane()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// create only once
			if (pane != null)
				return true;

			// try getting the output window
			var w = GetOutputWindow();
			if (w == null)
				return false;

			// create a new pane for this extension; this adds an entry in the
			// "show output from" combo box
			var guid = new System.Guid(Strings.ExtensionGuid);
			w.CreatePane(
				ref guid, Strings.ExtensionName,
				Convert.ToInt32(true), Convert.ToInt32(false));

			// try to get the pane that was just created
			w.GetPane(guid, out pane);
			if (pane == null)
				return false;

			return true;
		}

		// returns the Output window
		//
		private static IVsOutputWindow GetOutputWindow()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			CustomTabNames.Instance.ServiceProvider.QueryService(
				typeof(SVsOutputWindow), out var w);

			if (w != null)
			{
				if (w is IVsOutputWindow)
					return (IVsOutputWindow)w;
			}

			return null;
		}
	}
}
