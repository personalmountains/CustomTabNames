using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using System;
using System.Globalization;

namespace CustomTabNames
{
	// logs strings to the Output window, in a specific pane; levels are not
	// implemented yet
	//
	class Logger
	{
		// output pane
		static IVsOutputWindowPane pane = null;

		private static Options Options
		{
			get
			{
				return Package.Instance.Options;
			}
		}

		// logs the given string by calling String.Format()
		//
		public static void Error(string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			LogImpl(0, format, args);
		}

		// logs the given string by calling String.Format()
		//
		public static void Warn(string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			LogImpl(1, format, args);
		}

		// logs the given string by calling String.Format()
		//
		public static void Log(string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			LogImpl(2, format, args);
		}

		// logs the given string by calling String.Format()
		//
		public static void Trace(string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			LogImpl(3, format, args);
		}

		private static void LogImpl(
			int level, string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			try
			{
				// don't do anything if logging is disabled
				if (!Options.Logging)
					return;

				// don't log if level is too high
				if (level > Options.LoggingLevel)
					return;

				// make sure the pane exists
				if (!CheckPane())
					return;

				pane.OutputString(String.Format(
					CultureInfo.InvariantCulture, format, args) + "\n");
			}
			catch (System.FormatException e)
			{
				pane.OutputString(
					"failed to log string '" + format + "' " +
					"with " + args.Length + " arguments, " +
					e.Message + "\n");
			}
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
			var w = Package.Instance.OutputWindow;
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
	}
}
