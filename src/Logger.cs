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

		// logs the given string and error code by calling String.Format()
		public static void ErrorCode(int e, string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var s = SafeFormat(format, args);
			LogImpl(0, "{0}, error 0x{1:X}", new object[] { s, e });
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

		// logs the given string by calling String.Format()
		//
		public static void VariableTrace(string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			LogImpl(4, format, args);
		}


		private static void LogImpl(int level, string format, object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// don't do anything if logging is disabled
			if (!Options.Logging)
				return;

			// don't log if level is too high
			if (level > Options.LoggingLevel)
				return;

			// make sure the pane exists
			if (!CheckPane())
				return;

			pane.OutputString(SafeFormat(format, args) + "\n");
		}

		public static string SafeFormat(string format, object[] args)
		{
			try
			{
				return string.Format(
					CultureInfo.InvariantCulture, format, args);
			}
			catch (System.FormatException e)
			{
				return
					"failed to log string '" + format + "' " +
					"with " + args.Length + " arguments, " +
					e.Message;
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


	public abstract class LoggingContext
	{
		public LoggingContext()
		{
		}

		protected abstract string LogPrefix();

		public void Error(string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Error("{0}", MakeString(format, args));
		}

		public void ErrorCode(int e, string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.ErrorCode(e, "{0}", MakeString(format, args));
		}

		public void Warn(string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Warn("{0}", MakeString(format, args));
		}

		public void Log(string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Log("{0}", MakeString(format, args));
		}

		public void Trace(string format, params object[] args)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Logger.Trace("{0}", MakeString(format, args));
		}

		private string MakeString(string format, object[] args)
		{
			var s = LogPrefix();
			if (s.Length > 0)
				s += ": ";

			s += Logger.SafeFormat(format, args);

			return s;
		}
	}
}
