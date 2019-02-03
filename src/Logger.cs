﻿using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using System;
using System.Globalization;

namespace CustomTabNames
{
	public interface ILogger
	{
		void Output(string s);
	}


	// logs strings to the Output window, in a specific pane; levels are not
	// implemented yet
	//
	public class Logger
	{
		public static Logger Instance { get; private set; }
		private readonly ILogger impl;

		public Logger(ILogger impl)
		{
			Instance = this;
			this.impl = impl;
		}

		// logs the given string by calling String.Format()
		//
		public void Error(string format, params object[] args)
		{
			LogImpl(0, format, args);
		}

		// logs the given string and error code by calling String.Format()
		public void ErrorCode(int e, string format, params object[] args)
		{
			var s = SafeFormat(format, args);
			LogImpl(0, "{0}, error 0x{1:X}", new object[] { s, e });
		}

		// logs the given string by calling String.Format()
		//
		public void Warn(string format, params object[] args)
		{
			LogImpl(1, format, args);
		}

		// logs the given string by calling String.Format()
		//
		public void Log(string format, params object[] args)
		{
			LogImpl(2, format, args);
		}

		// logs the given string by calling String.Format()
		//
		public void Trace(string format, params object[] args)
		{
			LogImpl(3, format, args);
		}

		// logs the given string by calling String.Format()
		//
		public void VariableTrace(string format, params object[] args)
		{
			LogImpl(4, format, args);
		}

		// always logs the given string, regardless of level, as long as
		// logging is enabled
		//
		public void LogAlways(string format, params object[] args)
		{
			LogImpl(-1, format, args);
		}


		private void LogImpl(int level, string format, object[] args)
		{
			// don't do anything if logging is disabled
			if (!Package.Instance.Options.Logging)
				return;

			// don't log if level is too high, always log if level is -1,
			// see LogAlways()
			if (level >= 0 && level > Package.Instance.Options.LoggingLevel)
				return;

			impl.Output(SafeFormat(format, args));
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
	}


	public abstract class LoggingContext
	{
		private readonly Logger lg;

		public LoggingContext()
		{
			this.lg = Logger.Instance;
		}

		protected abstract string LogPrefix();

		public void Error(string format, params object[] args)
		{
			lg.Error("{0}", MakeString(format, args));
		}

		public void ErrorCode(int e, string format, params object[] args)
		{
			lg.ErrorCode(e, "{0}", MakeString(format, args));
		}

		public void Warn(string format, params object[] args)
		{
			lg.Warn("{0}", MakeString(format, args));
		}

		public void Log(string format, params object[] args)
		{
			lg.Log("{0}", MakeString(format, args));
		}

		public void Trace(string format, params object[] args)
		{
			lg.Trace("{0}", MakeString(format, args));
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
