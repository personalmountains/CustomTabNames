using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace CustomTabNames
{
	public class VSLogger : ILogger
	{
		// output pane
		private IVsOutputWindowPane pane = null;

		public void Output(string s)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			// make sure the pane exists
			if (!CheckPane())
				return;

			pane.OutputString(s + "\n");
		}

		// creates the pane in the output window if necessary, returns whether
		// the pane is available
		//
		private bool CheckPane()
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
