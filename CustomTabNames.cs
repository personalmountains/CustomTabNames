using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using EnvDTE80;
using System.Collections.Generic;
using Task = System.Threading.Tasks.Task;
using OLE = Microsoft.VisualStudio.OLE;

namespace CustomTabNames
{
	using VariablesDictionary = Dictionary<string, Func<Document, string>>;

	[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
	[ProvideService(typeof(CustomTabNames), IsAsyncQueryable = true)]
	[ProvideAutoLoad(UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
	[Guid(CustomTabNames.Guid)]
	public sealed class CustomTabNames : AsyncPackage
	{
		public const string Guid = "BEE6C21E-FBF8-49B1-A0F8-89D7DFA732EE";

		private DTE2 dte;
		private ServiceProvider sp;
		private DocumentEvents docEvents;
		private WindowEvents winEvents;

		private VariablesDictionary variables;
		private int tries = 0;
		private Timer timer = null;

		private readonly string template =
			"$(ProjectName ':')$(ParentDir)$(Filename)";

		public CustomTabNames()
		{
		}

		protected override async Task InitializeAsync(
			CancellationToken ct, IProgress<ServiceProgressData> p)
		{
			await this.JoinableTaskFactory.SwitchToMainThreadAsync(ct);

			this.dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
			this.sp = new ServiceProvider((OLE.Interop.IServiceProvider)dte);
			this.docEvents = dte.Events.DocumentEvents;
			this.winEvents = dte.Events.WindowEvents;
			this.variables = Variables.MakeDictionary();

			SetEvents();
			_ = FixAllDocumentsAsync();
		}

		void SetEvents()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			docEvents.DocumentOpened += (d) => { OnDocumentOpened(d); };
			winEvents.WindowCreated += (w) => { OnWindowCreated(w); };
			winEvents.WindowActivated += (got, lost) => { OnWindowActivated(got, lost); };
		}

		private void OnDocumentOpened(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (!FixCaption(d))
				Defer();
		}

		private void OnWindowCreated(Window w)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (w.Document != null)
			{
				if (!FixCaption(w.Document))
					Defer();
			}
		}

		private void OnWindowActivated(Window w, Window lost)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (w.Document != null)
			{
				if (!FixCaption(w.Document))
					Defer();
			}
		}

		private void FixAllDocuments()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			bool failed = false;

			foreach (Document d in dte.Documents)
			{
				if (!FixCaption(d))
					failed = true;
			}

			if (failed)
			{
				++tries;
				if (tries == 10)
					return;

				Defer();
			}
		}

		private bool FixCaption(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var f = DocumentFrame(d);
			if (f == null)
				return false;

			var caption = MakeCaption(d);
			if (caption != null)
				SetCaption(f, caption);

			return true;
		}

		private void Defer()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (timer == null)
				timer = new Timer(OnTimer, null, 2000, Timeout.Infinite);
			else
				timer.Change(2000, Timeout.Infinite);
		}

		private void OnTimer(object o)
		{
			timer = null;
			_ = FixAllDocumentsAsync();
		}

		private async Task FixAllDocumentsAsync()
		{
			await JoinableTaskFactory.SwitchToMainThreadAsync();

			tries = 0;
			FixAllDocuments();
		}

		private IVsWindowFrame DocumentFrame(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			VsShellUtilities.IsDocumentOpen(
				sp, d.FullName, VSConstants.LOGVIEWID.Primary_guid,
				out IVsUIHierarchy h, out uint id,
				out IVsWindowFrame frame);

			return frame;
		}

		private string MakeCaption(Document d)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var s = template;
			var re = new Regex(@"\$\((.*?)\s*(?:'(.*)')?\)");

			while (true)
			{
				var m = re.Match(s);
				if (!m.Success)
					break;

				var name = m.Groups[1].Value;
				var text = m.Groups[2].Value;

				var v = variables[name];

				// not found
				if (v == null)
					continue;

				string replacement = v(d);

				// don't append the text if the result was empty
				if (replacement != "")
					replacement += text;

				s = Replace(s, m, replacement);
			}

			return s;
		}

		private string Replace(string String, Match m, string Replacement)
		{
			var ns = String.Substring(0, m.Index);
			ns += Replacement;
			ns += String.Substring(m.Index + m.Length);
			return ns;
		}

		private void SetCaption(IVsWindowFrame f, string s)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			f.SetProperty((int)VsFramePropID.EditorCaption, null);
			f.SetProperty((int)VsFramePropID.OwnerCaption, s);
		}
	}
}
