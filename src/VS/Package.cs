using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

// the Package class is the package; it starts the DocumentManager,
// waits for events and calls FixCaption() on documents
//
// the DocumentManager and class registers events like opening documents and
// notifies Package that a document caption needs fixing
//
// Logger has simple static functions to log strings to the output window and
// Strings has most of the localizable strings
//
// Options has all the options and acts as a DialogPage that can be shown in the
// options dialog
//
// Variables has all available variables and can expand them based on a template
// and a document

namespace CustomTabNames
{
	[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
	[InstalledProductRegistration(Strings.ExtensionName, Strings.ExtensionDescription, Strings.ExtensionVersion)]
	[ProvideService(typeof(Package), IsAsyncQueryable = true)]
	[ProvideAutoLoad(UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideOptionPage(typeof(VSOptions), Strings.ExtensionName, Strings.OptionsCategory, 0, 0, true)]
	[ProvideProfile(typeof(VSOptions), Strings.ExtensionName, Strings.OptionsCategory, 0, 0, isToolsOptionPage: true)]
	[Guid(Strings.ExtensionGuid)]
	[ProvideMenuResource("Menus.ctmenu", 1)]
	public sealed class Package : AsyncPackage
	{
		public static Package Instance { get; private set; }
		public Main Main { get; private set; }
		public VSOptions VSOptions { get; private set; }

		// services
		IVsSolution solution = null;
		IVsRunningDocumentTable rdt = null;
		IVsRunningDocumentTable4 rdt4 = null;
		IVsOutputWindow outputWindow = null;


		public Package()
		{
			Instance = this;
		}

		protected override async Task InitializeAsync(
			CancellationToken ct, IProgress<ServiceProgressData> p)
		{
			await JoinableTaskFactory.SwitchToMainThreadAsync(ct);

			if (Solution == null || RDT == null)
			{
				Main.Logger.Error("bailing out");
				return;
			}

			VSOptions = (VSOptions)GetDialogPage(typeof(VSOptions));

			Main = new Main(
				VSOptions, new VSLogger(), new VSSolution(),
				new VSDocumentManager());

			await Commands.Init.InitializeAllAsync(this);
		}

		public IVsSolution Solution
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (solution == null)
				{
					solution = GetService(typeof(SVsSolution)) as IVsSolution;
					if (solution == null)
						Main.Logger.Error("failed to get IVsSolution");
				}

				return solution;
			}
		}

		public IVsRunningDocumentTable RDT
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (rdt == null)
				{
					rdt = GetService(typeof(SVsRunningDocumentTable))
						as IVsRunningDocumentTable;

					if (rdt == null)
						Main.Logger.Error("can't get IVsRunningDocumentTable");
				}

				return rdt;
			}
		}

		public IVsRunningDocumentTable4 RDT4
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (rdt4 == null)
				{
					rdt4 = GetService(typeof(SVsRunningDocumentTable))
						as IVsRunningDocumentTable4;

					if (rdt4 == null)
						Main.Logger.Error("can't get IVsRunningDocumentTable4");
				}

				return rdt4;
			}
		}

		public IVsOutputWindow OutputWindow
		{
			get
			{
				ThreadHelper.ThrowIfNotOnUIThread();

				if (outputWindow == null)
				{
					outputWindow = GetService(typeof(SVsOutputWindow))
						as IVsOutputWindow;

					if (outputWindow == null)
						Main.Logger.Error("can't get IVsOutputWindow");
				}

				return outputWindow;
			}
		}
	}
}
