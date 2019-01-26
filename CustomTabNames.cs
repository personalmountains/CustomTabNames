using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using EnvDTE80;
using Task = System.Threading.Tasks.Task;

namespace CustomTabNames
{
	[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
	[ProvideService(typeof(CustomTabNames), IsAsyncQueryable = true)]
	[ProvideAutoLoad(UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
	[Guid(CustomTabNames.Guid)]
	public sealed class CustomTabNames : AsyncPackage
	{
		public const string Guid = "BEE6C21E-FBF8-49B1-A0F8-89D7DFA732EE";
		private readonly DTE2 dte;

		public CustomTabNames()
		{
			this.dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
		}

		protected override async Task InitializeAsync(
			CancellationToken ct, IProgress<ServiceProgressData> p)
		{
			await this.JoinableTaskFactory.SwitchToMainThreadAsync(ct);
		}
	}
}
