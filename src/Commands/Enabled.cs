using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using Task = System.Threading.Tasks.Task;

namespace CustomTabNames.Commands
{
	public static class Init
	{
		public static async Task InitializeAllAsync(AsyncPackage package)
		{
			await Enabled.InitializeAsync(package);
		}
	}


	public sealed class Enabled
	{
		// from CustomTabNames.vsct
		public const int ID = 0x0100;
		public static readonly Guid Set =
			new Guid("408d4aff-1925-494c-922d-d360f2705ae2");

		// reference to the button so its checked status can be changed
		private readonly MenuCommand cmd;


		private Enabled(OleMenuCommandService cs)
		{
			cmd = new MenuCommand(Execute, new CommandID(Set, ID));
			cs.AddCommand(cmd);

			cmd.Checked = Main.Instance.Options.Enabled;
			Main.Instance.Options.EnabledChanged += () =>
			{
				cmd.Checked = Main.Instance.Options.Enabled;
			};
		}

		public static async Task InitializeAsync(AsyncPackage package)
		{
			await ThreadHelper.JoinableTaskFactory
				.SwitchToMainThreadAsync(package.DisposalToken);

			var cs = await package.GetServiceAsync(typeof(IMenuCommandService));
			new Enabled(cs as OleMenuCommandService);
		}

		private void Execute(object sender, EventArgs e)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			Main.Instance.Options.Enabled = !Main.Instance.Options.Enabled;
		}
	}
}
