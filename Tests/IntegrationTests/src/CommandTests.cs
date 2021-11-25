using Microsoft.VisualStudio.TestTools.UnitTesting;
using EnvDTE;

namespace CustomTabNames.Tests
{
	[TestClass]
	public class CommandTests
	{
		[TestMethod]
		public void Enabled()
		{
			var cmd = Global.Operations.GetCommand(
				Commands.Enabled.Set, Commands.Enabled.ID);

			Assert.IsNotNull(cmd);

			Assert.AreEqual(true, Global.Operations.GetBoolOption("Enabled"));
			//Assert.AreEqual(true, IsChecked(cmd));

			Global.Operations.ExecuteCommand(
				Commands.Enabled.Set, Commands.Enabled.ID);

			Assert.AreEqual(false, Global.Operations.GetBoolOption("Enabled"));
			//Assert.AreEqual(false, IsChecked(cmd));

			Global.Operations.ExecuteCommand(
				Commands.Enabled.Set, Commands.Enabled.ID);

			Assert.AreEqual(true, Global.Operations.GetBoolOption("Enabled"));
			//Assert.AreEqual(true, IsChecked(cmd));
		}

		private bool IsChecked(Command c)
		{
			// there doesn't seem to be any way to get the checked status for
			// a command
			return false;
		}
	}
}
