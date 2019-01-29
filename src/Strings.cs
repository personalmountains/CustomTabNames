namespace CustomTabNames
{
	// some strings like ExtensionName and ExtensionDescription are used in
	// various places (CustomTabNames.cs, AssemblyInfo.cs, etc.); others are
	// specific to Options.cs
	//
	// some places cannot use these strings because they're not C# code, like
	// in the .vsixmanifest file
	//
	// anyways, they're all here, might help with localization one day
	//
	public class Strings
	{
		// update .vsixmanifest if this changes
		public const string ExtensionGuid = "bee6c21e-fbf8-49b1-a0f8-89d7dfa732ee";
		public const string ExtensionName = "CustomTabNames";
		public const string ExtensionDescription = "Customizes editor tab names";
		public const string ExtensionVersion = "1.0";
		public const string AssemblyVersion = ExtensionVersion + ".0.0";
		public const string Company = "personalmountains";
		public const string Copyright = "CC0 1.0 Universal";
		public const string ExtensionID = ExtensionName + "." + ExtensionGuid;


		// option strings
		public const string OptionsCategory = "General";

		public const string OptionEnabled = "Enabled";
		public const string OptionEnabledDescription =
			"Whether the extension is enabled";

		public const string OptionTemplate = "Template";
		public const string OptionTemplateDescription =
			"Variables: ProjectName, ParentDir, Filename, FullPath, " +
			"FilterPath, ParentFilter.\r\nText between quotes inside the " +
			"variable will only be added if the variable expansion is not " +
			"empty.";

		public const string OptionLogging = "Logging";
		public const string OptionLoggingDescription =
			"Whether logging to output window is enabled";
	}
}
