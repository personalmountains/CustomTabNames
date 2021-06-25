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
	public static class Strings
	{
		// update .vsixmanifest if this changes
#if CUSTOM_TAB_NAMES_2019
		public const string ExtensionGuid = "bee6c21e-fbf8-49b1-a0f8-89d7dfa732ee";
#elif CUSTOM_TAB_NAMES_2022
		public const string ExtensionGuid = "6fcb1c98-6501-4d92-80d2-6a0f88b03d37";
#endif

		public const string ExtensionName = "CustomTabNames";
		public const string ExtensionDescription = "Customizes editor tab names";
		public const string ExtensionVersion = "1.2";
		public const string AssemblyVersion = ExtensionVersion + ".0.0";
		public const string Company = "personalmountains";
		public const string Copyright = "CC0 1.0 Universal";
		public const string ExtensionID = ExtensionName + "." + ExtensionGuid;


		// option strings
		public const string OptionsCategory = "General";

		public const string OptionEnabled = "Enabled";
		public const string OptionEnabledDescription =
			"Whether the extension is enabled.";

		public const string OptionTemplate = "Template";
		public const string OptionTemplateDescription =
			"Variables: ProjectName, ParentDir, Filename, FullPath, " +
			"FolderPath, ParentFolder.\r\nText between quotes inside the " +
			"variable will only be added if the variable expansion is not " +
			"empty.";

		public const string OptionIgnoreBuiltinProjects =
			"Ignore built-in projects";
		public const string OptionIgnoreBuiltinProjectsDescription =
			"Don't display built-in project names like 'Miscellaneous Files'.";

		public const string OptionIgnoreSingleProject = "Ignore single project";
		public const string OptionIgnoreSingleProjectDescription =
			"Don't expand $(ProjectName) when the solution only has one " +
			"project.";

		public const string OptionLogging = "Logging";
		public const string OptionLoggingDescription =
			"Whether logging to output window is enabled.";

		public const string OptionLoggingLevel = "Logging level";
		public const string OptionLoggingLevelDescription =
			"When logging is enabled, determines the highest level to log.\n" +
			"0=Error, 1=Warn, 2=Log, 3=Trace 4=Variable expansions";
	}
}
