using System;
using System.Collections.Generic;

namespace CustomTabNames
{
	public interface IOptionsBackend
	{
		bool Enabled { get; set; }
		string Template { get; set; }
		bool IgnoreBuiltinProjects { get; set; }
		bool IgnoreSingleProject { get; set; }
		bool Logging { get; set; }
		int LoggingLevel { get; set; }
		void RegisterCallback(Action<string> a);
	}

	public interface ISolution
	{
		List<ITreeItem> ProjectItems
		{
			get;
		}

		List<IDocument> Documents
		{
			get;
		}

		bool HasSingleProject
		{
			get;
		}
	}

	public interface IDocument
	{
		string Path
		{
			get;
		}

		string Name
		{
			get;
		}

		IProject Project
		{
			get;
		}

		ITreeItem TreeItem
		{
			get;
		}

		void SetCaption(string s);
		void ResetCaption();
	}

	public interface IProject
	{
		string Name
		{
			get;
		}

		bool IsBuiltIn
		{
			get;
		}
	}

	public interface ITreeItem
	{
		string Name
		{
			get;
		}

		ITreeItem Parent
		{
			get;
		}

		bool IsFolder
		{
			get;
		}

		string DebugName
		{
			get;
		}
	}
}
