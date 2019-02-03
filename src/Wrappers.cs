using System.Collections.Generic;

namespace CustomTabNames
{
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
