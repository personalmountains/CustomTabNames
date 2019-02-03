namespace CustomTabNames
{
	public class Main
	{
		public static Main Instance { get; private set; }
		public Options Options { get; private set; }
		public Logger Logger { get; private set; }
		public ISolution Solution { get; private set; }
		public DocumentManager DocumentManager { get; private set; }

		// whether Start() has already been called
		private bool started = false;

		public Main(
			IOptionsBackend o,
			ILogger logger, ISolution solution, DocumentManager dm)
		{
			Instance = this;

			Options = new Options(o);
			Logger = new Logger(logger);
			Solution = solution;
			DocumentManager = dm;

			DocumentManager.DocumentChanged += OnDocumentChanged;
			DocumentManager.ContainersChanged += OnContainersChanged;

			Options.EnabledChanged += OnEnabledChanged;
			Options.TemplateChanged += OnTemplateChanged;
			Options.IgnoreBuiltinProjectsChanged += OnIgnoreBuiltinProjectsChanged;
			Options.IgnoreSingleProjectChanged += OnIgnoreSingleProjectChanged;
			Options.LoggingChanged += OnLoggingChanged;

			if (Options.Enabled)
			{
				Logger.Log("initialized");
				Start();
			}
			else
			{
				Logger.Log("initialized but disabled in the options");
			}
		}

		// starts the manager (which register handlers for documents) and
		// do a first pass on all opened documents
		//
		public void Start()
		{
			Logger.Trace("starting");

			if (started)
			{
				// shouldn't happen
				Logger.Error("already started");
				return;
			}

			started = true;
			DocumentManager.Start();
			FixAllDocuments();
		}

		// stops the manager (kills events for documents) and resets all the
		// captions to their original value
		//
		public void Stop()
		{
			Logger.Trace("stopping");

			if (!started)
			{
				Logger.Error("already stopped");
				return;
			}

			started = false;
			DocumentManager.Stop();
			ResetAllDocuments();
		}

		// fired when the template option changed, fixes all currently opened
		// documents
		//
		private void OnTemplateChanged()
		{
			if (!Options.Enabled)
				return;

			Logger.Log("template option changed");
			FixAllDocuments();
		}

		// fired when the enabled option changed, either starts or stops
		//
		private void OnEnabledChanged()
		{
			Logger.Log("enabled option changed");

			if (Options.Enabled)
				Start();
			else
				Stop();
		}

		// fired when the ignore built-in projects option changed, fixes all
		// currently opened documents
		//
		private void OnIgnoreBuiltinProjectsChanged()
		{
			if (!Options.Enabled)
				return;

			Logger.Log("ignore built-in projects option changed");
			FixAllDocuments();
		}

		// fired when the ignore single project option changed, fixes all
		// currently opened documents
		//
		private void OnIgnoreSingleProjectChanged()
		{
			if (!Options.Enabled)
				return;

			Logger.Log("ignore single project option changed");
			FixAllDocuments();
		}

		// fired when the logging flag is changed
		//
		private void OnLoggingChanged()
		{
			if (Options.Logging)
				Logger.LogAlways("logging enabled");
		}

		// fired when a document or window has been opened
		//
		private void OnDocumentChanged(IDocument d)
		{
			FixCaption(d);
		}

		private void OnContainersChanged()
		{
			Logger.Log("containers changed");
			FixAllDocuments();
		}

		// walks through all opened documents and tries to set the caption for
		// each of them
		//
		public void FixAllDocuments()
		{
			Logger.Log("fixing all documents");

			foreach (var d in Solution.Documents)
				FixCaption(d);
		}

		// walks through all opened documents and resets the caption for each
		// of them; failure is ignored
		//
		private void ResetAllDocuments()
		{
			Logger.Log("reseting all documents");

			foreach (var d in Solution.Documents)
				d.ResetCaption();
		}

		// called on each document by FixAllDocuments() and on documents given
		// by the DocumentChanged event from the DocumentManager
		//
		// creates a caption and tries to set it on the document
		//
		// may fail if the document doesn't have a frame yet
		//
		private void FixCaption(IDocument d)
		{
			d.SetCaption(Variables.Expand(d, Options.Template));
		}
	}
}
