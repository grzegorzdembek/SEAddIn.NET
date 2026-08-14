using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class FitAndUpdateViewsProcessor
    {
        private readonly SeAssembly _assembly;
        private readonly SeApp _application;

        public FitAndUpdateViewsProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _application = _assembly.Application;
        }

        public bool Initialize()
        {
            return true;
        }

        public void Process()
        {
            SeOccurrences occurrences = null;
            HashSet<string> filePaths = new (StringComparer.OrdinalIgnoreCase);

            try
            {
                occurrences = _assembly.Occurrences;
                DataUtils.BuildDataForCorrectModelsView(occurrences, filePaths);

                foreach (string filePath in filePaths)
                {
                    bool isOpen = false;

                    SeDocument document = null;
                    SeWindow window = null;
                    try
                    {
                        document = Helpers.GetOpenDocument(_application, filePath);
                        isOpen = true;
        
                        window = (SeWindow)_application.ActiveWindow;
                        if (window != null)
                        {
                            SeView view = window.View;
                            if (view != null)
                            {
                                try
                                {
                                    view.Update();
                                    view.ApplyNamedView("iso");
                                    view.Fit();
                                    document.Save();
                                }
                                finally
                                {
                                    Helpers.ReleaseCom(ref view);
                                }
                            }
                        }
                    }
                    finally
                    {
                        Helpers.ReleaseCom(ref window);

                        try
                        {
                            if (isOpen)
                            {
                                document?.Close(false);
                            }
                        }
                        catch
                        {
                        }

                        Helpers.ReleaseCom(ref document);

                        System.Windows.Forms.Application.DoEvents();
                        _application.DoIdle();
                    }
                }
            }
            finally
            {
                Helpers.ReleaseCom(ref occurrences);
            }
        }
    }
}