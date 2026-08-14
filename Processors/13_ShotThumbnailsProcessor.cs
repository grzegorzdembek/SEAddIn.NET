using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class ShotThumbnailsProcessor
    {
        private readonly SeAssembly _assembly;
        private readonly SeApp _application;

        private string _assemblyPath;
        private string _projectDirectory;
        private string _thumbnailsDirectory;

        public ShotThumbnailsProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _application = _assembly.Application;
        }

        public bool Initialize()
        {
            _assemblyPath = _assembly.FullName;
            _projectDirectory = Path.GetDirectoryName(_assemblyPath);
            _thumbnailsDirectory = Path.Combine(_projectDirectory, Constants.Folders.Thumbnails);
            return true;
        }

        public void Process()
        {
            SeOccurrences occurrences = null;
            HashSet<string> filePaths = new (StringComparer.OrdinalIgnoreCase);
            try
            {
                Directory.CreateDirectory(_thumbnailsDirectory);

                occurrences = _assembly.Occurrences;
                DataUtils.BuildDataForCorrectModelsView(occurrences, filePaths);
                foreach (string filePath in filePaths)
                {
                    bool isOpen = false;

                    SeDocument document = null;
                    SeWindow window = null;

                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string thumbnailPath = Path.Combine(_thumbnailsDirectory, fileName + ".jpg");

                    if (File.Exists(thumbnailPath))
                    {
                        continue;
                    }

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
                                    ReportUtils.SaveThumbnail(thumbnailPath, window);
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
