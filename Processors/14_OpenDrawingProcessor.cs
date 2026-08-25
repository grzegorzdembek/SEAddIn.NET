using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class OpenDrawingProcessor
    {
        private readonly SeDocument _document;

        private string _documentPath;  
        private string _projectDirectory;
        private string _documentName;

        private SeAssembly _assembly;
        List<string> _namesToFind;

        public OpenDrawingProcessor(SeDocument document)
        {
            _document = document;
        }

        public bool Initialize()
        {
            _documentPath = _document.FullName;
            _projectDirectory = Path.GetDirectoryName(_documentPath);
            _documentName = Path.GetFileNameWithoutExtension(_documentPath);

            if (!IsLoaded_Document()) 
            { 
                return false; 
            }

            return true;
        }

        public void Process()
        {
            foreach (string name in _namesToFind)
            {             
                string drawingPath = Path.Combine(_projectDirectory, name + ".pdf");
                if (File.Exists(drawingPath))
                {
                    System.Diagnostics.Process.Start(drawingPath);
                }
            }
        }

        public bool IsLoaded_Document()
        {
            _namesToFind = new();

            if (_document is SeAssembly assemblyDocument)
            {
                _assembly = assemblyDocument;
                SelectSet selectSet = null;
                try
                {
                    selectSet = _assembly.SelectSet;

                    if (selectSet.Count > 0)
                    {
                        for (int i = 1; i <= selectSet.Count; i++)
                        {
                            object selectedItem = null;
                            try
                            {
                                selectedItem = selectSet.Item(i);
                                if (selectedItem != null && selectedItem is SeOccurrence occurrence)
                                {
                                    string rawOccurrenceName = occurrence.Name;
                                    string nameWithoutInstance = rawOccurrenceName.Contains(":") ? rawOccurrenceName.Split(':')[0] : rawOccurrenceName;
                                    string OccurrenceName = Path.GetFileNameWithoutExtension(nameWithoutInstance);

                                    _namesToFind.Add(OccurrenceName);
                                }
                            }
                            finally
                            {
                                Helpers.ReleaseCom(ref selectedItem);
                            }
                        }

                        if (_namesToFind.Count == 0)
                        {
                            return false;
                        }

                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                finally
                {
                    Helpers.ReleaseCom(ref selectSet);
                }               
            }
            else
            {
                _namesToFind.Add(_documentName);
                return true;
            }
        }
    }
}
