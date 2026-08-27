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
            List<string> missingPdfs = new ();
            var uniqueNames = _namesToFind.Distinct().ToList();

            foreach (string name in uniqueNames)
            {
                string pdfPath = Path.Combine(_projectDirectory, name + ".pdf");
                if (File.Exists(pdfPath))
                {
                    System.Diagnostics.Process.Start(pdfPath);
                }
                else
                { 
                    missingPdfs.Add(name);
                }
            }

            if (missingPdfs.Count > 0)
            {
                string missingList = string.Join(System.Environment.NewLine, missingPdfs);
                MessageBox.Show($"Nie znaleziono rysunków (.pdf) dla {missingPdfs.Count} elementów:\n\n{missingList}");                              
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
                                    string occurrencePath = occurrence.OccurrenceFileName;
                                    string occurrenceName = Path.GetFileNameWithoutExtension(occurrencePath);

                                    _namesToFind.Add(occurrenceName);
                                }
                            }
                            finally
                            {
                                Helpers.ReleaseCom(ref selectedItem);
                            }
                        }

                        if (_namesToFind.Count == 0)
                        {
                            MessageBox.Show("Nie znaleziono możliwych rysunków do otwarcia.");
                            return false;
                        }

                        return true;
                    }
                    else
                    {
                        MessageBox.Show("Brak zaznaczonych elementów.");
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
