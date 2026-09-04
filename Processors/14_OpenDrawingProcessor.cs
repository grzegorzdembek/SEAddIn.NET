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
                                ProcessSelectedItem(selectedItem);
                            }
                            finally
                            {
                                Helpers.ReleaseCom(ref selectedItem);
                            }
                        }
                     
                        if (_namesToFind.Count == 0) // zalozmy ze chcial otworzyc rysnek zlozenia 
                        {
                            _namesToFind.Add(_documentName);
                        }
                    }
                    else // jesli nic nie zaznaczyl, chce otworzyc rysunek dla zlozenia
                    {
                        _namesToFind.Add(_documentName);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Wystąpił błąd podczas przetwarzania dokumentu: {ex.Message}");
                    return false;
                }
                finally
                {
                    Helpers.ReleaseCom(ref selectSet);
                }               
            }
            else
            {
                _namesToFind.Add(_documentName);
            }

            return true;
        }

        private void ProcessSelectedItem(object item)
        {
            if (item == null)
            {
                return;
            }

            object extractedObject = null;
            SeDocument occurrenceDocument = null;

            try
            {
                SeOccurrence occurrence = null;
                string comTypeName = System.ComponentModel.TypeDescriptor.GetClassName(item);

                if (comTypeName == "Reference")
                {
                    dynamic dynRef = item;
                    extractedObject = dynRef.Object;
                    occurrence = extractedObject as SeOccurrence;
                }
                else
                {
                    occurrence = item as SeOccurrence;
                }

                if (occurrence == null)
                {
                    return;
                }

                string occurrencePath = occurrence.OccurrenceFileName;

                if (string.IsNullOrEmpty(occurrencePath))
                {
                    return;
                }

                string occurrenceName = Path.GetFileNameWithoutExtension(occurrencePath);
                _namesToFind.Add(occurrenceName);

                occurrenceDocument = (SeDocument)occurrence.OccurrenceDocument;

                if (!occurrencePath.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) || occurrenceDocument is not SeAssembly subAssembly)
                {
                    return;
                }

                SeOccurrences subOccurrences = null;
                try
                {
                    subOccurrences = subAssembly.Occurrences;
                    for (int i = 1; i <= subOccurrences.Count; i++)
                    {
                        object childItem = null;
                        try
                        {
                            childItem = subOccurrences.Item(i);
                            ProcessSelectedItem(childItem);
                        }
                        finally
                        {
                            Helpers.ReleaseCom(ref childItem);
                        }
                    }
                }
                finally
                {
                    Helpers.ReleaseCom(ref subOccurrences);
                }
            }
            finally
            {
                Helpers.ReleaseCom(ref occurrenceDocument);

                if (extractedObject != null)
                {
                    Helpers.ReleaseCom(ref extractedObject);
                }
            }
        }
    }
}
