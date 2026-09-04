using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class OrganiseDrawingsProcessor
    {
        private readonly SeAssembly _assembly;

        private readonly Dictionary<string, FileData> _data;

        private string _assemblyFilePath;
        private string _projectDirectory;

        private string _drawingsDirectory;

        private string _targetDirectory;

        private Dictionary<string, string> _pdfFiles;
        private Dictionary<string, string> _dxfFiles;

        public OrganiseDrawingsProcessor(SeAssembly assembly)
        {
            _assembly = assembly;

            _data = new (StringComparer.OrdinalIgnoreCase);
        }

        public bool Initialize()
        {
            _assemblyFilePath = _assembly.FullName;
            _projectDirectory = Path.GetDirectoryName(_assemblyFilePath);

            if (!IsDrawingsDirectory_Loaded())
            {
                return false;
            }

            if (!IsTargetDirectory_Loaded())
            {
                return false;
            }

            if (!IsFiles_Loaded())
            {
                return false;
            }

            if (!IsData_Loaded())
            {
                return false;
            }

            return true;
        }

        public void Process()
        {
            HashSet<string> processedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in _data)
            {
                string fileName = item.Value.Name;
                string type = item.Value.Type;

                if (processedFileNames.Contains(fileName))
                {
                    continue;
                }

                string subDirectoryName = GetTargetSubDirectory(type);

                if (string.IsNullOrEmpty(subDirectoryName))
                {
                    continue;
                }

                bool hasPdf = _pdfFiles.TryGetValue(fileName, out string pdfPath);
                bool hasDxf = _dxfFiles.TryGetValue(fileName, out string dxfPath);

                if (!hasPdf && !hasDxf)
                {
                    processedFileNames.Add(fileName);
                    continue;
                }

                string currentTargetDirectory = Path.Combine(_targetDirectory, subDirectoryName);
                Directory.CreateDirectory(currentTargetDirectory);

                if (hasPdf)
                {
                    string expectedPdfPath = Path.Combine(currentTargetDirectory, fileName + ".pdf");
                    File.Copy(pdfPath, expectedPdfPath, true);
                }

                if (hasDxf)
                {
                    string expectedDxfPath = Path.Combine(currentTargetDirectory, fileName + ".dxf");
                    File.Copy(dxfPath, expectedDxfPath, true);
                }

                processedFileNames.Add(fileName);
            }
        }

        private string GetTargetSubDirectory(string type)
        {
            if (type == Constants.PartTypes.SheetMetal)
            {
                return Constants.Styles.SheetMetal;
            }

            if (type == Constants.PartTypes.Part)
            {
                return Constants.Styles.Part;
            }

            if (type == Constants.PartTypes.Steelmaking)
            {
                return Constants.Styles.Steelmaking;
            }

            if (type == Constants.PartTypes.Assembly)
            {
                return Constants.Styles.Assembly;
            }

            if (type == Constants.PartTypes.Commercial)
            {
                return Constants.Styles.Commercial;
            }

            return null;
        }

        private bool IsDrawingsDirectory_Loaded()
        {
            _drawingsDirectory = Path.Combine(_projectDirectory, Constants.Folders.Drawings);
            Directory.CreateDirectory(_drawingsDirectory);

            return true;
        }

        private bool IsTargetDirectory_Loaded()
        {
            string assemblyFileName = Path.GetFileNameWithoutExtension(_assemblyFilePath);
            _targetDirectory = Path.Combine(_drawingsDirectory, assemblyFileName);
            Directory.CreateDirectory(_targetDirectory);

            return true;
        }

        private bool IsFiles_Loaded()
        {
            _pdfFiles = Directory.EnumerateFiles(_projectDirectory, "*.pdf", SearchOption.TopDirectoryOnly)
                .ToDictionary(f => Path.GetFileNameWithoutExtension(f), f => f, StringComparer.OrdinalIgnoreCase);

            _dxfFiles = Directory.EnumerateFiles(_projectDirectory, "*.dxf", SearchOption.TopDirectoryOnly)
                .ToDictionary(f => Path.GetFileNameWithoutExtension(f), f => f, StringComparer.OrdinalIgnoreCase);

            if (_pdfFiles.Count == 0 && _dxfFiles.Count == 0)
            {
                MessageBox.Show("Nie znaleziono plików PDF ani DXF w katalogu projektu.");
                return false;
            }

            return true;
        }

        private bool IsData_Loaded()
        {
            SeOccurrences occurrences = null;

            try
            {
                occurrences = _assembly.Occurrences;
                DataUtils.BuildDataForOrganiseDrawings(occurrences, _data);
            }
            finally
            {
                Helpers.ReleaseCom(ref occurrences);
            }

            if (_data.Count == 0)
            {
                MessageBox.Show("Nie znaleziono wystąpień do przetworzenia.");
                return false;
            }

            return true;
        }
    }
}