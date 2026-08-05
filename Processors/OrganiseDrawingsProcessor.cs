using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class OrganiseDrawingsProcessor
    {
        private readonly SeAssembly _assembly;
        private List<(string FileName, string Path)> _pdfFiles;
        private List<(string FileName, string Path)> _dxfFiles;
        private readonly Dictionary<string, FileData> _occurrencesData;

        private string _projectDirectory;
        private string _drawingsDirectory;
        private string _targetDirectory;

        public OrganiseDrawingsProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _pdfFiles = new List<(string FileName, string Path)>();
            _dxfFiles = new List<(string FileName, string Path)>();
            _occurrencesData = new Dictionary<string, FileData>(StringComparer.OrdinalIgnoreCase);
        }

        public bool Initialize()
        {
            if (_assembly == null || string.IsNullOrEmpty(_assembly.FullName)) { return false; }

            string assemblyFilePath = _assembly.FullName;
            _projectDirectory = Path.GetDirectoryName(assemblyFilePath);

            if (string.IsNullOrEmpty(_projectDirectory)) { return false; }

            _drawingsDirectory = Path.Combine(_projectDirectory, Constants.Folders.Drawings);
            if (!Directory.Exists(_drawingsDirectory)) { Directory.CreateDirectory(_drawingsDirectory); }

            string assemblyFileName = Path.GetFileNameWithoutExtension(assemblyFilePath);
            _targetDirectory = Path.Combine(_drawingsDirectory, assemblyFileName);
            if (!Directory.Exists(_targetDirectory)) { Directory.CreateDirectory(_targetDirectory); }

            _pdfFiles = Directory.GetFiles(_projectDirectory, "*.pdf", SearchOption.TopDirectoryOnly)
                                .Select(f => (Path.GetFileNameWithoutExtension(f), f)).ToList();

            _dxfFiles = Directory.GetFiles(_projectDirectory, "*.dxf", SearchOption.TopDirectoryOnly)
                                 .Select(f => (Path.GetFileNameWithoutExtension(f), f)).ToList();

            LoadOccurrencesData();
            return true;
        }

        public void Process()
        {
            if (_occurrencesData == null || _occurrencesData.Count == 0) { return; }

            HashSet<string> processedFileNames = new(StringComparer.OrdinalIgnoreCase);

            foreach (var item in _occurrencesData)
            {
                string fileName = item.Value.FileName;
                string type = item.Value.Type;

                if (processedFileNames.Contains(fileName)) { continue; }

                string subDirectoryName = GetTargetSubDirectory(type);
                if (string.IsNullOrEmpty(subDirectoryName)) { continue; }

                string currentTargetDirectory = Path.Combine(_targetDirectory, subDirectoryName);
                if (!Directory.Exists(currentTargetDirectory)) { Directory.CreateDirectory(currentTargetDirectory); }

                // PDF
                var matchingPdfs = _pdfFiles.Where(f => f.FileName.Equals(fileName, StringComparison.OrdinalIgnoreCase)).ToList();
                foreach (var pdf in matchingPdfs)
                {
                    string expectedPdfPath = Path.Combine(currentTargetDirectory, pdf.FileName + ".pdf");
                    File.Copy(pdf.Path, expectedPdfPath, true);
                }

                // DXF
                var matchingDxfs = _dxfFiles.Where(f => f.FileName.Equals(fileName, StringComparison.OrdinalIgnoreCase)).ToList();
                foreach (var dxf in matchingDxfs)
                {
                    string expectedDxfPath = Path.Combine(currentTargetDirectory, dxf.FileName + ".dxf");
                    File.Copy(dxf.Path, expectedDxfPath, true);
                }

                processedFileNames.Add(fileName);
            }
        }

        private string GetTargetSubDirectory(string type)
        {
            if (type == Constants.PartTypes.SheetMetal) { return Constants.Styles.SheetMetal; }
            if (type == Constants.PartTypes.Part) { return Constants.Styles.Part; }
            if (type == Constants.PartTypes.Steelmaking) { return Constants.Styles.Steelmaking; }
            if (type == Constants.PartTypes.Assembly) { return Constants.Styles.Assembly; }
            if (type == Constants.PartTypes.Commercial) { return Constants.Styles.Commercial; }
            return null;
        }

        private void LoadOccurrencesData()
        {
            SeOccurrences occurrences = null;
            try
            {
                occurrences = _assembly.Occurrences;
                DataUtils.BuildDataForOrganiseDrawings(occurrences, _occurrencesData);
            }
            finally { Helpers.ReleaseCom(ref occurrences); }
        }
    }
}