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

        private List<(string FileName, string Path)> _pdfFiles;
        private List<(string FileName, string Path)> _dxfFiles;
        
        public OrganiseDrawingsProcessor(SeAssembly assembly)
        {
            _assembly = assembly;

            _data = new Dictionary<string, FileData>(StringComparer.OrdinalIgnoreCase);
        }

        public bool Initialize()
        {
            _assemblyFilePath = _assembly.FullName;
            _projectDirectory = Path.GetDirectoryName(_assemblyFilePath);

            _drawingsDirectory = Path.Combine(_projectDirectory, Constants.Folders.Drawings);
            Directory.CreateDirectory(_drawingsDirectory);

            string assemblyFileName = Path.GetFileNameWithoutExtension(_assemblyFilePath);
            _targetDirectory = Path.Combine(_drawingsDirectory, assemblyFileName);
            Directory.CreateDirectory(_targetDirectory);

            _pdfFiles = Directory.GetFiles(_projectDirectory, "*.pdf", SearchOption.TopDirectoryOnly)
                               .Select(f => (Path.GetFileNameWithoutExtension(f), f)).ToList();

            _dxfFiles = Directory.GetFiles(_projectDirectory, "*.dxf", SearchOption.TopDirectoryOnly)
                                 .Select(f => (Path.GetFileNameWithoutExtension(f), f)).ToList();

            if (_pdfFiles.Count == 0 && _dxfFiles.Count == 0) { MessageBox.Show("No PDF or DXF files found in the project directory.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning); return false; }

            LoadOccurrencesData();

            if (_data.Count == 0) { MessageBox.Show("No occurrences found to process.", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Information); return false; }

            return true;
        }

        public void Process()
        {
            HashSet<string> processedFileNames = new(StringComparer.OrdinalIgnoreCase);

            foreach (var item in _data)
            {
                string fileName = item.Value.Name;
                string type = item.Value.Type;

                if (processedFileNames.Contains(fileName)) { continue; }

                string subDirectoryName = GetTargetSubDirectory(type);
                if (string.IsNullOrEmpty(subDirectoryName)) { continue; }

                var matchingPdfs = _pdfFiles.Where(f => f.FileName.Equals(fileName, StringComparison.OrdinalIgnoreCase));
                var matchingDxfs = _dxfFiles.Where(f => f.FileName.Equals(fileName, StringComparison.OrdinalIgnoreCase));

                if (!matchingPdfs.Any() && !matchingDxfs.Any()) { processedFileNames.Add(fileName); continue; }

                string currentTargetDirectory = Path.Combine(_targetDirectory, subDirectoryName);
                Directory.CreateDirectory(currentTargetDirectory);

                foreach (var pdf in matchingPdfs) { string expectedPdfPath = Path.Combine(currentTargetDirectory, pdf.FileName + ".pdf"); File.Copy(pdf.Path, expectedPdfPath, true); }

                foreach (var dxf in matchingDxfs) { string expectedDxfPath = Path.Combine(currentTargetDirectory, dxf.FileName + ".dxf"); File.Copy(dxf.Path, expectedDxfPath, true); }

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
                DataUtils.BuildDataForOrganiseDrawings(occurrences, _data);
            }
            finally { Helpers.ReleaseCom(ref occurrences); }
        }
    }
}