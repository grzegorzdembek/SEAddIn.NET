namespace SolidEdgeAdd_In.Processors
{
    public class MakeDrawingsDirectoryProcessor
    {
        private readonly SeAssembly _assembly;
        private string _assemblyPath;
        private string _projectDirectory;

        private readonly Dictionary<string, string>  _drawings;
        private string _directory;

        public MakeDrawingsDirectoryProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _drawings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public bool Initialize()
        {
            _assemblyPath = _assembly.FullName;
            _projectDirectory = Path.GetDirectoryName(_assemblyPath);

            if (!IsDirectory_Loaded())
            {
                return false;
            }

            if (!IsDrawings_Loaded())
            {
                return false;
            }

            return true;
        }

        public void Process()
        {
            foreach (var drawing in _drawings)
            {
                string drawingName = drawing.Key;
                string drawingPath = drawing.Value; 
              
                string newDrawingPath = Path.Combine(_directory, drawingName);
                if (File.Exists(newDrawingPath))
                {
                    continue;
                }

                File.Copy(drawingPath, newDrawingPath, true);
            }
        }

        private bool IsDirectory_Loaded()
        {
            _directory = Path.Combine(_projectDirectory, "Rysunki PDF i DXF");
            Directory.CreateDirectory(_directory);
            return true;
        }

        private bool IsDrawings_Loaded()
        {
            IEnumerable<string> allFiles = Directory.EnumerateFiles(_projectDirectory, "*.*", SearchOption.AllDirectories);

            foreach (var file in allFiles)
            {
                // pominmy aby nie obciazac pamieci, zajmijmy sie nowymi itemami
                if (file.StartsWith(_directory, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!file.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) &&
                    !file.EndsWith(".dxf", StringComparison.OrdinalIgnoreCase))
                { 
                    continue; 
                }

                string fileName = Path.GetFileName(file);
                if (!_drawings.ContainsKey(fileName))
                {
                    _drawings.Add(fileName, file);
                }
            }

            if (_drawings.Count == 0)
            {
                return false;
            }

            return true;
        }
    }
}
