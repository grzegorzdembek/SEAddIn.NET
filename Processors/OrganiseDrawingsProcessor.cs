using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class OrganiseDrawingsProcessor
    {
        private readonly SeAssembly _assembly;
        private string _projectDirectory;
        private string _drawingsDirectory;
        private string _targetDirectory;

        private List<string> _pdfFiles;
        private List<string> _dxfFiles;

        public OrganiseDrawingsProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _pdfFiles = new List<string>();
            _dxfFiles = new List<string>();
        }

        public bool Initialize()
        {
            if (_assembly == null || string.IsNullOrEmpty(_assembly.FullName))
            {
                return false;
            }

            string projectPath = _assembly.FullName;
            _projectDirectory = Path.GetDirectoryName(projectPath);

            if (string.IsNullOrEmpty(_projectDirectory))
            {
                return false;
            }

            _drawingsDirectory = Path.Combine(_projectDirectory, "Drawings");

            if (!Directory.Exists(_drawingsDirectory))
            {
                Directory.CreateDirectory(_drawingsDirectory);
            }

            string assemblyName = Path.GetFileNameWithoutExtension(projectPath);
            _targetDirectory = Path.Combine(_drawingsDirectory, assemblyName);

            if (!Directory.Exists(_targetDirectory))
            {
                Directory.CreateDirectory(_targetDirectory);
            }

            var pdfFiles = Directory.GetFiles(_projectDirectory, "*.pdf", SearchOption.TopDirectoryOnly);
            var dxfFiles = Directory.GetFiles(_projectDirectory, "*.dxf", SearchOption.TopDirectoryOnly);

            _pdfFiles = pdfFiles.Select(f => f).ToList();
            _dxfFiles = dxfFiles.Select(f => f).ToList();

            return true;
        }

        public void Process()
        {
            // Możesz teraz pisać logikę biznesową. 
            // Masz bezpośredni dostęp do _pdfFiles, _dxfFiles, _targetDirectory, itd.
            // Bez przekazywania zmiennych i parametrów z zewnątrz!
        }
    }
}