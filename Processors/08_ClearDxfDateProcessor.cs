using Microsoft.Office.Interop.Excel;
using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class ClearDxfDateProcessor
    {
        private readonly SeAssembly _assembly;

        private string _assemblyPath;
        private string _projectDirectory;

        private readonly Logger _logger;
       
        public ClearDxfDateProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _logger = new Logger();
        }

        public bool Initialize()
        {
            _assemblyPath = _assembly.FullName;
            _projectDirectory = Path.GetDirectoryName(_assemblyPath);

            return true;
        }

        public void Process()
        {
            HashSet<string> processedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            SeOccurrences occurrences = null;

            try
            {
                occurrences = _assembly.Occurrences;

                DataUtils.ClearDxfDates(occurrences, processedPaths, _logger);
            }
            catch (Exception ex)
            {
                _logger.LogError("Main Process", $"Critical error during execution: {ex.Message}");
                throw;
            }
            finally
            {
                Helpers.ReleaseCom(ref occurrences);

                if (!string.IsNullOrEmpty(_projectDirectory))
                {
                    _logger.SaveReport(_projectDirectory);
                }
            }
        }
    }
}