using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class ClearDxfDateProcessor
    {
        private readonly SeAssembly _assembly;

        public ClearDxfDateProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
        }

        public void Process()
        {
            HashSet<string> processedPaths = new(StringComparer.OrdinalIgnoreCase);
            SeOccurrences occurrences = null;
            try
            {
                occurrences = _assembly.Occurrences;
                DataUtils.ClearDxfDates(occurrences, processedPaths);
            }
            finally { Helpers.ReleaseCom(ref occurrences); }
        }
    }
}