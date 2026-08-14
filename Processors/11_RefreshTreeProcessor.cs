using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class RefreshTreeProcessor
    {
        private readonly SeAssembly _assembly;

        public RefreshTreeProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
        }

        public bool Initialize()
        {
            return true;
        }

        public void Process()
        {
            int previousCount = -1;
            int currentCount = 0;
            int maxAttempts = 3;
            int attempts = 0;

            do
            {
                previousCount = currentCount;
                attempts++;

                HashSet<string> paths = new (StringComparer.OrdinalIgnoreCase);

                SeOccurrences occurrences = null;
                try
                {
                    occurrences = _assembly.Occurrences;
                    DataUtils.RefreshTree(occurrences, paths);
                    currentCount = paths.Count;
                }
                finally
                {
                    Helpers.ReleaseCom(ref occurrences);
                }
            }
            while (currentCount > previousCount && attempts < maxAttempts);
        }
    }
}