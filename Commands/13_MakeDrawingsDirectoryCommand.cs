using SolidEdgeAdd_In.Processors;

namespace SolidEdgeAdd_In.Commands
{
    public class MakeDrawingsDirectoryCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                MakeDrawingsDirectoryProcessor processor = new (assembly);

                if (processor.Initialize())
                {
                    processor.Process();
                }

                stopwatch.Stop();
                string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff");

                MessageBox.Show($"Execution time: {elapsedTime}.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception: {ex.Message}.");
            }
        }
    }
}
