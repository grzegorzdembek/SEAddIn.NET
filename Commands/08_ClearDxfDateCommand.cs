using SolidEdgeAdd_In.Processors;
// Eighth Command

namespace SolidEdgeAdd_In.Commands
{
    public class ClearDxfDateCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            SeApp application = assembly.Application;

            try
            {
                ClearDxfDateProcessor processor = new (assembly);
                processor.Process();

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
