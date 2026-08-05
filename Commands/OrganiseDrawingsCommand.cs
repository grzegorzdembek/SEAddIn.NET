using SolidEdgeAdd_In.Processors;

namespace SolidEdgeAdd_In.Commands
{
    public class OrganiseDrawingsCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();

                OrganiseDrawingsProcessor processor = new (assembly);
                if (processor.Initialize()) { processor.Process(); }

                stopwatch.Stop();
                string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff");

                MessageBox.Show($"Execution time: {elapsedTime}", 
                    "Completed", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Information);
            }
            catch (Exception ex){ MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}
