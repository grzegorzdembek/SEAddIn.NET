using SolidEdgeAdd_In.Processors;
// Sixth Command

namespace SolidEdgeAdd_In.Commands
{
    public class ExportOccurrencesListCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            SeApp application = assembly.Application;
            try
            {
                ExportOccurrencesListProcessor processor = new (assembly);
                if (processor.Initialize()) 
                { 
                    application.DelayCompute = true;
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
            finally 
            { 
                application.DelayCompute = false; 
            }
        }
    }
}
