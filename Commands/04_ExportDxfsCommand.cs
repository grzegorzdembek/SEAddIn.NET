using SolidEdgeAdd_In.Processors;
// Fourth Command

namespace SolidEdgeAdd_In.Commands
{
    public class ExportDxfsCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            SeApp application = assembly.Application;
            try
            {
                ExportDxfsProcessor processor = new(assembly);
                if (processor.Initialize()) 
                {
                    application.DelayCompute = true;
                    application.ScreenUpdating = false;
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
                application.ScreenUpdating = true; 
            }
        }
    }
}