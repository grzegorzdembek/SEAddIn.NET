using SolidEdgeAdd_In.Processors;
// Seventh Command

namespace SolidEdgeAdd_In.Commands
{
    public class SetCountPropertyCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            SeApp application = assembly.Application;

            try
            {
                SetCountPropertyProcessor processor = new (assembly);
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