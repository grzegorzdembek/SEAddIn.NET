using SolidEdgeAdd_In.Processors;

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
                application.DelayCompute = true;
                application.ScreenUpdating = false;

                SetCountPropertyProcessor processor = new (assembly);
                if (processor.Initialize()) { processor.Process(); }

                stopwatch.Stop();
                string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff");

                MessageBox.Show($"Execution time: {elapsedTime}", 
                    "Completed", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
            finally { application.DelayCompute = false; application.ScreenUpdating = true; }
        }
    }
}