using SolidEdgeAdd_In.Processors;
// Ninth Command

namespace SolidEdgeAdd_In.Commands
{
    public class CopyDrawingsCommand
    {
        public static void Execute(SeDocument document)
        {
            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();

                CopyDrawingsProcessor processor = new (document);
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
