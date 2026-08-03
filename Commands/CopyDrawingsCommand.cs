using SolidEdgeAdd_In.Processors;

namespace SolidEdgeAdd_In.Commands
{
    public class CopyDrawingsCommand
    {
        public static void Execute(SeDocument document)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                var processor = new CopyDrawingsProcessor(document);

                if (processor.Initialize())
                {
                    processor.Process();
                }

                stopwatch.Stop();
                string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff");
                MessageBox.Show($"Execution time: {elapsedTime}", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception: {ex.Message}");
            }
        }
    }
}
