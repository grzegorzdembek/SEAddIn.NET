using SolidEdgeAdd_In.Processors;

namespace SolidEdgeAdd_In.Commands
{
    public class SaveAsDxfAndPdfCommand
    {
        public static void Execute(SeDraft draft)
        {
            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();

                SaveAsDxfAndPdfProcessor processor = new (draft);
                if (processor.Initialize()) { processor.Process(); }

                stopwatch.Stop();
                string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff");

                MessageBox.Show($"Execution time: {elapsedTime}", 
                    "Completed", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}