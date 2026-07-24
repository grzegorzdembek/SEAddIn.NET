using Helper = SolidEdgeAdd_In.Helpers.SaveAsStepHelper;

namespace SolidEdgeAdd_In.Commands
{
    public class SaveAsStepCommand
    {
        public static void Execute(SeDocument document)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            try
            {
                string path = Helper.GetPath(document); (bool isConfirmed, string editedPath) = Helper.GetDecisionAndEditedPath(path);

                Helper.Save(document, isConfirmed, editedPath);

                stopwatch.Stop(); string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff"); MessageBox.Show($"Czas wykonywania: {elapsedTime}", "Zakończono", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}
