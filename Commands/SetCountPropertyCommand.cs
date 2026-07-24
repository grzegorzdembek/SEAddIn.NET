using Helper = SolidEdgeAdd_In.Helpers.SetCountPropertyHelper;

namespace SolidEdgeAdd_In.Commands
{
    public class SetCountPropertyCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                var (isConfirmed, multiplier) = Helper.GetMultiplier(assembly); if (!isConfirmed) return;

                var occurrences = Helper.GetData(assembly); var feedback = Helper.SetAndGetFeedback(assembly, occurrences, multiplier);

                Helper.DisplayFeedback(feedback);

                stopwatch.Stop(); string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff"); MessageBox.Show($"Czas wykonywania: {elapsedTime}", "Zakończono", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}