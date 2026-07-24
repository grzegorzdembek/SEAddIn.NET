using Helper = SolidEdgeAdd_In.Helpers.ClearDxfDateHelper;

namespace SolidEdgeAdd_In.Commands
{
    public class ClearDxfDateCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            Stopwatch stopwatch = Stopwatch.StartNew(); SeApp application = assembly.Application;

            try
            {
                application.DelayCompute = true; application.ScreenUpdating = false;

                Helper.ProcessClearing(assembly);

                stopwatch.Stop(); string elapsedTime = stopwatch.Elapsed.ToString(@"mm\:ss\.fff"); MessageBox.Show($"Czas wykonywania: {elapsedTime}", "Zakończono", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
            finally { application.DelayCompute = false; application.ScreenUpdating = true; }
        }
    }
}
