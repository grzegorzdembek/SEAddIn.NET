using SolidEdgeAdd_In.Helpers;

namespace SolidEdgeAdd_In.Main
{
    public class SetCountPropertyCommand
    {
        public static void Execute(SeAssembly assembly)
        {
            try
            {
                var (isConfirmed, multiplier) = SetCountPropertyHelper.GetMultiplier(assembly);
                if (!isConfirmed) return;

                var occurrences = SetCountPropertyHelper.GetOccurrences(assembly);
                var feedback = SetCountPropertyHelper.SetAndGetFeedback(assembly, occurrences, multiplier);
                SetCountPropertyHelper.DisplayFeedback(feedback);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}