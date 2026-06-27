using SolidEdgeAdd_In.Helpers.AssemblyEnviroment; 
using System;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Main.AssemblyEnviroment
{
    public class SetCountProperty
    {
        public static void AddIn(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            try
            {
                int multiplier = SetCountPropertyHelper.GetMultiplier(assembly);
                var occurrences = SetCountPropertyHelper.GetOccurrences(assembly);
                var feedback = SetCountPropertyHelper.SetAndGetFeedback(occurrences, multiplier);
                SetCountPropertyHelper.DisplayFeedback(feedback);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}