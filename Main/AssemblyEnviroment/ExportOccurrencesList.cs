using System;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Main.AssemblyEnviroment
{
    public class ExportOccurrencesList
    {
        public static void
           AddIn
           (SolidEdgeAssembly.AssemblyDocument assembly)
        {
            try
            {
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }       
    }
}
