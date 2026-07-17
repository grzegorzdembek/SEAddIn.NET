using SolidEdgeAdd_In.Helpers; 

namespace SolidEdgeAdd_In.Main
{
    public class SaveAsDxfCommand
    {
        public static void Execute(SeDocument document)
        {
            try
            {
                string path = SaveAsDxfHelper.GetPath(document);
                (bool isConfirmed, string editedPath) = SaveAsDxfHelper.GetDecisionAndEditedPath(path);
                SaveAsDxfHelper.Save(document, isConfirmed, editedPath);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}
