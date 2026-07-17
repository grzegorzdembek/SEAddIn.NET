using SolidEdgeAdd_In.Helpers;

namespace SolidEdgeAdd_In.Main
{
    public class SaveAsStepCommand
    {
        public static void Execute(SeDocument document)
        {
            try
            {
                string path = SaverAsStepHelper.GetPath(document);
                (bool isConfirmed, string editedPath) = SaverAsStepHelper.GetDecisionAndEditedPath(path);
                SaverAsStepHelper.Save(document, isConfirmed, editedPath);
            }
            catch (Exception ex) { MessageBox.Show($"Exception: {ex.Message}"); }
        }
    }
}
