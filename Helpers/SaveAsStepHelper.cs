using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Helpers
{
    public class SaveAsStepHelper
    {
        public static string GetPath(SeDocument document)
        {
            if (string.IsNullOrEmpty(document.FullName)) { MessageBox.Show("Zapisz najpierw plik w Solid Edge, aby móc wyeksportować format STEP.", "Wymagany Zapis", MessageBoxButtons.OK, MessageBoxIcon.Warning); return null; }

            using var properties = new PropertyProvider(document);       
            string name = $"{properties.MaterialName}_{properties.Count}szt_{properties.Material}_{Path.GetFileNameWithoutExtension(document.FullName)}.step";

            return Path.Combine(Path.GetDirectoryName(document.FullName), name);
        }

        public static (bool isConfirmed, string editedPath) GetEditedPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return (false, null);

            return DialogService.GetDecisionAndEditedStepPath(path);
        }

        public static void Save(SeDocument document, bool isConfirmed, string path)
        { 
            if (!isConfirmed || string.IsNullOrEmpty(path)) return; 
            
            document.SaveAs(path); 
        }
    }
}