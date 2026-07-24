using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Helpers
{
    public class SaveAsDxfHelper
    {
        public static string GetPath(SeDocument document)
        {
            if (string.IsNullOrEmpty(document.FullName)) { MessageBox.Show("Zapisz najpierw plik w Solid Edge.", "Wymagany Zapis"); return null; }

            using var properties = new PropertyProvider(document); string name = $"{properties.Thickness}mm_{properties.Count}szt_{properties.Material}_{Path.GetFileNameWithoutExtension(document.FullName)}.dxf";

            return Path.Combine(Path.GetDirectoryName(document.FullName), name);
        }

        public static (bool isConfirmed, string editedPath) GetDecisionAndEditedPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return (false, null);

            return DialogService.GetDecisionAndEditedDxfPath(path);
        }

        public static void Save(SeDocument document, bool isConfirmed, string path)
        {
            if (!isConfirmed || string.IsNullOrEmpty(path)) return;

            SeModels models = null; SeFlatPatternModels flatPatterns = null;

            try
            {
                if (document is SePart part) { models = part.Models; flatPatterns = part.FlatPatternModels; }
                else if (document is SeSheetMetal sheetMetal) { models = sheetMetal.Models; flatPatterns = sheetMetal.FlatPatternModels; }

                if (flatPatterns == null || models == null || flatPatterns.Count == 0 || models.Count == 0) { MessageBox.Show("Nie można wykonać zapisu DXF - brak rozwinięcia."); }
                else { using var properties = new PropertyProvider(document); properties.UpdateDxfDate(); models.SaveAsFlatDXFEx(path, null, null, null, true); }
            }
            finally { CoreUtils.ReleaseCom(ref flatPatterns); CoreUtils.ReleaseCom(ref models); }
        }
    }
}