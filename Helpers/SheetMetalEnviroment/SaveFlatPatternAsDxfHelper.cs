using SolidEdgeAdd_In.Utils;
using System.Windows.Forms;
using Path = System.IO.Path;

namespace SolidEdgeAdd_In.Helpers.SheetMetalEnviroment
{
    public class SaveFlatPatternAsDxfHelper
    {
        public static string GetPath(SolidEdgePart.SheetMetalDocument metalSheet)
        {
            string name =
                $"{PropertyProvider.GetThickness((SolidEdgeFramework.SolidEdgeDocument)metalSheet)}mm_" +
                $"{PropertyProvider.GetCount((SolidEdgeFramework.SolidEdgeDocument)metalSheet)}szt_" +
                $"{PropertyProvider.GetMaterial((SolidEdgeFramework.SolidEdgeDocument)metalSheet)}_" +
                $"{Path.GetFileNameWithoutExtension(metalSheet.FullName)}.dxf";

            string path = Path.Combine(Path.GetDirectoryName(metalSheet.FullName), name);
            return path;
        }

        public static (bool isConfirmed, string editedPath) GetDecisionAndEditedPath(string path)
        {
            return DialogService.GetDecisionAndEditedDxfPath(path);
        }

        public static void
          Save
          (SolidEdgePart.SheetMetalDocument metalSheet, bool isConfirmed, string path)
        {
            if (!isConfirmed) { return; }
            SolidEdgePart.Models models = null;
            SolidEdgePart.FlatPatternModels flatPatterns = null;
            try
            {
                models = metalSheet.Models;
                flatPatterns = metalSheet.FlatPatternModels;

                if (flatPatterns.Count == 0 || models.Count == 0) MessageBox.Show("Nie można wykonać zapisu DXF tego elementu, poniewż nie ma ono utworzonego rozwinięcia.");
                else
                {
                    models.SaveAsFlatDXFEx(path, null, null, null, true);
                    PropertyProvider.SetDxfDate((SolidEdgeFramework.SolidEdgeDocument)metalSheet);
                }
            }
            finally { CoreUtils.ReleaseCom(ref flatPatterns); CoreUtils.ReleaseCom(ref models); }
        }
    }
}
