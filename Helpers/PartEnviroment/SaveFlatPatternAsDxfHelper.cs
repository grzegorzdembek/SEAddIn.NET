using SolidEdgeAdd_In.Utils;
using System.Windows.Forms;
using Path = System.IO.Path;

namespace SolidEdgeAdd_In.Helpers.PartEnviroment
{
    public class SaveFlatPatternAsDxfHelper
    {
        public static string GetPath(SolidEdgePart.PartDocument part)
        {
            string name =
                 $"{PropertyProvider.GetThickness((SolidEdgeFramework.SolidEdgeDocument)part)}mm_" +
                 $"{PropertyProvider.GetCount((SolidEdgeFramework.SolidEdgeDocument)part)}szt_" +
                 $"{PropertyProvider.GetMaterial((SolidEdgeFramework.SolidEdgeDocument)part)}_" +
                 $"{Path.GetFileNameWithoutExtension(part.FullName)}.dxf";

            string path = Path.Combine(Path.GetDirectoryName(part.FullName), name);
            return path;
        }

        public static (bool isConfirmed, string editedPath) GetDecisionAndEditedPath(string path)
        {
            return DialogService.GetDecisionAndEditedDxfPath(path);
        }

        public static void
          Save
          (SolidEdgePart.PartDocument part, bool isConfirmed, string path)
        {
            if (!isConfirmed) { return; }
            SolidEdgePart.Models models = null;
            SolidEdgePart.FlatPatternModels flatPatterns = null;
            try
            {
                models = part.Models;
                flatPatterns = part.FlatPatternModels;

                if (flatPatterns.Count == 0 || models.Count == 0) MessageBox.Show("Nie można wykonać zapisu DXF tego elementu, poniewż nie ma ono utworzonego rozwinięcia.");
                else
                {
                    models.SaveAsFlatDXFEx(path, null, null, null, true);
                    PropertyProvider.SetDxfDate((SolidEdgeFramework.SolidEdgeDocument)part);
                }
            }
            finally { CoreUtils.ReleaseCom(ref flatPatterns); CoreUtils.ReleaseCom(ref models); }
        }
    }
}