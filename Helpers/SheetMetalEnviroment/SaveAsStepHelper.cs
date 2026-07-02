using SolidEdgeAdd_In.Utils;
using System.IO;

namespace SolidEdgeAdd_In.Helpers.SheetMetalEnviroment
{
    public class SaveAsStepHelper
    {
        public static string GetPath(SolidEdgePart.SheetMetalDocument metalSheet)
        {
            string name =
                $"{PropertyProvider.GetMaterialName((SolidEdgeFramework.SolidEdgeDocument)metalSheet)}_" +
                $"{PropertyProvider.GetCount((SolidEdgeFramework.SolidEdgeDocument)metalSheet)}szt_" +
                $"{PropertyProvider.GetMaterial((SolidEdgeFramework.SolidEdgeDocument)metalSheet)}_" +
                $"{Path.GetFileNameWithoutExtension(metalSheet.FullName)}.step";

            string path = Path.Combine(Path.GetDirectoryName(metalSheet.FullName), name);
            return path;
        }

        public static (bool isConfirmed, string editedPath) GetDecisionAndEditedPath(string path)
        {
            return DialogService.GetDecisionAndEditedStepPath(path);
        }

        public static void
            Save
            (SolidEdgePart.SheetMetalDocument metalSheet, bool isConfirmed, string path)
        {
            if (!isConfirmed) { return; }
            metalSheet.SaveAs(path);
        }
    }
}
