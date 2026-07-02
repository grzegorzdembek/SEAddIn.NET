using SolidEdgeAdd_In.Utils;
using System.IO;

namespace SolidEdgeAdd_In.Helpers.PartEnviroment
{
    public class SaveAsStepHelper
    {
        public static string GetPath(SolidEdgePart.PartDocument part)
        {
            string name =
                 $"{PropertyProvider.GetMaterialName((SolidEdgeFramework.SolidEdgeDocument)part)}_" +
                 $"{PropertyProvider.GetCount((SolidEdgeFramework.SolidEdgeDocument)part)}szt_" +
                 $"{PropertyProvider.GetMaterial((SolidEdgeFramework.SolidEdgeDocument)part)}_" +
                 $"{Path.GetFileNameWithoutExtension(part.FullName)}.step";

            string path = Path.Combine(Path.GetDirectoryName(part.FullName), name);
            return path;
        }

        public static (bool isConfirmed, string editedPath) GetDecisionAndEditedPath(string path)
        {
            return DialogService.GetDecisionAndEditedStepPath(path);
        }

        public static void
            Save
            (SolidEdgePart.PartDocument part, bool isConfirmed, string path)
        {
            if (!isConfirmed) { return; }
            part.SaveAs(path);
        }
    }
}
