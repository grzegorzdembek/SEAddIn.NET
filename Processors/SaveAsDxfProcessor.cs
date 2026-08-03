using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class SaveAsDxfProcessor
    {
        private readonly SeDocument _document;
        private string _path;

        public SaveAsDxfProcessor(SeDocument document)
        {
            _document = document;
        }

        public bool Initialize()
        {
            if (string.IsNullOrEmpty(_document.FullName))
            {
                MessageBox.Show("Save the file in Solid Edge first.", "Save Required");
                return false;
            }

            using var properties = new PropertyProvider(_document);
            string name = $"{properties.Thickness}mm_{properties.Count}pcs_{properties.Material}_{Path.GetFileNameWithoutExtension(_document.FullName)}.dxf";

            string initialPath = Path.Combine(Path.GetDirectoryName(_document.FullName), name);

            var result = DialogService.GetDecisionAndEditedDxfPath(initialPath);

            if (!result.isConfirmed || string.IsNullOrEmpty(result.dxfPath))
            {
                return false;
            }

            _path = result.dxfPath;
            return true;
        }

        public void Process()
        {
            SeModels models = null;
            SeFlatPatternModels flatPatterns = null;

            try
            {
                if (_document is SePart part)
                {
                    models = part.Models;
                    flatPatterns = part.FlatPatternModels;
                }
                else if (_document is SeSheetMetal sheetMetal)
                {
                    models = sheetMetal.Models;
                    flatPatterns = sheetMetal.FlatPatternModels;
                }

                if (flatPatterns == null || models == null || flatPatterns.Count == 0 || models.Count == 0)
                {
                    MessageBox.Show("Cannot save DXF - missing flat pattern.");
                }
                else
                {
                    using var properties = new PropertyProvider(_document);
                    properties.UpdateDxfDate();
                    models.SaveAsFlatDXFEx(_path, null, null, null, true);
                    _document.Save();
                }
            }
            finally
            {
                CoreUtils.ReleaseCom(ref flatPatterns);
                CoreUtils.ReleaseCom(ref models);
            }
        }
    }
}