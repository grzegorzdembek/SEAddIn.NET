using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class SaveFlatPatternAsDxfProcessor
    {
        private readonly SeDocument _document;
        private string _dxfFilePath;

        public SaveFlatPatternAsDxfProcessor(SeDocument document)
        {
            _document = document;
        }

        public bool Initialize()
        {
            string documentFilePath = _document.FullName;

            if (string.IsNullOrEmpty(documentFilePath))
            {
                MessageBox.Show("Save the file in Solid Edge first to export flat pattern.", "Save Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string fileName = Path.GetFileNameWithoutExtension(documentFilePath);
            string projectDirectory = Path.GetDirectoryName(documentFilePath);

            using var properties = new PropertyUtils(_document);
            string dxfFileName = $"{properties.Thickness}mm_{properties.Count}szt_{properties.Material}_{fileName}.dxf";
            string dxfFilePath = Path.Combine(projectDirectory, dxfFileName);

            (bool isConfirmed, string editedPath) = DialogUtils.GetEditedPath(dxfFilePath);
            if (!isConfirmed || string.IsNullOrEmpty(editedPath)) { return false; }

            _dxfFilePath = editedPath;

            return true;
        }

        public void Process()
        {
            SeModels models = null;
            SeFlatPatternModels flatPatterns = null;

            try
            {
                if (_document is SePart part) { models = part.Models; flatPatterns = part.FlatPatternModels; }
                else if (_document is SeSheetMetal sheetMetal) { models = sheetMetal.Models; flatPatterns = sheetMetal.FlatPatternModels; }

                if (flatPatterns == null || models == null || flatPatterns.Count == 0 || models.Count == 0)
                {
                    MessageBox.Show("Cannot save DXF - missing flat pattern.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    using var properties = new PropertyUtils(_document);
                    properties.UpdateDxfDate();
                    models.SaveAsFlatDXFEx(_dxfFilePath, null, null, null, true);
                    _document.Save();
                }
            }
            finally { Helpers.ReleaseCom(ref flatPatterns); Helpers.ReleaseCom(ref models); }
        }
    }
}