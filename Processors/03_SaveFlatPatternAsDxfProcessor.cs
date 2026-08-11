using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class SaveFlatPatternAsDxfProcessor
    {
        private readonly SeDocument _document;
        private string _documentPath;

        private string _projectDirectory;
        private string _documentName;

        private string _dxfPath;

        public SaveFlatPatternAsDxfProcessor(SeDocument document)
        {
            _document = document;
        }

        public bool Initialize()
        {
            _documentPath = _document.FullName;
            if (string.IsNullOrEmpty(_documentPath)) { MessageBox.Show("Save the file first to export flat pattern."); return false; }

            try { _document.Save(); } catch { MessageBox.Show("Cannot update document."); return false; }

            _projectDirectory = Path.GetDirectoryName(_documentPath);
            _documentName = Path.GetFileNameWithoutExtension(_documentPath);

            using PropertyUtils properties = new (_document);

            string dxfName = $"{properties.Thickness}mm_{properties.Count}szt_{properties.Material}_{_documentName}.dxf";
            string dxfPath = Path.Combine(_projectDirectory, dxfName);

            (bool isConfirmed, string editedPath) = DialogUtils.GetEditedPath(dxfPath);
            if (!isConfirmed || string.IsNullOrEmpty(editedPath)) { return false; }

            _dxfPath = editedPath;

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
                    MessageBox.Show("Cannot save DXF - missing flat pattern.");
                }
                else
                {
                    using PropertyUtils properties = new (_document);
                    properties.UpdateDxfDate();
                    models.SaveAsFlatDXFEx(_dxfPath, null, null, null, true);
                    _document.Save();
                }
            }
            finally { Helpers.ReleaseCom(ref flatPatterns); Helpers.ReleaseCom(ref models); }
        }
    }
}