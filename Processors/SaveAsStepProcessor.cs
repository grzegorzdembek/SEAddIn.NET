using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class SaveAsStepProcessor
    {
        private readonly SeDocument _document;
        private string _stepFilePath;

        public SaveAsStepProcessor(SeDocument document)
        {
            _document = document;
        }

        public bool Initialize()
        {
            string documentFilePath = _document.FullName;

            if (string.IsNullOrEmpty(documentFilePath))
            {
                MessageBox.Show("Save the file in Solid Edge first to export STEP format.", "Save Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string fileName = Path.GetFileNameWithoutExtension(documentFilePath);
            string projectDirectory = Path.GetDirectoryName(documentFilePath);

            using PropertyUtils properties = new(_document);
            string stepFileName = $"{properties.MaterialName}_{properties.Count}szt_{properties.Material}_{fileName}.step";
            string stepFilePath = Path.Combine(projectDirectory, stepFileName);

            (bool isConfirmed, string editedPath) = DialogUtils.GetEditedPath(stepFilePath);
            if (!isConfirmed || string.IsNullOrEmpty(editedPath)) { return false; }

            _stepFilePath = editedPath;

            return true;
        }

        public void Process()
        {
            _document.SaveAs(_stepFilePath);
        }
    }
}