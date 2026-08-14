using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class SaveAsStepProcessor
    {
        private readonly SeDocument _document;
        private string _documentPath;

        private string _projectDirectory;
        private string _documentName;

        private string _stepPath;

        public SaveAsStepProcessor(SeDocument document)
        {
            _document = document;
        }

        public bool Initialize()
        {
            _documentPath = _document.FullName;

            if (string.IsNullOrEmpty(_documentPath))
            {
                MessageBox.Show("Save the file first to export STEP format.");
                return false;
            }

            try
            {
                _document.Save();
            }
            catch
            {
                MessageBox.Show("Cannot update document.");
                return false;
            }

            _projectDirectory = Path.GetDirectoryName(_documentPath);
            _documentName = Path.GetFileNameWithoutExtension(_documentPath);

            using PropertyUtils properties = new (_document);

            string stepName = $"{properties.MaterialName}_{properties.Count}szt_{properties.Material}_{_documentName}.step";
            string stepPath = Path.Combine(_projectDirectory, stepName);

            (bool isConfirmed, string editedPath) = DialogUtils.GetEditedPath(stepPath);

            if (!isConfirmed || string.IsNullOrEmpty(editedPath))
            {
                return false;
            }

            _stepPath = editedPath;

            return true;
        }

        public void Process()
        {
            _document.SaveAs(_stepPath);
        }
    }
}