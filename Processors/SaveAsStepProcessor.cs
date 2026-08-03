using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class SaveAsStepProcessor
    {
        private readonly SeDocument _document;
        private string _path;

        public SaveAsStepProcessor(SeDocument document)
        {
            _document = document;
        }

        public bool Initialize()
        {
            if (string.IsNullOrEmpty(_document.FullName))
            {
                MessageBox.Show("Save the file in Solid Edge first to export STEP format.", "Save Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            using var properties = new PropertyProvider(_document);
            string name = $"{properties.MaterialName}_{properties.Count}pcs_{properties.Material}_{Path.GetFileNameWithoutExtension(_document.FullName)}.step";

            string initialPath = Path.Combine(Path.GetDirectoryName(_document.FullName), name);

            var result = DialogService.GetDecisionAndEditedStepPath(initialPath);

            if (!result.isConfirmed || string.IsNullOrEmpty(result.stepPath))
            {
                return false;
            }

            _path = result.stepPath;
            return true;
        }

        public void Process()
        {
            _document.SaveAs(_path);
        }
    }
}