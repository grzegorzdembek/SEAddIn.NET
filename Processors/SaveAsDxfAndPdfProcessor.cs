namespace SolidEdgeAdd_In.Processors
{
    public class SaveAsDxfAndPdfProcessor
    {
        private readonly SeDraft _draft;

        private string _dxfFilePath;
        private string _pdfFilePath;

        public SaveAsDxfAndPdfProcessor(SeDraft draft)
        {
            _draft = draft;
        }

        public bool Initialize()
        {
            string documentFilePath = _draft.FullName;

            if (string.IsNullOrEmpty(documentFilePath))
            {
                MessageBox.Show("Save the drawing first to export DXF and PDF formats.", "Save Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string fileName = Path.GetFileNameWithoutExtension(documentFilePath);
            string projectDirectory = Path.GetDirectoryName(documentFilePath);
            
            _dxfFilePath = Path.Combine(projectDirectory, fileName + ".dxf");
            _pdfFilePath = Path.Combine(projectDirectory, fileName + ".pdf");

            return true;
        }

        public void Process()
        {
            _draft.SaveAs(_dxfFilePath, 14);
            _draft.SaveAs(_pdfFilePath, 5);
        }
    }
}