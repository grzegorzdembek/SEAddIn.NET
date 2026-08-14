namespace SolidEdgeAdd_In.Processors
{
    public class SaveAsDxfAndPdfProcessor
    {
        private readonly SeDraft _draft;
        private string _draftPath;

        private string _projectDirectory;
        private string _draftName;

        private string _dxfPath;
        private string _pdfPath;

        public SaveAsDxfAndPdfProcessor(SeDraft draft)
        {
            _draft = draft;
        }

        public bool Initialize()
        {
            _draftPath = _draft.FullName;

            if (string.IsNullOrEmpty(_draftPath))
            {
                MessageBox.Show("Save draft first to export DXF and PDF formats.");
                return false;
            }

            try
            {
                _draft.Save();
            }
            catch
            {
                MessageBox.Show("Cannot update draft.");
                return false;
            }

            _projectDirectory = Path.GetDirectoryName(_draftPath);
            _draftName = Path.GetFileNameWithoutExtension(_draftPath);

            _dxfPath = Path.Combine(_projectDirectory, _draftName + ".dxf");
            _pdfPath = Path.Combine(_projectDirectory, _draftName + ".pdf");

            return true;
        }

        public void Process()
        {
            _draft.SaveAs(_dxfPath, 14);
            _draft.SaveAs(_pdfPath, 5);
        }
    }
}