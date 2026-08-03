using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class SaveAsDxfAndPdfProcessor
    {
        private readonly SeDraft _draft;
        private string _dxfPath;
        private string _pdfPath;

        public SaveAsDxfAndPdfProcessor(SeDraft draft)
        {
            _draft = draft;
        }

        public bool Initialize()
        {
            if (string.IsNullOrEmpty(_draft.FullName))
            {
                MessageBox.Show("Save the drawing first to export DXF and PDF formats.", "Save Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string dir = Path.GetDirectoryName(_draft.FullName);
            string file = Path.GetFileNameWithoutExtension(_draft.FullName);

            _dxfPath = Path.Combine(dir, file + ".dxf");
            _pdfPath = Path.Combine(dir, file + ".pdf");

            return true;
        }

        public void Process()
        {
            _draft.SaveAs(_dxfPath, 14);
            _draft.SaveAs(_pdfPath, 5);
        }
    }
}