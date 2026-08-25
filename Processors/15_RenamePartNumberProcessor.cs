namespace SolidEdgeAdd_In.Processors
{
    public class RenamePartNumberProcessor
    {
        private readonly SeDocument _document;

        private string _documentPath;
        private string _projectDirectory;
        private string _documentName;

        public RenamePartNumberProcessor(SeDocument document)
        {
            _document = document;
        }

        public bool Initialize()
        {
            _documentPath = _document.FullName;
            _projectDirectory = Path.GetDirectoryName(_documentPath);
            _documentName = Path.GetFileNameWithoutExtension(_documentPath);



            return true;
        }

        public void Process()
        {
            
        }

      
    }
}
