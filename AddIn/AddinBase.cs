namespace SolidEdgeAdd_In
{
    public abstract class SolidEdgeAddIn
    {
        public abstract SeApp Application { get; }
        public abstract SeAddIn AddIn { get; }
        public abstract SeISEAddInEx AddInEx { get; }
        public abstract Guid Guid { get; }
        public abstract string NativeResourcesDllPath { get; }
    }

    public class SolidEdgeAddInWrapper : SolidEdgeAddIn
    {
        private readonly AddIn _addIn;

        public SolidEdgeAddInWrapper(AddIn addIn)
        {
            _addIn = addIn;
        }

        public override SeApp Application => _addIn.m_application;
        public override SeAddIn AddIn => _addIn.m_addin;
        public override SeISEAddInEx AddInEx => (SeISEAddInEx)_addIn.m_addin;
        public override Guid Guid => new("384E5BCF-DD43-49AC-BF01-F99CC009B35F");
        public override string NativeResourcesDllPath => Assembly.GetExecutingAssembly().Location;
    }
}