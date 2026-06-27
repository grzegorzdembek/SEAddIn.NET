using System;
using System.Reflection;

namespace SolidEdgeAdd_In
{
    public abstract class SolidEdgeAddIn
    {
        public abstract SolidEdgeFramework.Application Application { get; }
        public abstract SolidEdgeFramework.AddIn AddIn { get; }
        public abstract SolidEdgeFramework.ISEAddInEx AddInEx { get; }
        public abstract Guid Guid { get; }
        public abstract string NativeResourcesDllPath { get; }
    }

    public class SolidEdgeAddInWrapper : SolidEdgeAddIn
    {
        private readonly AddIn _addIn;
        public SolidEdgeAddInWrapper(AddIn addIn) { _addIn = addIn; }
        public override SolidEdgeFramework.Application Application => _addIn.m_application;
        public override SolidEdgeFramework.AddIn AddIn => _addIn.m_addin;
        public override SolidEdgeFramework.ISEAddInEx AddInEx => (SolidEdgeFramework.ISEAddInEx)_addIn.m_addin;
        public override Guid Guid => new Guid("6CD67BC6-6FD8-4922-AEC7-1731EA4F629E");
        public override string NativeResourcesDllPath => Assembly.GetExecutingAssembly().Location;
    }
}