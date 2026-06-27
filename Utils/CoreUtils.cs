using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidEdgeFramework;

namespace SolidEdgeAdd_In.Utils
{
    public class CoreUtils
    {
        public static void ManageCoordinateSystemsInPart(SolidEdgePart.PartDocument part, bool visible)
        {
            foreach (SolidEdgePart.RefPlane refPlane in part.RefPlanes) { refPlane.Visible = visible; }
            foreach (SolidEdgePart.RefAxis refAxis in part.RefAxes) { refAxis.Visible = visible; }
            foreach (SolidEdgePart.CoordinateSystem cs in part.CoordinateSystems) { cs.Visible = visible; }
        }

        public static void ManageCoordinateSystemsInSheetMetal(SolidEdgePart.SheetMetalDocument sheetMetal, bool visible)
        {
            foreach (SolidEdgePart.RefPlane refPlane in sheetMetal.RefPlanes) { refPlane.Visible = visible; }
            foreach (SolidEdgePart.RefAxis refAxis in sheetMetal.RefAxes) { refAxis.Visible = visible; }
            foreach (SolidEdgePart.CoordinateSystem cs in sheetMetal.CoordinateSystems) { cs.Visible = visible; }
        }

        public static void ManageCoordinateSystemsInAssembly(SolidEdgeAssembly.AssemblyDocument assembly, bool visible)
        {
            foreach (SolidEdgeAssembly.AsmRefPlane refPlane in assembly.AsmRefPlanes) { refPlane.Visible = visible; }
        }

        public static void ReleaseCom<T>(ref T comObject) where T : class
        {
            if (comObject != null)
            {
                try { Marshal.ReleaseComObject(comObject); }
                finally { comObject = null; }
            }
        }

        public static int GetCount(Dictionary<string, int> dict, string path)
        {
            return dict[path];
        }

        public static SolidEdgeFramework.SolidEdgeDocument GetOpenDocument(SolidEdgeFramework.Application application, string filePath)
        {
            SolidEdgeDocument document = null;
            try
            {
                application.DisplayAlerts = false;
                int seOpenNoAssemblyContext = 32;
                int seOpenNoVisible = 128;
                int openFlags = seOpenNoAssemblyContext | seOpenNoVisible;
                document = (SolidEdgeDocument)application.Documents.Open(filePath, openFlags);
            }
            finally { application.DisplayAlerts = true; }
            return document;
        }
    }

}