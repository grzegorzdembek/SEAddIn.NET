namespace SolidEdgeAdd_In.Utils
{
    public class Logger
    {
        private readonly StringBuilder _logBuilder = new();
        private int _successCount = 0;
        private int _skipCount = 0;
        private int _errorCount = 0;

        public void LogSkip(string fileName, string reason)
        {
            _logBuilder.AppendLine($"[SKIPPED] {fileName} -> {reason}");
            _skipCount++;
        }

        public void LogSuccess(string fileName)
        {
            _logBuilder.AppendLine($"[SUCCESS] {fileName}");
            _successCount++;
        }

        public void LogError(string fileName, string errorMessage)
        {
            _logBuilder.AppendLine($"[ERROR]   {fileName} -> {errorMessage}");
            _errorCount++;
        }

        public void SaveReport(string directoryPath)
        {
            string reportPath = Path.Combine(directoryPath, "Report.txt");
            var finalReport = new StringBuilder();

            finalReport.AppendLine("==================================================");
            finalReport.AppendLine("                     REPORT                       ");
            finalReport.AppendLine("==================================================");
            finalReport.AppendLine($"Date: {DateTime.Now}");
            finalReport.AppendLine($"Processed files:    {_successCount}");
            finalReport.AppendLine($"Skipped files:      {_skipCount}");
            finalReport.AppendLine($"Errors encountered: {_errorCount}");
            finalReport.AppendLine("==================================================\n");

            finalReport.Append(_logBuilder.ToString());
            File.WriteAllText(reportPath, finalReport.ToString());
        }
    }

    public class Helpers
    {
        public static void ReleaseCom<T>(ref T comObject) where T : class
        {
            if (comObject != null)
            {
                try { Marshal.ReleaseComObject(comObject); }
                finally { comObject = null; }
            }
        }

        public static bool IsMessageAccepted(string message)
        {
            DialogResult isAccepted = MessageBox.Show(message, "Kontynuować?", MessageBoxButtons.OKCancel);
            return isAccepted == DialogResult.OK;
        }

        public static SeDocument GetOpenDocument(SeApp application, string filePath)
        {
            SeDocument document = null;
            SeDocuments documents = null;

            try
            {
                application.DisplayAlerts = false;
                int seOpenNoAssemblyContext = 32;
                int seOpenNoVisible = 128;
                int openFlags = seOpenNoAssemblyContext | seOpenNoVisible;

                documents = application.Documents;
                document = (SeDocument)documents.Open(filePath, openFlags);
            }
            finally
            {
                application.DisplayAlerts = true;
                ReleaseCom(ref documents);
            }

            return document;
        }

        private static void SetVisibilitySafe(object comObj, bool visible)
        {
            if (comObj == null) return;
            try
            {
                dynamic d = comObj;
                d.Visible = visible;
            }
            catch
            {
                try
                {
                    dynamic d = comObj;
                    d.Hidden = !visible;
                }
                catch
                {
                    try
                    {
                        dynamic d = comObj;
                        if (visible) d.Show();
                        else d.Hide();
                    }
                    catch {  }
                }
            }
        }

        public static void ManageNonModelElementsInPart(SePart part, bool visible)
        {
            SeRefPlanes planes = null;
            SeRefAxes axes = null;
            SeCoordinateSystems coords = null;
            Sketchs sketches = null;
            SePMI pmi = null;
            SeDimensions pmiDims = null;

            try
            {
                planes = part.RefPlanes;
                if (planes != null)
                {
                    for (int i = 1; i <= planes.Count; i++)
                    {
                        SeRefPlane plane = null;
                        try { plane = (SeRefPlane)planes.Item(i); SetVisibilitySafe(plane, visible); }
                        finally { ReleaseCom(ref plane); }
                    }
                }

                axes = part.RefAxes;
                if (axes != null)
                {
                    for (int i = 1; i <= axes.Count; i++)
                    {
                        SeRefAxis axis = null;
                        try { axis = (SeRefAxis)axes.Item(i); SetVisibilitySafe(axis, visible); }
                        finally { ReleaseCom(ref axis); }
                    }
                }

                coords = part.CoordinateSystems;
                if (coords != null)
                {
                    for (int i = 1; i <= coords.Count; i++)
                    {
                        SeCoordinateSystem cs = null;
                        try { cs = (SeCoordinateSystem)coords.Item(i); SetVisibilitySafe(cs, visible); }
                        finally { ReleaseCom(ref cs); }
                    }
                }

                try { sketches = (Sketchs)((dynamic)part).Sketches; } catch { }
                if (sketches != null)
                {
                    for (int i = 1; i <= sketches.Count; i++)
                    {
                        Sketch sketch = null;
                        try { sketch = (Sketch)sketches.Item(i); SetVisibilitySafe(sketch, visible); }
                        finally { ReleaseCom(ref sketch); }
                    }
                }

                pmi = (SePMI)part.PMI;
                if (pmi != null)
                {
                    pmiDims = (SeDimensions)pmi.Dimensions;
                    if (pmiDims != null)
                    {
                        for (int i = 1; i <= pmiDims.Count; i++)
                        {
                            SeDimension dim = null;
                            try { dim = (SeDimension)pmiDims.Item(i); SetVisibilitySafe(dim, visible); }
                            finally { ReleaseCom(ref dim); }
                        }
                    }
                }
            }
            finally
            {
                ReleaseCom(ref pmiDims);
                ReleaseCom(ref pmi);
                ReleaseCom(ref sketches);
                ReleaseCom(ref coords);
                ReleaseCom(ref axes);
                ReleaseCom(ref planes);
            }
        }

        public static void ManageNonModelElementsInSheetMetal(SeSheetMetal sheetMetal, bool visible)
        {
            SeRefPlanes planes = null; 
            SeRefAxes axes = null;
            SeCoordinateSystems coords = null;
            Sketchs sketches = null;
            SePMI pmi = null;
            SeDimensions pmiDims = null;

            try
            {
                planes = sheetMetal.RefPlanes;
                if (planes != null)
                {
                    for (int i = 1; i <= planes.Count; i++)
                    {
                        SeRefPlane plane = null;
                        try { plane = (SeRefPlane)planes.Item(i); SetVisibilitySafe(plane, visible); }
                        finally { ReleaseCom(ref plane); }
                    }
                }

                axes = sheetMetal.RefAxes;
                if (axes != null)
                {
                    for (int i = 1; i <= axes.Count; i++)
                    {
                        SeRefAxis axis = null;
                        try { axis = (SeRefAxis)axes.Item(i); SetVisibilitySafe(axis, visible); }
                        finally { ReleaseCom(ref axis); }
                    }
                }

                coords = sheetMetal.CoordinateSystems;
                if (coords != null)
                {
                    for (int i = 1; i <= coords.Count; i++)
                    {
                        SeCoordinateSystem cs = null;
                        try { cs = (SeCoordinateSystem)coords.Item(i); SetVisibilitySafe(cs, visible); }
                        finally { ReleaseCom(ref cs); }
                    }
                }

                try { sketches = (Sketchs)((dynamic)sheetMetal).Sketches; } catch { }
                if (sketches != null)
                {
                    for (int i = 1; i <= sketches.Count; i++)
                    {
                        Sketch sketch = null;
                        try { sketch = (Sketch)sketches.Item(i); SetVisibilitySafe(sketch, visible); }
                        finally { ReleaseCom(ref sketch); }
                    }
                }

                pmi = (SePMI)sheetMetal.PMI;
                if (pmi != null)
                {
                    pmiDims = (SeDimensions)pmi.Dimensions;
                    if (pmiDims != null)
                    {
                        for (int i = 1; i <= pmiDims.Count; i++)
                        {
                            SeDimension dim = null;
                            try { dim = (SeDimension)pmiDims.Item(i); SetVisibilitySafe(dim, visible); }
                            finally { ReleaseCom(ref dim); }
                        }
                    }
                }
            }
            finally
            {
                ReleaseCom(ref pmiDims);
                ReleaseCom(ref pmi);
                ReleaseCom(ref sketches);
                ReleaseCom(ref coords);
                ReleaseCom(ref axes);
                ReleaseCom(ref planes);
            }
        }

        public static void ManageNonModelElementsInAssembly(SeAssembly assembly, bool visible)
        {
            SeAsmRefPlanes planes = null;
            SeAsmCoordinateSystems coords = null;
            Sketchs sketches = null;
            SePMI pmi = null;
            SeDimensions pmiDims = null;

            try
            {
                planes = assembly.AsmRefPlanes;
                if (planes != null)
                {
                    for (int i = 1; i <= planes.Count; i++)
                    {
                        SeAsmRefPlane plane = null;
                        try { plane = (SeAsmRefPlane)planes.Item(i); SetVisibilitySafe(plane, visible); }
                        finally { ReleaseCom(ref plane); }
                    }
                }

                coords = assembly.CoordinateSystems;
                if (coords != null)
                {
                    for (int i = 1; i <= coords.Count; i++)
                    {
                        SeAsmCoordinateSystem cs = null;
                        try { cs = (SeAsmCoordinateSystem)coords.Item(i); SetVisibilitySafe(cs, visible); }
                        finally { ReleaseCom(ref cs); }
                    }
                }

               
                try { sketches = (Sketchs)((dynamic)assembly).Sketches; } catch { }
                if (sketches == null) { try { sketches = (Sketchs)((dynamic)assembly).AsmSketches; } catch { } }

                if (sketches != null)
                {
                    for (int i = 1; i <= sketches.Count; i++)
                    {
                        Sketch sketch = null;
                        try { sketch = (Sketch)sketches.Item(i); SetVisibilitySafe(sketch, visible); }
                        finally { ReleaseCom(ref sketch); }
                    }
                }

                pmi = (SePMI)assembly.PMI;
                if (pmi != null)
                {
                    pmiDims = (SeDimensions)pmi.Dimensions;
                    if (pmiDims != null)
                    {
                        for (int i = 1; i <= pmiDims.Count; i++)
                        {
                            SeDimension dim = null;
                            try { dim = (SeDimension)pmiDims.Item(i); SetVisibilitySafe(dim, visible); }
                            finally { ReleaseCom(ref dim); }
                        }
                    }
                }
            }
            finally
            {
                ReleaseCom(ref pmiDims);
                ReleaseCom(ref pmi);
                ReleaseCom(ref sketches);
                ReleaseCom(ref coords);
                ReleaseCom(ref planes);
            }
        }
    }
}