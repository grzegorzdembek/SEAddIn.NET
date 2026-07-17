namespace SolidEdgeAdd_In.Utils
{
    public class DxfExportLogger
    {
        private readonly StringBuilder _logBuilder = new ();
        private int _successCount = 0;
        private int _skipCount = 0;
        private int _errorCount = 0;

        public void LogSkip(string fileName, string reason)
        {
            _logBuilder.AppendLine($"[POMINIĘTO] {fileName} -> {reason}");
            _skipCount++;
        }

        public void LogSuccess(string fileName)
        {
            _logBuilder.AppendLine($"[SUKCES]    {fileName} -> Wyeksportowano pomyślnie");
            _successCount++;
        }

        public void LogError(string fileName, string errorMessage)
        {
            _logBuilder.AppendLine($"[BŁĄD]      {fileName} -> {errorMessage}");
            _errorCount++;
        }

        public void SaveReport(string directoryPath)
        {
            try
            {
                string reportPath = Path.Combine(directoryPath, "Raport_Eksportu_DXF.txt");
                var finalReport = new StringBuilder();

                finalReport.AppendLine("==================================================");
                finalReport.AppendLine("             RAPORT Z EKSPORTU DXF                ");
                finalReport.AppendLine("==================================================");
                finalReport.AppendLine($"Data wygenerowania: {DateTime.Now}");
                finalReport.AppendLine($"Liczba plików wyeksportowanych: {_successCount}");
                finalReport.AppendLine($"Liczba plików pominiętych:      {_skipCount}");
                finalReport.AppendLine($"Liczba błędów krytycznych:      {_errorCount}");
                finalReport.AppendLine("==================================================\n");

                finalReport.Append(_logBuilder.ToString());

                File.WriteAllText(reportPath, finalReport.ToString());
            }
            catch { }
        }
    }


    public class CoreUtils
    {

        public static void ManageCoordinateSystemsInPart(SePart part, bool visible)
        {
            foreach (SeRefPlane refPlane in part.RefPlanes) { refPlane.Visible = visible; }
            foreach (SeRefAxis refAxis in part.RefAxes) { refAxis.Visible = visible; }
            foreach (SeCoordinateSystem cs in part.CoordinateSystems) { cs.Visible = visible; }
        }

        public static void ManageCoordinateSystemsInSheetMetal(SeSheetMetal sheetMetal, bool visible)
        {
            foreach (SeRefPlane refPlane in sheetMetal.RefPlanes) { refPlane.Visible = visible; }
            foreach (SeRefAxis refAxis in sheetMetal.RefAxes) { refAxis.Visible = visible; }
            foreach (SeCoordinateSystem cs in sheetMetal.CoordinateSystems) { cs.Visible = visible; }
        }

        public static void ManageCoordinateSystemsInAssembly(SeAssembly assembly, bool visible)
        {
            foreach (SeAsmRefPlane refPlane in assembly.AsmRefPlanes) { refPlane.Visible = visible; }
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

        public static SeDocument GetOpenDocument(SeApp application, string filePath)
        {
            SeDocument document = null;
            try
            {
                application.DisplayAlerts = false;

                int seOpenNoAssemblyContext = 32;
                int seOpenNoVisible = 128;
                int openFlags = seOpenNoAssemblyContext | seOpenNoVisible;

                document = (SeDocument)application.Documents.Open(filePath, openFlags);
            }
            finally
            {
                application.DisplayAlerts = true;
            }
            return document;
        }
    }
}