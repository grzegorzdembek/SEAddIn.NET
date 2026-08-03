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
            try
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
            catch
            {
            }
        }
    }

    public class CoreUtils
    {
        public static void ManageCoordinateSystemsInPart(SePart part, bool visible)
        {
            SolidEdgePart.RefPlanes planes = null;
            SolidEdgePart.RefAxes axes = null;
            SolidEdgePart.CoordinateSystems coords = null;

            try
            {
                planes = part.RefPlanes;

                if (planes != null)
                {
                    for (int i = 1; i <= planes.Count; i++)
                    {
                        SeRefPlane plane = null;

                        try
                        {
                            plane = (SeRefPlane)planes.Item(i);
                            plane.Visible = visible;
                        }
                        finally
                        {
                            ReleaseCom(ref plane);
                        }
                    }
                }

                axes = part.RefAxes;

                if (axes != null)
                {
                    for (int i = 1; i <= axes.Count; i++)
                    {
                        SeRefAxis axis = null;

                        try
                        {
                            axis = (SeRefAxis)axes.Item(i);
                            axis.Visible = visible;
                        }
                        finally
                        {
                            ReleaseCom(ref axis);
                        }
                    }
                }

                coords = part.CoordinateSystems;

                if (coords != null)
                {
                    for (int i = 1; i <= coords.Count; i++)
                    {
                        SeCoordinateSystem cs = null;

                        try
                        {
                            cs = (SeCoordinateSystem)coords.Item(i);
                            cs.Visible = visible;
                        }
                        finally
                        {
                            ReleaseCom(ref cs);
                        }
                    }
                }
            }
            finally
            {
                ReleaseCom(ref coords);
                ReleaseCom(ref axes);
                ReleaseCom(ref planes);
            }
        }

        public static void ManageCoordinateSystemsInSheetMetal(SeSheetMetal sheetMetal, bool visible)
        {
            SolidEdgePart.RefPlanes planes = null;
            SolidEdgePart.RefAxes axes = null;
            SolidEdgePart.CoordinateSystems coords = null;

            try
            {
                planes = sheetMetal.RefPlanes;

                if (planes != null)
                {
                    for (int i = 1; i <= planes.Count; i++)
                    {
                        SeRefPlane plane = null;

                        try
                        {
                            plane = (SeRefPlane)planes.Item(i);
                            plane.Visible = visible;
                        }
                        finally
                        {
                            ReleaseCom(ref plane);
                        }
                    }
                }

                axes = sheetMetal.RefAxes;

                if (axes != null)
                {
                    for (int i = 1; i <= axes.Count; i++)
                    {
                        SeRefAxis axis = null;

                        try
                        {
                            axis = (SeRefAxis)axes.Item(i);
                            axis.Visible = visible;
                        }
                        finally
                        {
                            ReleaseCom(ref axis);
                        }
                    }
                }

                coords = sheetMetal.CoordinateSystems;

                if (coords != null)
                {
                    for (int i = 1; i <= coords.Count; i++)
                    {
                        SeCoordinateSystem cs = null;

                        try
                        {
                            cs = (SeCoordinateSystem)coords.Item(i);
                            cs.Visible = visible;
                        }
                        finally
                        {
                            ReleaseCom(ref cs);
                        }
                    }
                }
            }
            finally
            {
                ReleaseCom(ref coords);
                ReleaseCom(ref axes);
                ReleaseCom(ref planes);
            }
        }

        public static void ManageCoordinateSystemsInAssembly(SeAssembly assembly, bool visible)
        {
            SolidEdgeAssembly.AsmRefPlanes planes = null;

            try
            {
                planes = assembly.AsmRefPlanes;

                if (planes != null)
                {
                    for (int i = 1; i <= planes.Count; i++)
                    {
                        SeAsmRefPlane plane = null;

                        try
                        {
                            plane = (SeAsmRefPlane)planes.Item(i);
                            plane.Visible = visible;
                        }
                        finally
                        {
                            ReleaseCom(ref plane);
                        }
                    }
                }
            }
            finally
            {
                ReleaseCom(ref planes);
            }
        }

        public static void ReleaseCom<T>(ref T comObject) where T : class
        {
            if (comObject != null)
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                finally
                {
                    comObject = null;
                }
            }
        }

        public static int GetCount(Dictionary<string, int> dict, string path)
        {
            return dict[path];
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
                CoreUtils.ReleaseCom(ref documents);
            }

            return document;
        }
    }
}