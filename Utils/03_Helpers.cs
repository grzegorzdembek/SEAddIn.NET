namespace SolidEdgeAdd_In.Utils
{
    public class Logger
    {
        private readonly StringBuilder _logBuilder = new ();
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
    }
}