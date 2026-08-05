using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class SetCountPropertyProcessor
    {
        private readonly SeAssembly _assembly;
        private readonly Dictionary<string, FileData> _data;
        private readonly StringBuilder _feedback;

        private int _multiplier;

        public SetCountPropertyProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _data = new Dictionary<string, FileData>(StringComparer.OrdinalIgnoreCase);
            _feedback = new StringBuilder();
        }

        public bool Initialize()
        {
            SeDocument document = (SeDocument)_assembly;

            using var properties = new PropertyUtils(document);
            int count = properties.Count;

            // MULTIPLIER
            if (count == 0)
            {
                var (isConfirmed, multiplier) = DialogUtils.GetMultiplier();
                if (isConfirmed) { properties.Count = multiplier; _multiplier = multiplier; }
                else { return false; }
            }
            else { _multiplier = count; }

            // SCAN DATA
            SeOccurrences occurrences = null;
            try { occurrences = _assembly.Occurrences; DataUtils.BuildDataForSetCount(occurrences, _data); }
            finally { Helpers.ReleaseCom(ref occurrences); }

            return true;
        }

        public void Process()
        {
            _feedback.AppendLine($"SELECTED MULTIPLIER:{_multiplier}");
            _feedback.AppendLine($"{"FILE NAME",-30} | T |  M  |  I  |  C  | M*C |");

            string M = _multiplier.ToString("D3");
            int missFiles = 0;

            HashSet<string> processedPaths = new(StringComparer.OrdinalIgnoreCase);
            SeOccurrences occurrences = null;

            // APPLY COUNTS
            try { occurrences = _assembly.Occurrences; DataUtils.ApplyCounts(occurrences, _data, _multiplier, processedPaths); }
            finally { Helpers.ReleaseCom(ref occurrences); }

            // REPORTING
            foreach (var item in _data)
            {
                try
                {
                    string type = item.Value.Type;
                    string fileName = item.Value.FileName;

                    int count = item.Value.OccurrenceCount;
                    int oldCount = item.Value.Count;
                    int targetCount = _multiplier * count;

                    if (!processedPaths.Contains(item.Key))
                    {
                        missFiles++;
                        _feedback.AppendLine($"{fileName,-30} | {type} | --- | --- | --- | --- |");
                        continue;
                    }

                    string I = oldCount.ToString("D3");
                    string C = count.ToString("D3");
                    string MC = targetCount.ToString("D3");

                    _feedback.AppendLine($"{fileName,-30} | {type} | {M} | {I} | {C} | {MC} |");
                }
                catch
                {
                    missFiles++;
                    string errorFileName = "Unknown file";
                    try { errorFileName = item.Value.FileName ?? Path.GetFileNameWithoutExtension(item.Key); } catch { }
                    _feedback.AppendLine($"{errorFileName,-30} | --- | --- | --- | --- | --- | --- |");
                    continue;
                }
            }

            // SUMMARY
            if (missFiles == 0) { _feedback.AppendLine($"Successfully added property - Quantity for all files."); }
            else { _feedback.AppendLine($"Skipped {missFiles} files."); }

            DisplayFeedback();
        }

        // DIALOG WINDOW
        private void DisplayFeedback()
        {
            using Form form = new()
            {
                Text = "Feedback",
                Width = 700,
                Height = 700,
                StartPosition = FormStartPosition.CenterScreen
            };

            TextBox textBox = new()
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Text = _feedback.ToString(),
                Font = new Font("Consolas", 10)
            };

            form.Controls.Add(textBox);
            form.ShowDialog();
        }
    }
}