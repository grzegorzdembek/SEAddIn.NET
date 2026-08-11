using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class SetCountPropertyProcessor
    {
        private readonly SeAssembly _assembly;

        private readonly Dictionary<string, FileData> _data;
        private int _dataCount;

        private int _multiplier;

        private readonly StringBuilder _feedback;

        public SetCountPropertyProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _data = new Dictionary<string, FileData>(StringComparer.OrdinalIgnoreCase);
            _feedback = new StringBuilder();
        }

        public bool Initialize()
        {
            if (!IsLoaded_Data()) return false;

            if (!IsLoaded_Multiplier()) return false;

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

            try { occurrences = _assembly.Occurrences; DataUtils.ApplyCounts(occurrences, _data, _multiplier, processedPaths); }
            finally { Helpers.ReleaseCom(ref occurrences); }

            foreach (var item in _data)
            {
                try
                {
                    string type = item.Value.Type;
                    string name = item.Value.Name;

                    int occurrenceCount = item.Value.OccurrenceCount;
                    int count = item.Value.Count;
                    int targetCount = _multiplier * count;

                    if (!processedPaths.Contains(item.Key))
                    {
                        missFiles++;
                        _feedback.AppendLine($"{name,-30} | {type} | --- | --- | --- | --- |");
                        continue;
                    }

                    string P = count.ToString("D3"); 
                    string C = occurrenceCount.ToString("D3"); 
                    string MC = targetCount.ToString("D3"); 

                    _feedback.AppendLine($"{name,-30} | {type} | {P} | {M} | {C} | {MC} |");
                }
                catch
                {
                    missFiles++;
                    string errorFileName = "Unknown file";
                    try { errorFileName = item.Value.Name ?? Path.GetFileNameWithoutExtension(item.Key); } catch { }
                    _feedback.AppendLine($"{errorFileName,-30} | --- | --- | --- | --- | --- | --- |");
                    continue;
                }
            }

            if (missFiles == 0) { _feedback.AppendLine($"Successfully added property - Count for all files."); }
            else { _feedback.AppendLine($"Skipped {missFiles} files."); }

            DisplayFeedback();
        }

        private bool IsLoaded_Data()
        {
            SeOccurrences occurrences = null;
            try
            {
                occurrences = _assembly.Occurrences;
                DataUtils.BuildDataForSetCountProperty(occurrences, _data);
            }
            finally { Helpers.ReleaseCom(ref occurrences); }

            _dataCount = _data.Count;
            if (!Helpers.IsMessageAccepted($"Files {_dataCount}.")) return false;

            return true;
        }

        private bool IsLoaded_Multiplier()
        {
            SeDocument document = (SeDocument)_assembly;
            using PropertyUtils properties = new(document);
            int count = properties.Count;

            if (count == 0)
            {
                (bool isConfirmed, int multiplier) = DialogUtils.GetMultiplier();
                if (isConfirmed)
                {
                    properties.Count = multiplier;
                    _multiplier = multiplier;
                    return true;
                }
                return false;
            }

            _multiplier = count;
            if (!Helpers.IsMessageAccepted($"Multiplier {_multiplier}.")) return false;

            return true;
        }

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