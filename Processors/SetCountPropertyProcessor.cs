using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Processors
{
    public class SetCountPropertyProcessor
    {
        private readonly SeAssembly _assembly;
        private int _multiplier;
        private Dictionary<string, FileData> _data;
        private StringBuilder _feedback;

        public SetCountPropertyProcessor(SeAssembly assembly)
        {
            _assembly = assembly;
            _data = new Dictionary<string, FileData>(StringComparer.OrdinalIgnoreCase);
            _feedback = new StringBuilder();
        }

        public bool Initialize()
        {
            SeDocument document = (SeDocument)_assembly;
            using var properties = new PropertyProvider(document);

            int count = properties.Count;

            if (count == 0)
            {
                var result = DialogService.GetMultiplier();
                if (result.isConfirmed)
                {
                    properties.Count = result.multiplier;
                    _multiplier = result.multiplier;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                _multiplier = count;
            }

            SeOccurrences occurrences = null;
            try
            {
                occurrences = _assembly.Occurrences;
                AssemblyTreeWalker.BuildDataForSetCount(occurrences, _data);
            }
            finally
            {
                CoreUtils.ReleaseCom(ref occurrences);
            }

            return true;
        }

        public void Process()
        {
            _feedback.AppendLine($"SELECTED MULTIPLIER:{_multiplier}");
            _feedback.AppendLine($"{"FILE NAME",-30} | T |  M  |  I  |  C  | M*C |");

            string M = _multiplier.ToString("D3");
            int missFiles = 0;
            int missTypes = 0;

            HashSet<string> processed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            SeOccurrences occurrences = null;

            try
            {
                occurrences = _assembly.Occurrences;
                AssemblyTreeWalker.ApplyCounts(occurrences, _data, _multiplier, processed);
            }
            finally
            {
                CoreUtils.ReleaseCom(ref occurrences);
            }

            foreach (var item in _data)
            {
                try
                {
                    string type = item.Value.Type;

                    if (type == null)
                    {
                        missTypes++;
                        continue;
                    }

                    if (type == Constants.PartTypes.Assembly || type == Constants.PartTypes.SheetMetal || type == Constants.PartTypes.Part || type == Constants.PartTypes.Steelmaking)
                    {
                        string name = item.Value.Name;
                        int count = item.Value.OccurrenceCount;
                        int oldCount = item.Value.Count;
                        int targetCount = _multiplier * count;

                        if (!processed.Contains(item.Key))
                        {
                            missFiles++;
                            _feedback.AppendLine($"{name,-30} | {type} | --- | --- | --- | --- |");
                            continue;
                        }

                        string I = oldCount.ToString("D3");
                        string C = count.ToString("D3");
                        string MC = targetCount.ToString("D3");

                        _feedback.AppendLine($"{name,-30} | {type} | {M} | {I} | {C} | {MC} |");
                    }
                }
                catch
                {
                    missFiles++;
                    string errorName = "Unknown file";
                    try
                    {
                        errorName = item.Value.Name ?? Path.GetFileNameWithoutExtension(item.Key);
                    }
                    catch
                    {
                    }
                    _feedback.AppendLine($"{errorName,-30} | --- | --- | --- | --- | --- | --- |");
                    continue;
                }
            }

            _feedback.AppendLine($"Number of files without Type property: {missTypes}");

            if (missFiles == 0)
            {
                _feedback.AppendLine($"Successfully added property - Quantity for all files.");
            }
            else
            {
                _feedback.AppendLine($"Skipped {missFiles} files.");
            }

            DisplayFeedback();
        }

        private void DisplayFeedback()
        {
            using Form form = new Form
            {
                Text = "Feedback",
                Width = 700,
                Height = 700,
                StartPosition = FormStartPosition.CenterScreen
            };

            TextBox textBox = new TextBox
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