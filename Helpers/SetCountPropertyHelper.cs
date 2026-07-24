using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Helpers
{
    public class SetCountPropertyHelper
    {
        public static (bool isConfirmed, int multiplier) GetMultiplier(SeAssembly assembly)
        {
            SeDocument document = (SeDocument)assembly; using var properties = new PropertyProvider(document);

            int count = properties.Count; if (count == 0) { var result = DialogService.GetMultiplier(); if (result.isConfirmed) { properties.Count = result.multiplier; return result; } return (false, 1); }

            return (true, count);
        }

        public static Dictionary<string, FileData> GetData(SeAssembly assembly)
        {
            Dictionary<string, FileData> data = new(StringComparer.OrdinalIgnoreCase); SeOccurrences occurrences = null;

            try { occurrences = assembly.Occurrences; AssemblyTreeWalker.BuildDataForSetCount(occurrences, data); } finally { CoreUtils.ReleaseCom(ref occurrences); }

            return data;
        }

        public static StringBuilder SetAndGetFeedback(SeAssembly assembly, Dictionary<string, FileData> data, int multiplier)
        {
            StringBuilder feedback = new(); feedback.AppendLine($"WYBRANO MNOŻNIK:{multiplier}"); feedback.AppendLine($"{"NAZWA PLIKU",-30} | T |  M  |  I  |  C  | M*C |");

            string M = multiplier.ToString("D3"); int missFiles = 0; int missTypes = 0;

            HashSet<string> processed = new(StringComparer.OrdinalIgnoreCase); SeOccurrences occurrences = null;

            try { occurrences = assembly.Occurrences; AssemblyTreeWalker.ApplyCounts(occurrences, data, multiplier, processed); } finally { CoreUtils.ReleaseCom(ref occurrences); }

            foreach (var item in data)
            {
                try
                {
                    string type = item.Value.Type; if (type == null) { missTypes++; continue; }

                    if (type == "A" || type == "B" || type == "C" || type == "K")
                    {
                        string name = item.Value.Name; int count = item.Value.OccurrenceCount; int oldCount = item.Value.Count; int targetCount = multiplier * count;

                        if (!processed.Contains(item.Key)) { missFiles++; feedback.AppendLine($"{name,-30} | {type} | --- | --- | --- | --- |"); continue; }

                        string I = oldCount.ToString("D3"); string C = count.ToString("D3"); string MC = targetCount.ToString("D3");

                        feedback.AppendLine($"{name,-30} | {type} | {M} | {I} | {C} | {MC} |");
                    }
                }
                catch { missFiles++; string errorName = "Nieznany plik"; try { errorName = item.Value.Name ?? Path.GetFileNameWithoutExtension(item.Key); } catch { } feedback.AppendLine($"{errorName,-30} | --- | --- | --- | --- | --- | --- |"); continue; }
            }

            feedback.AppendLine($"Liczba plików bez właściwości Typ: {missTypes}");

            if (missFiles == 0) feedback.AppendLine($"Pomyślnie dodano właściwość - Ilość dla wszystkich plików z typem - A,B,C,K."); else feedback.AppendLine($"Pominięto {missFiles} plików.");

            return feedback;
        }

        public static void DisplayFeedback(StringBuilder feedback)
        {
            using Form form = new() { Text = "Feedback", Width = 700, Height = 700, StartPosition = FormStartPosition.CenterScreen };

            TextBox textBox = new() { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, Text = feedback.ToString(), Font = new Font("Consolas", 10) };

            form.Controls.Add(textBox); form.ShowDialog();
        }
    }
}