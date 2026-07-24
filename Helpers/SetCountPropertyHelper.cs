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
            Dictionary<string, FileData> data = new(StringComparer.OrdinalIgnoreCase); AssemblyTreeWalker.BuildDataForSetCount(assembly.Occurrences, data); return data;
        }

        public static StringBuilder SetAndGetFeedback(SeAssembly assembly, Dictionary<string, FileData> data, int multiplier)
        {
            SeDocument document = null; StringBuilder feedback = new(); feedback.AppendLine($"WYBRANO MNOŻNIK:{multiplier}"); feedback.AppendLine("NAZWA PLIKU       | T |  M  |  I  |  C  | M*C |");

            string M = multiplier.ToString("D3"); int missFiles = 0; int missTypes = 0;

            foreach (var item in data)
            {
                try
                {
                    string type = item.Value.Type; if (type == null) { missTypes++; continue; }

                    if (type == "A" || type == "B" || type == "C" || type == "K")
                    {
                        string name = item.Value.Name; int count = item.Value.OccurrenceCount; int oldCount = item.Value.Count; int targetCount = multiplier * count; bool saveSuccess = true;

                        try { document = CoreUtils.GetOpenDocument(assembly.Application, item.Key); if (document != null) { using var properties = new PropertyProvider(document); properties.Count = targetCount; document.Save(); } else { saveSuccess = false; } }
                        catch { saveSuccess = false; }
                        finally { document?.Close(false); CoreUtils.ReleaseCom(ref document); }

                        if (!saveSuccess) { missFiles++; feedback.AppendLine($"{name}       | {type} | --- | --- | --- | --- |"); continue; }

                        string I = oldCount.ToString("D3"); string C = count.ToString("D3"); string MC = targetCount.ToString("D3");

                        feedback.AppendLine($"{name}       | {type} | {M} | {I} | {C} | {MC} |");
                    }
                }
                catch { missFiles++; string errorName = "Nieznany plik"; try { errorName = item.Value.Name ?? Path.GetFileNameWithoutExtension(item.Key); } catch { } feedback.AppendLine($"{errorName}       | --- | --- | --- | --- | --- | --- |"); continue; }
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