using SolidEdgeAdd_In.Utils; 
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Helpers.AssemblyEnviroment
{
    public class SetCountPropertyHelper
    {
        public static int GetMultiplier(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            SolidEdgeFramework.SolidEdgeDocument document = (SolidEdgeFramework.SolidEdgeDocument)assembly;
            int count = PropertyProvider.GetCount(document);
            if (count == 0) PropertyProvider.SetCount(document, count = DialogService.GetMultiplier());
            return count;
        }

        public static Dictionary<string, int> GetOccurrences(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            var occurrences = new Dictionary<string, int>();
            AssemblyTreeWalker.AllOccurrences(assembly.Occurrences, occurrences);
            return occurrences;
        }

        public static StringBuilder SetAndGetFeedback(Dictionary<string, int> occurrences, int multiplier)
        {
            StringBuilder feedback = new StringBuilder();
            feedback.AppendLine($"WYBRANO MNOŻNIK:{multiplier}");
            feedback.AppendLine("NAZWA PLIKU      | T |  M  |  I  |  C  | M*C |");

            string M = multiplier.ToString("D3");
    
            int missFiles = 0;
            int missTypes = 0;

            foreach (var occurrence in occurrences)
            {
                try
                {
                    string type = PropertyProvider.GetType(occurrence.Key);
                    if (type == null) { missTypes++; continue; }

                    if (type == "A" || type == "B" || type == "C" || type == "K")                       
                    {
                        string name = System.IO.Path.GetFileNameWithoutExtension(occurrence.Key);
                        int count = occurrence.Value;

                        int oldCount = PropertyProvider.GetCount(occurrence.Key);
                        int targetCount = multiplier * count;
                        bool saveSuccess = PropertyProvider.SetCount(occurrence.Key, targetCount);

                        if (!saveSuccess)
                        {
                            missFiles++;
                            feedback.AppendLine($"{name}   | {type} | --- | --- | --- | --- |");
                            continue;
                        }

                        int newCount = targetCount;

                        string I = oldCount.ToString("D3");
                        string C = count.ToString("D3");
                        string MC = newCount.ToString("D3");

                        feedback.AppendLine($"{name}   | {type} | {M} | {I} | {C} | {MC} |");
                    }
                }
                catch
                {
                    missFiles++;
                    string errorName = "NIEZNANY PLIK";
                    try { errorName = System.IO.Path.GetFileNameWithoutExtension(occurrence.Key); } catch { }
                    feedback.AppendLine($"{errorName}   | --- | --- | --- | --- | --- | --- |");
                    continue;
                }
            }

            feedback.AppendLine($"Liczba plików bez właściwości Typ: {missTypes}");
            if (missFiles == 0) feedback.AppendLine($"Pomyślnie dodano właściwość - Ilość dla wszystkich plików z typem - A,B,C,K.");
            else feedback.AppendLine($"Pominięto {missFiles} plików.");

            return feedback;
        }

        public static void DisplayFeedback(StringBuilder feedback)
        {
            Form form = new Form
            {
                Text = "Feedback",
                Width = 700,  
                Height = 700, 
                StartPosition = FormStartPosition.CenterScreen
            };

            // 2. Tworzymy pole tekstowe, które zajmie całe okno
            TextBox textBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Text = feedback.ToString(),
                Font = new System.Drawing.Font("Consolas", 10)
            };

            form.Controls.Add(textBox);
            form.ShowDialog();
        }
    }
}