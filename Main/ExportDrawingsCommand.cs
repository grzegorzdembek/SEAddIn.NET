using SolidEdgeAdd_In.Helpers;

namespace SolidEdgeAdd_In.Main
{
    public class ExportDrawingsCommand
    {
        public static void Execute(SeDocument document)
        {
            try
            {
                /*
                 *  folder projektu -> 
                 *  folder Paczki -> 
                 *  folder wybrany przez uzytkownika (tu musi znajdowac sie zestawienie) -> 
                 *  tworzenie folderu Rysunki (tu kopiujemy rysunki na podstawie zestawienia)
                 */
                string defaultPath = ExportDrawingsHelper.GetDefaultPath(document);
                if (string.IsNullOrEmpty(defaultPath)) { MessageBox.Show("Nie znaleziono folderu Paczki", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                string selectedPath = ExportDrawingsHelper.GetSelectedFolder(defaultPath);
                if (string.IsNullOrEmpty(selectedPath)) return;

                string excelListPath = ExportDrawingsHelper.GetSummaryExcelPath(selectedPath);
                if (string.IsNullOrEmpty(excelListPath)) { MessageBox.Show("Nie znaleziono JEDNEGO pliku z rozszerzeniem .xlsx", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                /*
                 * 
                 */
                ExportDrawingsHelper.ProcessDrawings(defaultPath, selectedPath, excelListPath);
                MessageBox.Show("Zakończono działanie makra", "Sukces", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Błąd podczas dodawania rysunków: {ex.Message}", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
