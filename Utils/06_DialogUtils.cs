namespace SolidEdgeAdd_In.Utils
{
    public class DialogUtils
    {
        public static (bool isConfirmed, string editedPath) GetEditedPath(string filePath)
        {
            bool isConfirmed = false;
            string editedPath = filePath;

            using Form prompt = new()
            {
                Width = 800,
                Height = 200,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Zapisać plik?",
                StartPosition = FormStartPosition.CenterScreen
            };

            Label label = new() { Left = 20, Top = 20, Width = 740, Text = "Wygenerowana scieżka (do edycji):" };
            TextBox textBox = new() { Left = 20, Top = 50, Width = 740, Text = filePath };

            Button yesButton = new() { Text = "Tak", Left = 550, Width = 100, Top = 100, DialogResult = DialogResult.Yes };
            Button noButton = new() { Text = "Nie", Left = 660, Width = 100, Top = 100, DialogResult = DialogResult.No };

            prompt.Controls.Add(label); prompt.Controls.Add(textBox);
            prompt.Controls.Add(yesButton); prompt.Controls.Add(noButton);
            prompt.AcceptButton = yesButton; prompt.CancelButton = noButton;

            if (prompt.ShowDialog() == DialogResult.Yes)
            {
                isConfirmed = true;
                editedPath = textBox.Text;
            }

            return (isConfirmed, editedPath);
        }

        public static (bool isConfirmed, int multiplier) GetMultiplier()
        {
            int multiplier = 1;
            bool isConfirmed = false;

            using Form prompt = new()
            {
                Width = 300,
                Height = 160,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Mnożnik złożenia",
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false
            };

            Label textLabel = new() { Left = 20, Top = 20, Text = "Podaj mnożnik:", AutoSize = true };
            TextBox textBox = new() { Left = 20, Top = 50, Width = 240 };
            Button confirmation = new() { Text = "OK", Left = 160, Width = 100, Top = 80, DialogResult = DialogResult.OK };

            prompt.Controls.Add(textLabel); prompt.Controls.Add(textBox); prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;

            if (prompt.ShowDialog() == DialogResult.OK)
            {
                if (int.TryParse(textBox.Text, out int result))
                {
                    multiplier = result;
                    isConfirmed = true;
                }
            }

            return (isConfirmed, multiplier);
        }

        public static List<string> GetSelectedTypes()
        {
            List<string> selectedTypes = new();
            List<string> availableOptions = new()
            {
                $"Złożenia ({Constants.PartTypes.Assembly})",
                $"Części ({Constants.PartTypes.Part})",
                $"Blachy ({Constants.PartTypes.SheetMetal})",
                $"Handlowe ({Constants.PartTypes.Commercial})",
                $"Hutnicze ({Constants.PartTypes.Steelmaking})",
                $"Normalia ({Constants.PartTypes.Standard})"
            };

            using Form prompt = new()
            {
                Width = 350,
                Height = 320,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Wybór typu",
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false
            };

            Label label = new() { Left = 20, Top = 20, Width = 300, Text = "Wybierz typ:", AutoSize = true };
            CheckedListBox checkedListBox = new() { Left = 20, Top = 50, Width = 290, Height = 160, CheckOnClick = true };

            foreach (var option in availableOptions) { checkedListBox.Items.Add(option); }

            Button okButton = new() { Text = "OK", Left = 100, Width = 100, Top = 230, DialogResult = DialogResult.OK };
            Button cancelButton = new() { Text = "Anuluj", Left = 210, Width = 100, Top = 230, DialogResult = DialogResult.Cancel };

            prompt.Controls.Add(label); prompt.Controls.Add(checkedListBox);
            prompt.Controls.Add(okButton); prompt.Controls.Add(cancelButton);
            prompt.AcceptButton = okButton; prompt.CancelButton = cancelButton;

            if (prompt.ShowDialog() == DialogResult.OK)
            {
                foreach (var item in checkedListBox.CheckedItems)
                {
                    if (item != null) { selectedTypes.Add(item.ToString()); }
                }
            }

            return selectedTypes;
        }

        public static bool IsGenerateThumbnails()
        {
            DialogResult result = MessageBox.Show(
                "Chcesz wygenerować miniatury?",
                "Miniatury",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            return result == DialogResult.Yes;
        }

        public static string GetPartsListType(List<string> savedSettings)
        {
            string selected = null;

            using var prompt = new Form
            {
                Width = 300,
                Height = 200,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Lista części",
                StartPosition = FormStartPosition.CenterScreen
            };

            ListBox listBox = new() { Left = 10, Top = 10, Width = 200, Height = 120 };
            foreach (var s in savedSettings) { listBox.Items.Add(s); }

            Button confirmation = new() { Text = "OK", Left = 10, Width = 80, Top = 135, DialogResult = DialogResult.OK };

            prompt.Controls.Add(listBox); prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;

            if (prompt.ShowDialog() == DialogResult.OK)
            {
                if (listBox.SelectedItem != null) { selected = listBox.SelectedItem.ToString(); }
            }

            return selected;
        }
    }
}