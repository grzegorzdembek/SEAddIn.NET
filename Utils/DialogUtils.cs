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
                Text = "Execute save?",
                StartPosition = FormStartPosition.CenterScreen
            };

            Label label = new() { Left = 20, Top = 20, Width = 740, Text = "Generated path (editable):" };
            TextBox textBox = new() { Left = 20, Top = 50, Width = 740, Text = filePath };

            Button yesButton = new() { Text = "Yes", Left = 550, Width = 100, Top = 100, DialogResult = DialogResult.Yes };
            Button noButton = new() { Text = "No", Left = 660, Width = 100, Top = 100, DialogResult = DialogResult.No };

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
                Text = "Data Input",
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false
            };

            Label textLabel = new() { Left = 20, Top = 20, Text = "ENTER MULTIPLIER:", AutoSize = true };
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
                Text = "Select Options",
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false
            };

            Label label = new() { Left = 20, Top = 20, Width = 300, Text = "Select Occurrences you need:", AutoSize = true };
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

        public static bool IsShotsNeeded()
        {
            DialogResult result = MessageBox.Show(
                "Generate thumbnails?",
                "Thumbnails",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            return result == DialogResult.Yes;
        }

        private static string PartsListHelper(List<string> savedSettings)
        {
            string selected = null;

            using var prompt = new Form
            {
                Width = 300,
                Height = 200,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Parts list table properties",
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

        public static string GetPartsListType(SeApp application, SeAssembly assembly)
        {
            SeDocuments documents = null;
            SeDraft draft = null;
            SeDraftSheet sheet = null;
            SeDrawingViews drawingViews = null;
            SeDrawingView drawingView = null;
            SeModelLinks modelLinks = null;
            SeModelLink modelLink = null;
            SePartsLists partsLists = null;
            SePartsList partsList = null;

            try
            {
                documents = application.Documents;
                draft = (SeDraft)documents.Add("SolidEdge.DraftDocument", Missing.Value);
                sheet = draft.ActiveSheet;
                modelLinks = draft.ModelLinks;
                modelLink = modelLinks.Add(assembly.FullName);
                drawingViews = sheet.DrawingViews;

                drawingView = drawingViews.AddAssemblyView(modelLink, SeViewOrientation.igFrontView, 0.1, 0.2, 0.2, SeAssemblyDrawingViewType.seAssemblyDesignedView);

                partsLists = draft.PartsLists;
                partsList = partsLists.AddEx(drawingView, 0, "", 0, 1);

                Array listOfSavedSettings = Array.CreateInstance(typeof(object), 0);
                partsList.GetListOfSavedSettings(out int numSavedSettings, ref listOfSavedSettings);

                var settingsList = new List<string>();

                if (listOfSavedSettings != null)
                {
                    foreach (var o in listOfSavedSettings)
                    {
                        if (o != null) { settingsList.Add(o.ToString()); }
                    }
                }

                if (settingsList.Count == 0) { settingsList.Add("<No saved parts list styles available>"); }

                return PartsListHelper(settingsList);
            }
            finally
            {
                Helpers.ReleaseCom(ref partsList); Helpers.ReleaseCom(ref partsLists);
                Helpers.ReleaseCom(ref modelLink); Helpers.ReleaseCom(ref modelLinks);
                Helpers.ReleaseCom(ref drawingView); Helpers.ReleaseCom(ref drawingViews);
                Helpers.ReleaseCom(ref sheet);

                if (draft != null)
                {
                    try { draft.Close(false); } catch { }
                }

                Helpers.ReleaseCom(ref draft); Helpers.ReleaseCom(ref documents);
            }
        }
    }
}