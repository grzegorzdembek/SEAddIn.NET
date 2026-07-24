namespace SolidEdgeAdd_In.Utils
{
    public class DialogService
    {
        public static bool IsFileLocationsConfirmed()
        {
            DialogResult result = MessageBox.Show("Upewnij się, że wszystkie pliki .par i .psm, które są w tym samym folderze co otwarte złożenie, mają typ B lub C!", "Wymagane potwierdzenie", MessageBoxButtons.YesNo, MessageBoxIcon.Warning); return result == DialogResult.Yes;
        }

        public static bool IsShotsNeeded()
        {
            DialogResult result = MessageBox.Show("Czy wygenerować miniatury?", "Miniatury", MessageBoxButtons.YesNo, MessageBoxIcon.Question); return result == DialogResult.Yes;
        }

        public static (bool isConfirmed, int multiplier) GetMultiplier()
        {
            int multiplier = 1; bool isConfirmed = false;

            using var prompt = new Form { Width = 300, Height = 160, FormBorderStyle = FormBorderStyle.FixedDialog, Text = "Wprowadzanie danych", StartPosition = FormStartPosition.CenterScreen, MinimizeBox = false, MaximizeBox = false };

            Label textLabel = new() { Left = 20, Top = 20, Text = "PODAJ MNOŻNIK:", AutoSize = true }; TextBox textBox = new() { Left = 20, Top = 50, Width = 240 }; Button confirmation = new() { Text = "OK", Left = 160, Width = 100, Top = 80, DialogResult = DialogResult.OK };

            prompt.Controls.Add(textLabel); prompt.Controls.Add(textBox); prompt.Controls.Add(confirmation); prompt.AcceptButton = confirmation;

            if (prompt.ShowDialog() == DialogResult.OK) { if (int.TryParse(textBox.Text, out int result)) { multiplier = result; } isConfirmed = true; }

            return (isConfirmed, multiplier);
        }

        public static (bool decision, string dxfPath) GetDecisionAndEditedDxfPath(string dxfPath)
        {
            bool decision = false; string dxfPathEdited = dxfPath;

            using var prompt = new Form { Width = 800, Height = 200, FormBorderStyle = FormBorderStyle.FixedDialog, Text = "Czy wykonać zapis Dxf?", StartPosition = FormStartPosition.CenterScreen };

            Label label = new() { Left = 20, Top = 20, Width = 740, Text = "Wygenerowana ścieżka (do edycji):" }; TextBox textBox = new() { Left = 20, Top = 50, Width = 740, Text = dxfPath }; Button yesButton = new() { Text = "Tak", Left = 550, Width = 100, Top = 100, DialogResult = DialogResult.Yes }; Button noButton = new() { Text = "Nie", Left = 660, Width = 100, Top = 100, DialogResult = DialogResult.No };

            prompt.Controls.Add(label); prompt.Controls.Add(textBox); prompt.Controls.Add(yesButton); prompt.Controls.Add(noButton); prompt.AcceptButton = yesButton; prompt.CancelButton = noButton;

            if (prompt.ShowDialog() == DialogResult.Yes) { decision = true; dxfPathEdited = textBox.Text; }

            return (decision, dxfPathEdited);
        }

        public static (bool decision, string stepPath) GetDecisionAndEditedStepPath(string stepPath)
        {
            bool decision = false; string stepPathEdited = stepPath;

            using var prompt = new Form { Width = 800, Height = 200, FormBorderStyle = FormBorderStyle.FixedDialog, Text = "Czy wykonać zapis pliku STEP?", StartPosition = FormStartPosition.CenterScreen };

            Label label = new() { Left = 20, Top = 20, Width = 740, Text = "Wygenerowana ścieżka (do edycji):" }; TextBox textBox = new() { Left = 20, Top = 50, Width = 740, Text = stepPath }; Button yesButton = new() { Text = "Tak", Left = 550, Width = 100, Top = 100, DialogResult = DialogResult.Yes }; Button noButton = new() { Text = "Nie", Left = 660, Width = 100, Top = 100, DialogResult = DialogResult.No };

            prompt.Controls.Add(label); prompt.Controls.Add(textBox); prompt.Controls.Add(yesButton); prompt.Controls.Add(noButton); prompt.AcceptButton = yesButton; prompt.CancelButton = noButton;

            if (prompt.ShowDialog() == DialogResult.Yes) { decision = true; stepPathEdited = textBox.Text; }

            return (decision, stepPathEdited);
        }

        private static string PartsListHelper(List<string> savedSettings)
        {
            string selected = null;

            using var prompt = new Form { Width = 300, Height = 200, FormBorderStyle = FormBorderStyle.FixedDialog, Text = "Właściwości tabeli listy części.", StartPosition = FormStartPosition.CenterScreen };

            ListBox listBox = new() { Left = 10, Top = 10, Width = 200, Height = 120 }; foreach (var s in savedSettings) { listBox.Items.Add(s); }

            Button confirmation = new() { Text = "OK", Left = 10, Width = 80, Top = 135, DialogResult = DialogResult.OK };

            prompt.Controls.Add(listBox); prompt.Controls.Add(confirmation); prompt.AcceptButton = confirmation;

            if (prompt.ShowDialog() == DialogResult.OK) { if (listBox.SelectedItem != null) { selected = listBox.SelectedItem.ToString(); } }

            return selected;
        }

        public static string GetPartsListType(SeApp application, SeAssembly assembly)
        {
            SeDocuments documents = null; SeDraft draft = null; SeDraftSheet sheet = null; SeDrawingViews drawingViews = null; SeDrawingView drawingView = null; SeModelLinks modelLinks = null; SeModelLink modelLink = null; SePartsLists partsLists = null; SePartsList partsList = null;

            try
            {
                documents = application.Documents; draft = (SeDraft)documents.Add("SolidEdge.DraftDocument", Missing.Value); sheet = draft.ActiveSheet; modelLinks = draft.ModelLinks; modelLink = modelLinks.Add(assembly.FullName); drawingViews = sheet.DrawingViews;

                drawingView = drawingViews.AddAssemblyView(modelLink, SeViewOrientation.igFrontView, 0.1, 0.2, 0.2, SeAssemblyDrawingViewType.seAssemblyDesignedView);

                partsLists = draft.PartsLists; partsList = partsLists.AddEx(drawingView, 0, "", 0, 1);

                Array listOfSavedSettings = Array.CreateInstance(typeof(object), 0); partsList.GetListOfSavedSettings(out int numSavedSettings, ref listOfSavedSettings);

                var settingsList = new List<string>(); if (listOfSavedSettings != null) { foreach (var o in listOfSavedSettings) { if (o != null) settingsList.Add(o.ToString()); } }

                if (settingsList.Count == 0) settingsList.Add("<No saved parts list styles available>"); string chosen = PartsListHelper(settingsList);

                return chosen;
            }
            finally
            {
                CoreUtils.ReleaseCom(ref partsList); CoreUtils.ReleaseCom(ref partsLists); CoreUtils.ReleaseCom(ref modelLink); CoreUtils.ReleaseCom(ref modelLinks); CoreUtils.ReleaseCom(ref drawingView); CoreUtils.ReleaseCom(ref drawingViews); CoreUtils.ReleaseCom(ref sheet); draft?.Close(false); CoreUtils.ReleaseCom(ref draft); CoreUtils.ReleaseCom(ref documents);
            }
        }
    }
}