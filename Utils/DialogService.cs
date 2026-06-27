using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Reflection;

namespace SolidEdgeAdd_In.Utils
{
    public class DialogService
    {
        public static bool IsFileLocationsConfirmed()
        {
            string message = "Upewnij się, że wszystkie pliki .par i .psm, które są w tym samym folderze co otwarte złożenie, mają typ B lub C!";
            string caption = "Wymagane potwierdzenie";
            DialogResult result = MessageBox.Show(
                message,
                caption,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );
            return result == DialogResult.Yes;
        }

        public static bool IsShotsNeeded()
        {
            DialogResult result = MessageBox.Show(
                "Czy wygenerować miniatury?",
                "Miniatury",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            return result == DialogResult.Yes;
        }

        public static int GetMultiplier()
        {
            int multiplier = 1;

            using (Form prompt = new Form())
            {
                prompt.Width = 300;
                prompt.Height = 160;
                prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
                prompt.Text = "Wprowadzanie danych";
                prompt.StartPosition = FormStartPosition.CenterScreen;
                prompt.MinimizeBox = false;
                prompt.MaximizeBox = false;

                Label textLabel = new Label() { Left = 20, Top = 20, Text = "PODAJ MNOŻNIK:", AutoSize = true };
                TextBox textBox = new TextBox() { Left = 20, Top = 50, Width = 240 };
                Button confirmation = new Button() { Text = "OK", Left = 160, Width = 100, Top = 80, DialogResult = DialogResult.OK };

                prompt.Controls.Add(textLabel);
                prompt.Controls.Add(textBox);
                prompt.Controls.Add(confirmation);
                prompt.AcceptButton = confirmation;

                if (prompt.ShowDialog() == DialogResult.OK)
                {
                    if (int.TryParse(textBox.Text, out int result))
                    {
                        multiplier = result;
                    }
                }
            }

            return multiplier;
        }

        public static (bool decision, string dxfPath) GetDecisionAndEditedDxfPath(string dxfPath)
        {
            bool decision = false;
            string dxfPathEdited = dxfPath;

            using (Form prompt = new Form())
            {
                prompt.Width = 800;
                prompt.Height = 200;
                prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
                prompt.Text = "Czy wykonać zapis Dxf?";
                prompt.StartPosition = FormStartPosition.CenterScreen;

                Label label = new Label()
                {
                    Left = 20,
                    Top = 20,
                    Width = 740,
                    Text = "Wygenerowana ścieżka (do edycji):"
                };

                TextBox textBox = new TextBox()
                {
                    Left = 20,
                    Top = 50,
                    Width = 740,
                    Text = dxfPath
                };

                Button yesButton = new Button()
                {
                    Text = "Tak",
                    Left = 550,
                    Width = 100,
                    Top = 100,
                    DialogResult = DialogResult.Yes
                };

                Button noButton = new Button()
                {
                    Text = "Nie",
                    Left = 660,
                    Width = 100,
                    Top = 100,
                    DialogResult = DialogResult.No
                };

                prompt.Controls.Add(label);
                prompt.Controls.Add(textBox);
                prompt.Controls.Add(yesButton);
                prompt.Controls.Add(noButton);
                prompt.AcceptButton = yesButton;
                prompt.CancelButton = noButton;

                if (prompt.ShowDialog() == DialogResult.Yes)
                {
                    decision = true;
                    dxfPathEdited = textBox.Text;
                }
            }

            return (decision, dxfPathEdited);
        }

        private static string PartsListHelper(List<string> savedSettings)
        {
            string selected = null;

            using (Form prompt = new Form())
            {
                prompt.Width = 300;
                prompt.Height = 200;
                prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
                prompt.Text = "Właściwości tabeli listy części.";
                prompt.StartPosition = FormStartPosition.CenterScreen;

                ListBox listBox = new ListBox()
                {
                    Left = 10,
                    Top = 10,
                    Width = 200,
                    Height = 120
                };

                foreach (var s in savedSettings)
                {
                    listBox.Items.Add(s);
                }

                Button confirmation = new Button()
                {
                    Text = "OK",
                    Left = 10,
                    Width = 80,
                    Top = 135,
                    DialogResult = DialogResult.OK
                };

                prompt.Controls.Add(listBox);
                prompt.Controls.Add(confirmation);
                prompt.AcceptButton = confirmation;

                if (prompt.ShowDialog() == DialogResult.OK)
                {
                    if (listBox.SelectedItem != null)
                    {
                        selected = listBox.SelectedItem.ToString();
                    }
                }
            }

            return selected;
        }
        public static string GetPartsListType(SolidEdgeFramework.Application application, SolidEdgeAssembly.AssemblyDocument assembly)
        {
            SolidEdgeFramework.Documents documents = null;
            SolidEdgeDraft.DraftDocument draft = null;
            SolidEdgeDraft.Sheet sheet = null;
            SolidEdgeDraft.DrawingViews drawingViews = null;
            SolidEdgeDraft.DrawingView drawingView = null;
            SolidEdgeDraft.ModelLinks modelLinks = null;
            SolidEdgeDraft.ModelLink modelLink = null;
            SolidEdgeDraft.PartsLists partsLists = null;
            SolidEdgeDraft.PartsList partsList = null;

            try
            {
                documents = application.Documents;
                draft = (SolidEdgeDraft.DraftDocument)documents.Add("SolidEdge.DraftDocument", Missing.Value);
                sheet = draft.ActiveSheet;
                modelLinks = draft.ModelLinks;
                modelLink = modelLinks.Add(assembly.FullName);
                drawingViews = sheet.DrawingViews;
                drawingView = drawingViews.AddAssemblyView(
                    modelLink,
                    SolidEdgeDraft.ViewOrientationConstants.igFrontView,
                    0.1, 0.2, 0.2,
                    SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView);
                partsLists = draft.PartsLists;
                partsList = partsLists.AddEx(drawingView, 0, "", 0, 1);

                int numSavedSettings = 0;
                System.Array listOfSavedSettings = Array.CreateInstance(typeof(object), 0);
                partsList.GetListOfSavedSettings(out numSavedSettings, ref listOfSavedSettings);
                var settingsList = new List<string>();

                if (listOfSavedSettings != null)
                {
                    foreach (var o in listOfSavedSettings)
                    {
                        if (o != null) settingsList.Add(o.ToString());
                    }
                }

                if (settingsList.Count == 0) settingsList.Add("<No saved parts list styles available>");
                string chosen = PartsListHelper(settingsList);

                return chosen;
            }
            finally
            {
                CoreUtils.ReleaseCom(ref partsList);
                CoreUtils.ReleaseCom(ref partsLists);
                CoreUtils.ReleaseCom(ref modelLink);
                CoreUtils.ReleaseCom(ref modelLinks);
                CoreUtils.ReleaseCom(ref drawingView);
                CoreUtils.ReleaseCom(ref drawingViews);
                CoreUtils.ReleaseCom(ref sheet);
                draft?.Close(false);
                CoreUtils.ReleaseCom(ref draft);
                CoreUtils.ReleaseCom(ref documents);
            }
        }
    }
}