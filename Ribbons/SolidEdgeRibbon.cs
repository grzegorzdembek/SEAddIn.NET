using SolidEdgeAdd_In.Commands;
using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Ribbons
{
    public class SolidEdgeRibbon : Ribbon
    {
        public SolidEdgeRibbon(SeApp application)
        {
            this.Application = application;
            var tab = AddTab("Dodatki");
            var draftGroup = tab.AddGroup("Rysunek");
            var partGroup = tab.AddGroup("Część");
            var assemblyGroup1 = tab.AddGroup("Złożenie");
            var assemblyGroup2 = tab.AddGroup("Złożenie");
            var generalGroup = tab.AddGroup("Ogólne");

            var saveDraftButton = new RibbonButton(1)
            {
                Label = "Zapisz PDF i DXF",
                ScreenTip = "Zapisuje otwarty rysunek jako PDF i DXF",
                SuperTip = "Zapisany zostanie PDF i DXF rysunku, w tej samej lokalizacji"
            };
            saveDraftButton.Click += (control) =>
            {
                try
                {
                    var document = application.ActiveDocument; // FIXED: Do not release ActiveDocument
                    if (document is SeDraft draft) SaveAsDxfAndPdfCommand.Execute(draft);
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            };
            draftGroup.AddControl(saveDraftButton);

            var saveFlatPatternButton = new RibbonButton(2)
            {
                Label = "Zapisz rozwinięcie",
                ScreenTip = "Zapisuje rozwinięcie jako DXF",
                SuperTip = "Zapisany zostanie DXF rozwinięcia otwartej blachy, w tej samej lokalizacji"
            };
            saveFlatPatternButton.Click += (control) =>
            {
                try
                {
                    var document = application.ActiveDocument;
                    if (document is SePart partDocument || document is SeSheetMetal sheetMetalDocument)
                    {
                        SaveAsDxfCommand.Execute(document);
                    }
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            };
            partGroup.AddControl(saveFlatPatternButton);

            var saveStepButton = new RibbonButton(3)
            {
                Label = "Zapisz STEP",
                ScreenTip = "Zapisuje otwarty plik jako STEP",
                SuperTip = "Zapisany zostanie STEP otwartego pliku, w tej samej lokalizacji"
            };
            saveStepButton.Click += (control) =>
            {
                try
                {
                    var document = application.ActiveDocument;
                    if (document is SePart partDocument || document is SeSheetMetal sheetMetalDocument)
                    {
                        SaveAsStepCommand.Execute(document);
                    }
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            };
            partGroup.AddControl(saveStepButton);

            var exportDxfsButton = new RibbonButton(4)
            {
                Label = "Eksportuj DFXy",
                ScreenTip = "Zapisuje DXFy wszystkich części (par) i blach (psm) z otwartego złożenia",
                SuperTip = "Zapisane zostaną DXFy wszystkich części (par) i blach (psm) z otwartego złożenia, w tej samej lokalizacji"
            };
            exportDxfsButton.Click += (control) =>
            {
                try
                {
                    var document = application.ActiveDocument;
                    if (document is SeAssembly assembly) ExportDxfsCommand.Execute(assembly);
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            };
            assemblyGroup1.AddControl(exportDxfsButton);

            var exportPartsListButton = new RibbonButton(5)
            {
                Label = "Eksportuj listę części",
                ScreenTip = "Zapisuje listę części z rysunku otwartego złożenia do Excela",
                SuperTip = "Zapisany zostanie arkusz Excela z tabelą (lista części) otwartego złożenia, w tej samej lokalizacji"
            };
            exportPartsListButton.Click += (control) =>
            {
                try
                {
                    var document = application.ActiveDocument;
                    if (document is SeAssembly assembly) ExportPartsListCommand.Execute(assembly);
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            };
            assemblyGroup1.AddControl(exportPartsListButton);

            var setCountPropertyButton = new RibbonButton(6)
            {
                Label = "Dodaj właściwość ilości",
                ScreenTip = "Dodaje nową właściwość (ilość) dla każdego wystąpienia z typem A,B,C",
                SuperTip = "Dodane zostaną właściwości (ilości) dla wszystkich wystąpień w otwartym złożeniu"
            };
            setCountPropertyButton.Click += (control) =>
            {
                try
                {
                    var document = application.ActiveDocument;
                    if (document is SeAssembly assembly) SetCountPropertyCommand.Execute(assembly);
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            };
            assemblyGroup2.AddControl(setCountPropertyButton);

            var copyDrawingsButton = new RibbonButton(7)
            {
                Label = "Dodaj rysunki",
                ScreenTip = "Dodaje folder Rysunki ",
                SuperTip = "Dodane zostaną rysunki na podstawie pliku Excel w folderze Paczki."
            };
            copyDrawingsButton.Click += (control) =>
            {
                try
                {
                    var document = application.ActiveDocument;
                    CopyDrawingsCommand.Execute(document);
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            };
            generalGroup.AddControl(copyDrawingsButton);

            var clearDxfDateButton = new RibbonButton(8)
            {
                Label = "Usuń właściwość DxfDate.",
                ScreenTip = "Usuwa właściwość dla plików w tym złożeniu",
                SuperTip = ""
            };
            clearDxfDateButton.Click += (control) =>
            {
                try
                {
                    var document = application.ActiveDocument;
                    if (document is SeAssembly assembly) ClearDxfDateCommand.Execute(assembly);
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            };
            assemblyGroup2.AddControl(clearDxfDateButton);
        }
    }
}