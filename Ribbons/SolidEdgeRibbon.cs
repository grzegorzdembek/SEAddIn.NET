using SolidEdgeAdd_In.Main;
using SolidEdgeAdd_In.Utils;
using System.Xml.Linq;

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
            var assemblyGroup = tab.AddGroup("Złożenie");
            var generalGroup = tab.AddGroup("Ogólne");

            var saveDraftButton = new RibbonButton(1)
            {
                Label = "Zapisz PDF i DXF",
                ScreenTip = "Zapisuje otwarty rysunek jako PDF i DXF",
                SuperTip = "Zapisany zostanie PDF i DXF rysunku, w tej samej lokalizacji"
            };
            saveDraftButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeDraft draft) SaveAsDxfAndPdfCommand.Execute(draft);
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                finally { CoreUtils.ReleaseCom(ref document); }
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
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SePart partDocument || document is SeSheetMetal sheetMetalDocument)
                    {
                        SaveAsDxfCommand.Execute(document);
                    }
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                finally { CoreUtils.ReleaseCom(ref document); }
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
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SePart partDocument || document is SeSheetMetal sheetMetalDocument)
                    {
                        SaveAsStepCommand.Execute(document);
                    }
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                finally { CoreUtils.ReleaseCom(ref document); }
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
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly) ExportDxfsCommand.Execute(assembly);
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                finally { CoreUtils.ReleaseCom(ref document); }
            };
            assemblyGroup.AddControl(exportDxfsButton);

            var exportPartsListButton = new RibbonButton(5)
            {
                Label = "Eksportuj listę części",
                ScreenTip = "Zapisuje listę części z rysunku otwartego złożenia do Excela",
                SuperTip = "Zapisany zostanie arkusz Excela z tabelą (lista części) otwartego złożenia, w tej samej lokalizacji"
            };
            exportPartsListButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly) ExportPartsListCommand.Execute(assembly);
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                finally { CoreUtils.ReleaseCom(ref document); }
            };
            assemblyGroup.AddControl(exportPartsListButton);

            var setCountPropertyButton = new RibbonButton(6)
            {
                Label = "Dodaj właściwość ilości",
                ScreenTip = "Dodaje nową właściwość (ilość) dla każdego wystąpienia z typem A,B,C",
                SuperTip = "Dodane zostaną właściwości (ilości) dla wszystkich wystąpień w otwartym złożeniu"
            };
            setCountPropertyButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly) SetCountPropertyCommand.Execute(assembly);
                    else MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                }
                finally { CoreUtils.ReleaseCom(ref document); }
            };
            assemblyGroup.AddControl(setCountPropertyButton);

            var Ribbon7 = new RibbonButton(7)
            {
                Label = "Dodaj rysunki",
                ScreenTip = "Dodaje folder Rysunki ",
                SuperTip = "Dodane zostaną rysunki na podstawie pliku Excel w folderze Paczki."
            };
            Ribbon7.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    ExportDrawingsCommand.Execute(document);
                }
                finally { CoreUtils.ReleaseCom(ref document); }
                
            };
            generalGroup.AddControl(Ribbon7);
        }
    }
}