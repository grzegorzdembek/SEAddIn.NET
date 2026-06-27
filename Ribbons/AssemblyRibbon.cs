using SolidEdgeAdd_In.Main.AssemblyEnviroment;

namespace SolidEdgeAdd_In.Ribbons 
{
    public class AssemblyRibbon : Ribbon
    {
        public AssemblyRibbon(SolidEdgeFramework.Application application)
        {
            this.Application = application;
            var tab = AddTab("Dodatki"); var metalSheetGroup = tab.AddGroup("Blachy"); var excelGroup = tab.AddGroup("Excel"); var propertiesGroup = tab.AddGroup("Właściwości i zmienne");
            /*---"BLACHY"---*/
            var SaveDxfOfPsmButton = new RibbonButton(4)
            {
                Label = "Zapisz DXFy blach.",
                ScreenTip = "Zapisuje DXFy wszystkich blach z otwartego złożenia.",
                SuperTip = "Zapisane zostaną DXY wszystkich blach (psm) z otwartego złożenia, w tej samej lokalizacji."
            };
            SaveDxfOfPsmButton.Click += (control) => { var assembly = (SolidEdgeAssembly.AssemblyDocument)application.ActiveDocument; SaveDxfOfPsm.AddIn(assembly);};
            metalSheetGroup.AddControl(SaveDxfOfPsmButton);
            /*---"BLACHY"---*/
            var SaveDxfOfPartsAndPsmButton = new RibbonButton(5)
            {
                Label = "Zapisz DXy części i blach.",
                ScreenTip = "Zapisuje DXFy wszystkich części (par) i blach (psm) z otwartego złożenia.",
                SuperTip = "Zapisane zostaną DXy wszystkich części (par) i blach (psm) z otwartego złożenia, w tej samej lokalizacji."
            };
            SaveDxfOfPartsAndPsmButton.Click += (control) => { var assembly = (SolidEdgeAssembly.AssemblyDocument)application.ActiveDocument; SaveDxfOfPartsAndPsm.AddIn(assembly);};
            metalSheetGroup.AddControl(SaveDxfOfPartsAndPsmButton);
            /*---"EXCEL"---*/
            var ExportPartsListButton = new RibbonButton(6)
            {
                Label = "Eksportuj listę części z rysunku do Excela.",
                ScreenTip = "Zapisuje listę części z rysunku otwartego złożenia do Excela.",
                SuperTip = "Zapisany zostanie arkusz Excela z tabelą (lista części) otwartego złożenia, w tej samej lokalizacji."
            };
            ExportPartsListButton.Click += (control) => { var assembly = (SolidEdgeAssembly.AssemblyDocument)application.ActiveDocument; ExportPartsList.AddIn(assembly);};
            excelGroup.AddControl(ExportPartsListButton);
            /*---"EXCEL"---*/
            var ExportOccurrencesListButton = new RibbonButton(7)
            {
                Label = "Eksportuj listę wystąpień do Excela.",
                ScreenTip = "Eksportuje listę wystąpień otwartego złożenia do Excela.",
                SuperTip = "Zapisany zostanie arkusz Excela z tabelą (lista wystąpień) otwartego złożenia, w tej samej lokalizacji."
            };
            ExportOccurrencesListButton.Click += (control) => { var assembly = (SolidEdgeAssembly.AssemblyDocument)application.ActiveDocument; ExportOccurrencesList.AddIn(assembly); };
            excelGroup.AddControl(ExportOccurrencesListButton);
            /*---"WŁAŚCIWOŚCI i ZMIENNE"---*/
            var SetCountPropertyButton = new RibbonButton(8)
            {
                Label = "Dodaj właściwość ilości.",
                ScreenTip = "Dodaje nową właściwość (ilość) dla każdego wystąpienia z typem A,B,C.",
                SuperTip = "Dodane zostaną właściwości (ilości) dla wszystkich wystąpień w otwartym złożeniu."
            };
            SetCountPropertyButton.Click += (control) => { var assembly = (SolidEdgeAssembly.AssemblyDocument)application.ActiveDocument; SetCountProperty.AddIn(assembly); };
            propertiesGroup.AddControl(SetCountPropertyButton);
            /*---"WŁAŚCIWOŚCI i ZMIENNE"---*/
            var PreparePartsAndMetalSheetsButton = new RibbonButton(9)
            {
                Label = "Przygotuj części tego złożenia.",
                ScreenTip = "Przygotuj zmienne części i ich typy.",
                SuperTip = "Przygotowane zostaną zmienne blach i ustawiony typ (B) oraz typ części (C)."
            };
            PreparePartsAndMetalSheetsButton.Click += (control) => { var assembly = (SolidEdgeAssembly.AssemblyDocument)application.ActiveDocument; PrepareParts.AddIn(assembly); };
            propertiesGroup.AddControl(PreparePartsAndMetalSheetsButton);
        }
    }
}