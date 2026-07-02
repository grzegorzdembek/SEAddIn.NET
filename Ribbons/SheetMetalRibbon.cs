using PartSaveFlatPatternAsDxf = SolidEdgeAdd_In.Main.PartEnviroment.SaveFlatPatternAsDxf;
using SheetMetalSaveFlatPatternAsDxf = SolidEdgeAdd_In.Main.SheetMetalEnviroment.SaveFlatPatternAsDxf;

using PartSaveAsStep = SolidEdgeAdd_In.Main.PartEnviroment.SaveAsStep;
using SheetMetalSaveAsStep = SolidEdgeAdd_In.Main.SheetMetalEnviroment.SaveAsStep;


namespace SolidEdgeAdd_In.Ribbons
{
    public class SheetMetalRibbon : Ribbon
    {
        public SheetMetalRibbon(SolidEdgeFramework.Application application)
        {
            this.Application = application;
            var tab = AddTab("Dodatki"); var metalSheetGroup = tab.AddGroup("Blachy");
            /*---"BLACHY"---*/
            var SaveFlatPatternAsDxfButton = new RibbonButton(4)
            {
                Label = "Zapisz rozwinięcie.",
                ScreenTip = "Zapisuje rozwinięcie jako DXF.",
                SuperTip = "Zapisany zostanie DXF rozwinięcia otwartej blachy, w tej samej lokalizacji."
            };
            SaveFlatPatternAsDxfButton.Click += (control) =>
            {
                try { var sheetMetal = (SolidEdgePart.SheetMetalDocument)application.ActiveDocument; SheetMetalSaveFlatPatternAsDxf.AddIn(sheetMetal); }
                catch { var part = (SolidEdgePart.PartDocument)application.ActiveDocument; PartSaveFlatPatternAsDxf.AddIn(part); }
            };
            metalSheetGroup.AddControl(SaveFlatPatternAsDxfButton);

            /*---"ZAPIS"---*/
            var savingGroup = tab.AddGroup("Zapis");
            var SaveAsStepButton = new RibbonButton(5)
            {
                Label = "Zapisz plik STEP.",
                ScreenTip = "Zapisuje otwarty plik jako STEP.",
                SuperTip = "Zapisany zostanie STEP otwartego pliku, w tej samej lokalizacji."
            };
            SaveAsStepButton.Click += (control) =>
            {
                try { var sheetMetal = (SolidEdgePart.SheetMetalDocument)application.ActiveDocument; SheetMetalSaveAsStep.AddIn(sheetMetal); }
                catch { var part = (SolidEdgePart.PartDocument)application.ActiveDocument; PartSaveAsStep.AddIn(part); }
            };
            savingGroup.AddControl(SaveAsStepButton);
        }
    }
}