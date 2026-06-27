using PartSaveFlatPatternAsDxf = SolidEdgeAdd_In.Main.PartEnviroment.SaveFlatPatternAsDxf;
using SheetMetalSaveFlatPatternAsDxf = SolidEdgeAdd_In.Main.SheetMetalEnviroment.SaveFlatPatternAsDxf;

namespace SolidEdgeAdd_In.Ribbons
{
    public class SheetMetalRibbon : Ribbon
    {
        public SheetMetalRibbon(SolidEdgeFramework.Application application)
        {
            this.Application = application;
            var tab = AddTab("Dodatki"); var metalSheetGroup = tab.AddGroup("Blachy");
            /*---"BLACHY"---*/
            var SaveFlatPatternAsDxfButton = new RibbonButton(3)
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
        }
    }
}