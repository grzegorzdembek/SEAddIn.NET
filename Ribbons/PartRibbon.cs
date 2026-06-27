using SolidEdgeAdd_In.Main.PartEnviroment;

namespace SolidEdgeAdd_In.Ribbons
{
    public class PartRibbon : Ribbon
    {
        public PartRibbon(SolidEdgeFramework.Application application)
        {
            this.Application = application;
            var tab = AddTab("Dodatki"); var metalSheetGroup = tab.AddGroup("Blachy");
            /*---"BLACHY"---*/
            var SaveFlatPatternAsDxfButton = new RibbonButton(2)
            {
                Label = "Zapisz rozwinięcie.",
                ScreenTip = "Zapisuje rozwinięcie jako DXF.",
                SuperTip = "Zapisany zostanie DXF rozwinięcia otwartej blachy, w tej samej lokalizacji."
            };
            SaveFlatPatternAsDxfButton.Click += (control) => { var part = (SolidEdgePart.PartDocument)application.ActiveDocument;SaveFlatPatternAsDxf.AddIn(part);};
            metalSheetGroup.AddControl(SaveFlatPatternAsDxfButton);
        }
    }
}