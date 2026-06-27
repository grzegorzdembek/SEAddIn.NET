using SolidEdgeAdd_In.Main.DraftEnviroment;

namespace SolidEdgeAdd_In.Ribbons
{
    public class DraftRibbon : Ribbon
    {
        public DraftRibbon(SolidEdgeFramework.Application application)
        {
            this.Application = application;
            var tab = AddTab("Dodatki"); var savingGroup = tab.AddGroup("Zapis");
            /*---"ZAPIS"---*/
            var SaveDraftAsDxfAndPdfButton = new RibbonButton(1)
            {
                Label = "Zapisz jako PDF i DXF.",
                ScreenTip = "Zapisuje otwarty rysunek jako PDF i DXF.",
                SuperTip = "Zapisany zostanie PDF i DXF rysunku, w tej samej lokalizacji."
            };
            SaveDraftAsDxfAndPdfButton.Click += (control) => { var draft = (SolidEdgeDraft.DraftDocument)application.ActiveDocument;SaveDraftAsDxfAndPdf.AddIn(draft);};
            savingGroup.AddControl(SaveDraftAsDxfAndPdfButton);
        }
    }
}