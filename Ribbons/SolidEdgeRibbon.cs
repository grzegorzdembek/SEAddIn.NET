using SolidEdgeAdd_In.Commands;

namespace SolidEdgeAdd_In.Ribbons
{
    public class SolidEdgeRibbon : Ribbon
    {
        public SolidEdgeRibbon(SeApp application)
        {
            this.Application = application;

            RibbonTab tab = AddTab("AddIn");

            RibbonGroup draftGroup = tab.AddGroup("Draft Environment"); // 1 Command 
            RibbonGroup partGroup = tab.AddGroup("Part Environment"); // 2 Command
            RibbonGroup assemblyGroup1 = tab.AddGroup("Assembly Environment"); // 3 Command
            RibbonGroup assemblyGroup2 = tab.AddGroup("Assembly Environment"); // 3 Command
            RibbonGroup generalGroup = tab.AddGroup("General"); // 1 Command

            /**************************************************************/
            /*- _____1_____DRAFT_______ -*/
            RibbonButton saveDraftButton = new (1)
            {
                Label = "Save PDF and DXF",
                ScreenTip = "Saves the active drawing as PDF and DXF.",
                SuperTip = "The drawing will be saved as both PDF and DXF in the project directory."
            };
            saveDraftButton.Click += (control) =>
            {
                try
                {
                    SeDocument document = application.ActiveDocument;
                    if (document is SeDraft draft)
                    { 
                        SaveAsDxfAndPdfCommand.Execute(draft);
                    }
                    else
                    { 
                        MessageBox.Show("AddIn will not execute for this document.");
                    }
                }
                catch (Exception ex) 
                { 
                    MessageBox.Show(ex.Message); 
                }
            };
            draftGroup.AddControl(saveDraftButton);

            /*- _____2_____PART/SheetMetal_______ -*/
            RibbonButton saveStepButton = new(2)
            {
                Label = "Save STEP",
                ScreenTip = "Saves the active document as a STEP file.",
                SuperTip = "The active document will be saved as a STEP file in the project directory."
            };

            saveStepButton.Click += (control) =>
            {
                try
                {
                    SeDocument document = application.ActiveDocument;
                    if (document is SePart partDocument || document is SeSheetMetal sheetMetalDocument)
                    {
                        SaveAsStepCommand.Execute(document);
                    }
                    else
                    {
                        MessageBox.Show("AddIn will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            };

            partGroup.AddControl(saveStepButton);

            /*- _____3_____PART/SHEETMETAL_______ -*/
            RibbonButton saveFlatPatternButton = new (3)
            {
                Label = "Save Flat Pattern",
                ScreenTip = "Saves the flat pattern as a DXF.",
                SuperTip = "The flat pattern will be saved as a DXF in the project directory."
            };

            saveFlatPatternButton.Click += (control) =>
            {
                try
                {
                    SeDocument document = application.ActiveDocument;
                    if (document is SePart partDocument || document is SeSheetMetal sheetMetalDocument)
                    {
                        SaveFlatPatternAsDxfCommand.Execute(document);
                    }
                    else
                    { 
                        MessageBox.Show("AddIn will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message); 
                }
            };
            partGroup.AddControl(saveFlatPatternButton);
            /**************************************************************/





            /**************************************************************/
            /*- _____4_____ASSEMBLY_______ -*/
            RibbonButton exportDxfsButton = new (4)
            {
                Label = "Export DXFs",
                ScreenTip = "Exports DXFs of all flat patterns for files - (.par) and (.psm) from the open assembly.",
                SuperTip = "The DXFs of all flat patterns for files - (par) and (psm) from the open assembly will be exported with excel report."
            };
            exportDxfsButton.Click += (control) =>
            {
                try
                {
                    SeDocument document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        ExportDxfsCommand.Execute(assembly);
                    }
                    else 
                    {
                        MessageBox.Show("AddIn will not execute for this document.");
                    }
                }
                catch (Exception ex) 
                { 
                    MessageBox.Show(ex.Message); 
                }
            };
            assemblyGroup1.AddControl(exportDxfsButton);

            /*- _____5_____ASSEMBLY_______ -*/
            RibbonButton exportPartsListButton = new (5)
            {
                Label = "Export Parts List",
                ScreenTip = "Exports the parts list from the open assembly to Excel.",
                SuperTip = "An Excel sheet containing the parts list of the open assembly will be saved in the project directory."
            };
            exportPartsListButton.Click += (control) =>
            {
                try
                {
                    SeDocument document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        ExportPartsListCommand.Execute(assembly);
                    }
                    else
                    { 
                        MessageBox.Show("AddIn will not execute for this document.");
                    }
                }
                catch (Exception ex) 
                { 
                    MessageBox.Show(ex.Message); 
                }
            };
            assemblyGroup1.AddControl(exportPartsListButton);

            /*- _____6_____ASSEMBLY_______ -*/
            RibbonButton exportOccurrencelListButton = new(6)
            {
                Label = "Export Occurrence List",
                ScreenTip = "Exports the occurrence list for the active assembly.",
                SuperTip = "Select the occurrence type and export them to the Excel."
            };
            exportOccurrencelListButton.Click += (control) =>
            {
                try
                {
                    SeDocument document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        ExportOccurrencesListCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("AddIn will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            };
            assemblyGroup2.AddControl(exportOccurrencelListButton);
            /**************************************************************/





            /**************************************************************/
            /*- _____7_____ASSEMBLY_______ -*/
            RibbonButton setCountPropertyButton = new (7)
            {
                Label = "Set Count Property",
                ScreenTip = "Adds a new property (count) for each instance with type A,B,C.",
                SuperTip = "Added properties (counts) will be available for all instances in the open assembly."
            };
            setCountPropertyButton.Click += (control) =>
            {
                try
                {
                    SeDocument document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    { 
                        SetCountPropertyCommand.Execute(assembly);
                    }
                    else
                    { 
                        MessageBox.Show("AddIn will not execute for this document.");
                    }
                }
                catch (Exception ex) 
                { 
                    MessageBox.Show(ex.Message); 
                }
            };
            assemblyGroup1.AddControl(setCountPropertyButton);

            /*- _____8_____ASSEMBLY_______ -*/
            RibbonButton clearDxfDateButton = new(8)
            {
                Label = "Clear DxfDate Property",
                ScreenTip = "Removes the DxfDate property for files in the current assembly.",
                SuperTip = "This can be used before exporting to DXF."
            };
            clearDxfDateButton.Click += (control) =>
            {
                try
                {
                    SeDocument document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        ClearDxfDateCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("AddIn will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            };
            assemblyGroup2.AddControl(clearDxfDateButton);
            /**************************************************************/




            /**************************************************************/
            /*- _____9_____GENERAL_______ -*/
            RibbonButton copyDrawingsButton = new (9)
            {
                Label = "Copy Drawings",
                ScreenTip = "Copies drawings to the selected directory.",
                SuperTip = "Copied drawings will be placed in the chosen directory and added column in excel report."
            };
            copyDrawingsButton.Click += (control) =>
            {
                try
                {
                    SeDocument document = application.ActiveDocument;
                    CopyDrawingsCommand.Execute(document);
                }
                catch (Exception ex) 
                { 
                    MessageBox.Show(ex.Message);
                }
            };
            generalGroup.AddControl(copyDrawingsButton);
        
            /*- _____10_____ASSEMBLY_______ -*/
            RibbonButton organizeDrawingsButton = new(10)
            {
                Label = "Organize Drawings",
                ScreenTip = "Organizes drawings in the project directory.",
                SuperTip = "Organizes drawings (files .dxf and .pdf) in the project directory."
            };
            organizeDrawingsButton.Click += (control) =>
            {
                try
                {
                    SeDocument document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        OrganiseDrawingsCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("AddIn will not execute for this document.");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            };
            assemblyGroup2.AddControl(organizeDrawingsButton);
            /**************************************************************/
        }
    }
}