using SolidEdgeAdd_In.Commands;

namespace SolidEdgeAdd_In.Ribbons
{
    public class SolidEdgeRibbon : Ribbon
    {
        public SolidEdgeRibbon(SeApp application)
        {
            this.Application = application;

            RibbonTab tab = AddTab("AddIns");

            RibbonGroup draftGroup = tab.AddGroup("Draft Environment");
            RibbonGroup partGroup = tab.AddGroup("Part Environment");
            RibbonGroup assemblyGroup1 = tab.AddGroup("Assembly Environment");
            RibbonGroup assemblyGroup2 = tab.AddGroup("Assembly Environment");
            RibbonGroup generalGroup = tab.AddGroup("General");

            /*- _____1_____ -*/
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

            /*- _____2_____ -*/
            RibbonButton saveFlatPatternButton = new (2)
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
                        SaveAsDxfCommand.Execute(document);
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

            /*- _____3_____ -*/
            RibbonButton saveStepButton = new (3)
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

            /*- _____4_____ -*/
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

            /*- _____5_____ -*/
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
                        MessageBox.Show("Makro nie zadziała dla tego dokumentu");
                    }
                }
                catch (Exception ex) 
                { 
                    MessageBox.Show(ex.Message); 
                }
            };

            assemblyGroup1.AddControl(exportPartsListButton);

            /*- _____6_____ -*/
            RibbonButton setCountPropertyButton = new (6)
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

            assemblyGroup2.AddControl(setCountPropertyButton);

            /*- _____7_____ -*/
            RibbonButton copyDrawingsButton = new (7)
            {
                Label = "Copy Drawings",
                ScreenTip = "Copies drawings to the choosen directory.",
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

            /*- _____8_____ -*/
            RibbonButton clearDxfDateButton = new (8)
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

            /*- _____9_____ -*/
            RibbonButton organizeDrawingsButton = new(9)
            {
                Label = "Organize Drawings",
                ScreenTip = "Organizes drawings in the project directory.",
                SuperTip = "Organizes drawings (files .dxf and .pdf) in the project directory."
            };

            organizeDrawingsButton.Click += (control) =>
            {
                try
                {
                    MessageBox.Show("not implemented yet.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            };

            assemblyGroup2.AddControl(organizeDrawingsButton);

            /*- _____10_____ -*/
            RibbonButton generateShotsButton = new(10)
            {
                Label = "Generate Shots",
                ScreenTip = "Generates shots of all parts in the assembly",
                SuperTip = "Generates shots of all parts in the assembly and lock them in the Miniatury directory."
            };

            generateShotsButton.Click += (control) =>
            {
                try
                {
                    MessageBox.Show("not implemented yet.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            };

            assemblyGroup2.AddControl(generateShotsButton);

            /*- _____11_____ -*/
            RibbonButton transformDataButton = new(11)
            {
                Label = "Transform Data",
                ScreenTip = "Transforms data from Excel.",
                SuperTip = "Transforms data from Excel."
            };

            transformDataButton.Click += (control) =>
            {
                try
                {
                    MessageBox.Show("not implemented yet.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }; 

            generalGroup.AddControl(transformDataButton);
        }
    }
}