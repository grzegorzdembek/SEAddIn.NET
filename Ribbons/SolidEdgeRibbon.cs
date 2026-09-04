using SolidEdgeAdd_In.Commands;
using SolidEdgeAdd_In.Utils;

namespace SolidEdgeAdd_In.Ribbons
{
    public class SolidEdgeRibbon : Ribbon
    {
        public SolidEdgeRibbon(SeApp application)
        {
            this.Application = application;
            RibbonTab tab = AddTab("AddIn");
            RibbonGroup draftGroup = tab.AddGroup("Draft Environment");
            RibbonGroup partGroup = tab.AddGroup("Part Environment");
            RibbonGroup assemblyGroup1 = tab.AddGroup("Assembly Environment");
            RibbonGroup assemblyGroup2 = tab.AddGroup("Assembly Environment");
            RibbonGroup assemblyGroup3 = tab.AddGroup("Assembly Environment");
            RibbonGroup assemblyGroup4 = tab.AddGroup("Assembly Environment");
            RibbonGroup generalGroup = tab.AddGroup("General");


            RibbonButton saveDraftButton = new (1)
            {
                Label = "Save PDF and DXF",
                ScreenTip = "Saves the active drawing as PDF and DXF.",
                SuperTip = "It requires the draft document to be open and saved."
            };
            saveDraftButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeDraft draft)
                    {
                        SaveAsDxfAndPdfCommand.Execute(draft);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            draftGroup.AddControl(saveDraftButton);


            RibbonButton saveStepButton = new(2)
            {
                Label = "Save As Step",
                ScreenTip = "Saves the active document as a STEP file.",
                SuperTip = "It requires the part or sheet metal document to be open and saved."
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
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            partGroup.AddControl(saveStepButton);


            RibbonButton saveFlatPatternButton = new(3)
            {
                Label = "Save Flat Pattern As Dxf",
                ScreenTip = "Saves the flat pattern of the active document as a DXF file.",
                SuperTip = "It requires the part or sheet metal document to be open, saved, and have a flat pattern prepared."
            };
            saveFlatPatternButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SePart partDocument || document is SeSheetMetal sheetMetalDocument)
                    {
                        SaveFlatPatternAsDxfCommand.Execute(document);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            partGroup.AddControl(saveFlatPatternButton);


            RibbonButton exportDxfsButton = new(4)
            {
                Label = "Export Dxfs",
                ScreenTip = "Exports DXF files of all flat patterns for sheet metal parts from the active assembly into a dated package folder.",
                SuperTip = "It requires the top level assembly to be open and accessible."
            };
            exportDxfsButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        ExportDxfsCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            assemblyGroup1.AddControl(exportDxfsButton);


            RibbonButton exportPartsListButton = new(5)
            {
                Label = "Export Parts List",
                ScreenTip = "Exports the parts list from the active assembly to Excel, optionally generating and inserting part thumbnails.",
                SuperTip = "It requires the top level assembly to be open and accessible."
            };
            exportPartsListButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        ExportPartsListCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            assemblyGroup1.AddControl(exportPartsListButton);


            RibbonButton exportOccurrencelListButton = new(6)
            {
                Label = "Export Occurrences List",
                ScreenTip = "Exports filtered occurrence lists grouped by selected part types into separate Excel files, including thumbnails.",
                SuperTip = "It requires the top level assembly to be open and accessible."
            };
            exportOccurrencelListButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        ExportOccurrencesListCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            assemblyGroup2.AddControl(exportOccurrencelListButton);


            RibbonButton setCountPropertyButton = new(7)
            {
                Label = "Set Count Property",
                ScreenTip = "Sets a new count property to specific instances (Assembly, Part, Sheet Metal, Steelmaking) in the active assembly.",
                SuperTip = "It requires the top level assembly to be open and accessible."
            };
            setCountPropertyButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        SetCountPropertyCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            assemblyGroup1.AddControl(setCountPropertyButton);


            RibbonButton clearDxfDateButton = new(8)
            {
                Label = "Clear DxfDate",
                ScreenTip = "Clears the DxfDate property exclusively for sheet metal (Type B) parts in the current assembly.",
                SuperTip = "It requires the top level assembly to be open and accessible."
            };
            clearDxfDateButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        ClearDxfDateCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            assemblyGroup2.AddControl(clearDxfDateButton);

            RibbonButton copyDrawingsButton = new(9)
            {
                Label = "Copy Drawings",
                ScreenTip = "Copies PDF and DXF drawings based on an Excel summary list into a selected directory, updating the Excel file with the copy status.",
                SuperTip = "It requires any document to be open, and a previously generated packages directory with an Excel summary."
            };
            copyDrawingsButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    CopyDrawingsCommand.Execute(document);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            generalGroup.AddControl(copyDrawingsButton);


            RibbonButton organizeDrawingsButton = new(10)
            {
                Label = "Organise Drawings",
                ScreenTip = "Organises PDF and DXF drawings by sorting them into subdirectories based on their part type within a new assembly folder.",
                SuperTip = "It requires the assembly document to be open, accessible, and have drawings (.dxf and .pdf) prepared in the Drawings folder."
            };
            organizeDrawingsButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        OrganiseDrawingsCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            assemblyGroup2.AddControl(organizeDrawingsButton);


            RibbonButton refreshTreeButton = new(11)
            {
                Label = "Refresh Tree",
                ScreenTip = "Refreshes and loads all files in the assembly tree to avoid missing references.",
                SuperTip = "It requires the top level assembly to be open and accessible."
            };
            refreshTreeButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        RefreshTreeCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            assemblyGroup3.AddControl(refreshTreeButton);


            RibbonButton fitAndUpdateViewButton = new(12)
            {
                Label = "Fit And Update Views",
                ScreenTip = "Fits and updates the views for all models found in the active assembly, applying the ISO named view.",
                SuperTip = "It requires the top level assembly to be open and accessible."
            };
            fitAndUpdateViewButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        FitAndUpdateViewsCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            assemblyGroup3.AddControl(fitAndUpdateViewButton);


            RibbonButton makeDrawingsDirectoryButton = new(13)
            {
                Label = "Make Drawings Directory",
                ScreenTip = "Makes a new directory for drawings and copies all new PDF and DXF files from the project into it.",
                SuperTip = "It requires the top level assembly to be open, accessible, and have drawings (.dxf and .pdf) prepared in the project folder."
            };
            makeDrawingsDirectoryButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        MakeDrawingsDirectoryCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            assemblyGroup3.AddControl(makeDrawingsDirectoryButton);
            

            RibbonButton openDrawingButton = new(14)
            {
                Label = "Open Drawing",
                ScreenTip = "Opens PDF drawings for the selected occurrences, including deeply nested parts, or for the active document itself.",
                SuperTip = "It requires a document to be open and accessible, with PDF drawings available in the project directory."
            };
            openDrawingButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    OpenDrawingCommand.Execute(document);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            generalGroup.AddControl(openDrawingButton);


            RibbonButton renamePartNumberButton = new(15)
            {
                Label = "Rename Part Number",
                ScreenTip = "Renames the part number for the selected occurrences, safely updating references, copying drawing files, and updating drawing links.",
                SuperTip = "It requires the top level assembly to be open and accessible, with specific occurrences selected."
            };
            renamePartNumberButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        RenamePartNumberCommand.Execute(assembly);
                    }
                    else
                    {
                        MessageBox.Show("Command will not execute for this document.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            assemblyGroup4.AddControl(renamePartNumberButton);

        }
    }
}