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
            RibbonGroup generalGroup = tab.AddGroup("General");


            RibbonButton saveDraftButton = new (1)
            {
                Label = "Save PDF and DXF",
                ScreenTip = "Saves the active drawing as PDF and DXF.",
                SuperTip = "The drawing will be saved as both PDF and DXF in the project directory."
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
                        MessageBox.Show("AddIn will not execute for this document.");
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


            RibbonButton saveStepButton = new (2)
            {
                Label = "Save STEP",
                ScreenTip = "Saves the active document as a STEP file.",
                SuperTip = "The active document will be saved as a STEP file in the project directory."
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
                        MessageBox.Show("AddIn will not execute for this document.");
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


            RibbonButton saveFlatPatternButton = new (3)
            {
                Label = "Save Flat Pattern",
                ScreenTip = "Saves the flat pattern as a DXF.",
                SuperTip = "The flat pattern will be saved as a DXF in the project directory."
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
                        MessageBox.Show("AddIn will not execute for this document.");
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


            RibbonButton exportDxfsButton = new (4)
            {
                Label = "Export DXFs",
                ScreenTip = "Exports DXFs of all flat patterns for files - (.par) and (.psm) from the open assembly.",
                SuperTip = "The DXFs of all flat patterns for files - (par) and (psm) from the open assembly will be exported with excel summary."
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
                        MessageBox.Show("AddIn will not execute for this document.");
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


            RibbonButton exportPartsListButton = new (5)
            {
                Label = "Export Parts List",
                ScreenTip = "Exports the parts list from the open assembly to Excel.",
                SuperTip = "An Excel sheet containing the parts list of the open assembly will be saved in the list directory."
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
                        MessageBox.Show("AddIn will not execute for this document.");
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


            RibbonButton exportOccurrencelListButton = new (6)
            {
                Label = "Export Occurrence List",
                ScreenTip = "Exports the occurrence list for the active assembly.",
                SuperTip = "Select the occurrence type and export them to the Excel."
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
                        MessageBox.Show("AddIn will not execute for this document.");
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


            RibbonButton setCountPropertyButton = new (7)
            {
                Label = "Set Count Property",
                ScreenTip = "Adds a new property (count) for each instance with type A,B,C.",
                SuperTip = "Added properties (counts) will be available for all instances in the open assembly."
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
                        MessageBox.Show("AddIn will not execute for this document.");
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


            RibbonButton clearDxfDateButton = new (8)
            {
                Label = "Clear DxfDate Property",
                ScreenTip = "Removes the DxfDate property for files in the current assembly.",
                SuperTip = "This can be used before exporting to DXF."
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
                        MessageBox.Show("AddIn will not execute for this document.");
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

            RibbonButton copyDrawingsButton = new (9)
            {
                Label = "Copy Drawings",
                ScreenTip = "Copies drawings to the selected directory.",
                SuperTip = "Copied drawings will be placed in the chosen directory and added column in excel report."
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


            RibbonButton organizeDrawingsButton = new (10)
            {
                Label = "Organize Drawings",
                ScreenTip = "Organizes drawings in the project directory.",
                SuperTip = "Organizes drawings (files .dxf and .pdf) in the project directory."
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
                        MessageBox.Show("AddIn will not execute for this document.");
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


            RibbonButton refreshTreeButton = new (11)
            {
                Label = "Refresh Tree",
                ScreenTip = "Loads all files.",
                SuperTip = "Use this to avoid missing files."
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
                        MessageBox.Show("AddIn will not execute for this document.");
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


            RibbonButton fitAndUpdateViewButton = new (12)
            {
                Label = "Fit And Update Views",
                ScreenTip = "Opens files to fit and update their views.",
                SuperTip = "Opens each document from the assembly, centers the model in the window (Fit), refreshes the display (Update), and closes the file."
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
                        MessageBox.Show("AddIn will not execute for this document.");
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

            /*
            RibbonButton shotThumbnailsButton = new (13)
            {
                Label = "Shot Thumbnails",
                ScreenTip = "Generates high quality thumbnail images.",
                SuperTip = "Temporarily hides planes and dimensions, fits view, takes a snapshot and restores the original state."
            };
            shotThumbnailsButton.Click += (control) =>
            {
                SeDocument document = null;
                try
                {
                    document = application.ActiveDocument;
                    if (document is SeAssembly assembly)
                    {
                        ShotThumbnailsCommand.Execute(assembly);
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
                finally
                {
                    Helpers.ReleaseCom(ref document);
                }
            };
            assemblyGroup3.AddControl(shotThumbnailsButton);
            */
        }
    }
}