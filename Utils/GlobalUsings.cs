global using System;
global using System.Collections.Generic;
global using System.IO;
global using System.Linq;
global using System.Text;
global using System.Drawing;
global using System.Windows.Forms;
global using System.Reflection;
global using System.Runtime.InteropServices;
global using Microsoft.Win32;
global using System.Runtime.InteropServices.ComTypes;
global using System.Diagnostics;

global using SolidEdgeFramework;

global using SeISolidEdgeAddIn = SolidEdgeFramework.ISolidEdgeAddIn;
global using SeConnectMode = SolidEdgeFramework.SeConnectMode;
global using SeDisconnectMode = SolidEdgeFramework.SeDisconnectMode;


global using SeAddIn = SolidEdgeFramework.AddIn;
global using SeISEAddInEx = SolidEdgeFramework.ISEAddInEx;

global using SeApp = SolidEdgeFramework.Application;
global using SeDocument = SolidEdgeFramework.SolidEdgeDocument;
global using SeWindow = SolidEdgeFramework.Window;

global using SeAssembly = SolidEdgeAssembly.AssemblyDocument;
global using SePart = SolidEdgePart.PartDocument;
global using SeSheetMetal = SolidEdgePart.SheetMetalDocument;
global using SeDraft = SolidEdgeDraft.DraftDocument;

global using SeOccurrences = SolidEdgeAssembly.Occurrences;
global using SeOccurrence = SolidEdgeAssembly.Occurrence;

global using SeBends = SolidEdgePart.Bends; 
global using SeModels = SolidEdgePart.Models;
global using SeModel = SolidEdgePart.Model;
global using SeFlatPatternModels = SolidEdgePart.FlatPatternModels;
global using SeFlatPatternModel = SolidEdgePart.FlatPatternModel;
global using SeFlanges = SolidEdgePart.Flanges;

global using SeRefPlane = SolidEdgePart.RefPlane;
global using SeRefAxis = SolidEdgePart.RefAxis;
global using SeCoordinateSystem = SolidEdgePart.CoordinateSystem;
global using SeAsmRefPlane = SolidEdgeAssembly.AsmRefPlane;

global using SeDocuments = SolidEdgeFramework.Documents;
global using SeDraftSheet = SolidEdgeDraft.Sheet;
global using SeDrawingViews = SolidEdgeDraft.DrawingViews;
global using SeDrawingView = SolidEdgeDraft.DrawingView;
global using SeModelLinks = SolidEdgeDraft.ModelLinks;
global using SeModelLink = SolidEdgeDraft.ModelLink;
global using SePartsLists = SolidEdgeDraft.PartsLists;
global using SePartsList = SolidEdgeDraft.PartsList;

global using SeViewOrientation = SolidEdgeDraft.ViewOrientationConstants;
global using SeAssemblyDrawingViewType = SolidEdgeDraft.AssemblyDrawingViewTypeConstants;

global using SeFilePropertySets = SolidEdgeFileProperties.PropertySets;
global using SeFileProperties = SolidEdgeFileProperties.Properties;
global using SeFileProperty = SolidEdgeFileProperties.Property;

global using SePropertySets = SolidEdgeFramework.PropertySets;
global using SeProperties = SolidEdgeFramework.Properties;
global using SeProperty = SolidEdgeFramework.Property;

global using SeView = SolidEdgeFramework.View;
global using SeImageQualityType = SolidEdgeFramework.SeImageQualityType;


global using ExcelXlPlacement = Microsoft.Office.Interop.Excel.XlPlacement;
global using ExcelXlDeleteShiftDirection = Microsoft.Office.Interop.Excel.XlDeleteShiftDirection;
global using MsoTriState = Microsoft.Office.Core.MsoTriState;


global using ExcelApp = Microsoft.Office.Interop.Excel.Application;
global using ExcelWorkbooks = Microsoft.Office.Interop.Excel.Workbooks;
global using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
global using ExcelSheets = Microsoft.Office.Interop.Excel.Sheets;
global using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
global using ExcelRange = Microsoft.Office.Interop.Excel.Range;
global using ExcelShapes = Microsoft.Office.Interop.Excel.Shapes;
global using ExcelStyles = Microsoft.Office.Interop.Excel.Styles;
global using ExcelStyle = Microsoft.Office.Interop.Excel.Style;
global using ExcelFormatConditions = Microsoft.Office.Interop.Excel.FormatConditions;
global using ExcelFormatCondition = Microsoft.Office.Interop.Excel.FormatCondition;
