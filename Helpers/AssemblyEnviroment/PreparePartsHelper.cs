using SolidEdgeAdd_In.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace SolidEdgeAdd_In.Helpers.AssemblyEnviroment
{
    public class PreparePartsHelper
    {
        public static bool IsConfirmedByUser()
        {
            return DialogService.IsFileLocationsConfirmed();
        }

        public static List<string> GetParts(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            string mainDir = Path.GetDirectoryName(assembly.FullName);
            List<string> partsAndMetalSheets = new List<string>();
            AssemblyTreeWalker.PartsAndMetalSheets(assembly.Occurrences, partsAndMetalSheets, mainDir);
            return partsAndMetalSheets;
        }

        public static (int listCount, int proccesssedFileCount) PrepareAndGetStats(SolidEdgeAssembly.AssemblyDocument assembly, List<string> occurrences)
        {
            SolidEdgeFramework.SolidEdgeDocument document = null;
            SolidEdgePart.SheetMetalDocument metalSheet = null;
            SolidEdgePart.PartDocument part = null;
            SolidEdgePart.Models models = null;
            SolidEdgePart.FlatPatternModels flatPattern = null;

            SolidEdgeFramework.VariableList variableList = null;
            SolidEdgeFramework.Variables variables = null;
            SolidEdgeFramework.variable variable = null;

            int filesCount = occurrences.Count;
            int proccessedFileCount = 0;

            foreach (var occurrence in occurrences)
            {
                bool isOpen = false;

                bool hasTypeProcessed = false;
                bool hasVarProcessed = false;
                try
                {
                    document = CoreUtils.GetOpenDocument(assembly.Application, occurrence); isOpen = true;

                    if (document is SolidEdgePart.PartDocument pDoc)
                    {
                        part = pDoc; models = part.Models; flatPattern = part.FlatPatternModels;
                        variables = (SolidEdgeFramework.Variables)part.Variables;
                    }
                    else if (document is SolidEdgePart.SheetMetalDocument msDoc)
                    {
                        metalSheet = msDoc; models = metalSheet.Models; flatPattern = metalSheet.FlatPatternModels;
                        variables = (SolidEdgeFramework.Variables)metalSheet.Variables;
                    }

                    int modelsCount = models.Count; int flatPatternCount = flatPattern.Count;

                    // Ustawinie typu:
                    if (flatPatternCount > 0) { PropertyProvider.SetType(document, "B"); PropertyProvider.SetSheetMaterial(document); }
                    else { PropertyProvider.SetType(document, "C"); }
                    hasTypeProcessed = true;

                    // Ustawienie zmiennych:
                    if (variables != null && modelsCount > 0)
                    {
                        int targetExposeValue = (flatPatternCount > 0) ? 1 : 0;
                        variableList = (SolidEdgeFramework.VariableList)variables.Query("*", null, null, false);
                        for (int i = 1; i <= variableList.Count; i++)
                        {
                            try
                            {
                                variable = (SolidEdgeFramework.variable)variableList.Item(i);
                                if (variable.Expose != targetExposeValue) variable.Expose = targetExposeValue;
                            }
                            catch { continue; }
                            finally { CoreUtils.ReleaseCom(ref variable); }
                        }
                    }
                    hasVarProcessed = true;
                    if (hasTypeProcessed && hasVarProcessed) proccessedFileCount++;
                }
                catch { continue; }
                finally
                {
                    if (isOpen) document?.Close(true);
                    CoreUtils.ReleaseCom(ref variable);
                    CoreUtils.ReleaseCom(ref variableList);
                    CoreUtils.ReleaseCom(ref variables);
                    CoreUtils.ReleaseCom(ref flatPattern);
                    CoreUtils.ReleaseCom(ref models);
                    CoreUtils.ReleaseCom(ref metalSheet);
                    CoreUtils.ReleaseCom(ref part);
                    CoreUtils.ReleaseCom(ref document);
                }
            }
            return (filesCount, proccessedFileCount);
        }

        public static void Report(int filesCount, int proccessedFileCount)
        {
            if (filesCount > 0 && proccessedFileCount > 0 && filesCount == proccessedFileCount) MessageBox.Show($"Wszystkie pliki zostały przygotowane.");
            else if (filesCount > 0 && proccessedFileCount > 0 && filesCount != proccessedFileCount) MessageBox.Show($"Istnieją pliki, które nie zostały przygotowane.");
            else MessageBox.Show($"Kapitulacja.");
        }
    }
}
