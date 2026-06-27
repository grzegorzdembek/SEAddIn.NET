using SolidEdgeAdd_In.Utils;
using System;
using System.Collections.Generic;
using System.IO;

namespace SolidEdgeAdd_In.Helpers.AssemblyEnviroment
{
    public class SaveDxfOfPartAndPsmHelper
    {
        public static string
            GetLocation
            (SolidEdgeAssembly.AssemblyDocument assembly)
        {
            string fileName = Path.GetFileNameWithoutExtension(assembly.FullName);
            string number = fileName.Length >= 4 ? fileName.Substring(0, 4) : fileName;
            string date = DateTime.Now.ToString("yyyy-MM-dd");

            string packagesFolder = Path.Combine(Path.GetDirectoryName(assembly.FullName), "Paczki");
            string subFolder = Path.Combine(packagesFolder, $"{number}_{date}");
            if (!Directory.Exists(subFolder)) Directory.CreateDirectory(subFolder);

            return subFolder;
        }

        public static Dictionary<string, int> GetPartsAndMetalSheets(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            Dictionary<string, int> partsAndMetalSheets = new Dictionary<string, int>();
            AssemblyTreeWalker.PartsAndMetalSheets(assembly.Occurrences, partsAndMetalSheets);
            return partsAndMetalSheets;
        }

        public static List<string> SaveAndGetDxfs(SolidEdgeAssembly.AssemblyDocument assembly, Dictionary<string, int> occurrences, string location)
        {
            SolidEdgeFramework.SolidEdgeDocument document = null;
            SolidEdgePart.SheetMetalDocument metalSheet = null;
            SolidEdgePart.PartDocument part = null;
            SolidEdgePart.Models models = null;
            SolidEdgePart.FlatPatternModels flatPatterns = null;

            List<string> exportedDxf = new List<string>();
            foreach (var occurrence in occurrences)
            {
                bool isSaved = false;
                bool isOpen = false;
                try
                {
                    if (!PropertyProvider.HasThickness(occurrence.Key)) continue;
                    if (!PropertyProvider.HasMaterial(occurrence.Key)) continue;
                    if (!PropertyProvider.IsStatusAvailable(occurrence.Key)) continue;

                    string dxfFilePath = RaportGenerationUtils.GetDxfPath(location, occurrence.Key, occurrences);
                    document = CoreUtils.GetOpenDocument(assembly.Application, occurrence.Key); isOpen = true;

                    if (document is SolidEdgePart.PartDocument pDoc)
                    {
                        part = (SolidEdgePart.PartDocument)pDoc;
                        models = part.Models;
                        flatPatterns = part.FlatPatternModels;
                        if (flatPatterns.Count == 0 || models.Count == 0) continue;
                    }
                    else if (document is SolidEdgePart.SheetMetalDocument msDoc)
                    {
                        metalSheet = (SolidEdgePart.SheetMetalDocument)msDoc;
                        models = metalSheet.Models;
                        flatPatterns = metalSheet.FlatPatternModels;
                        if (flatPatterns.Count == 0 || models.Count == 0) continue;
                    }

                    models.SaveAsFlatDXFEx(dxfFilePath, null, null, null, true); isSaved = true; exportedDxf.Add(dxfFilePath);
                    document.Close(true); isOpen = false;
                    if (!isOpen && isSaved) PropertyProvider.SetDxfDate(occurrence.Key);
                }
                catch { continue; }
                finally
                {
                    if (isOpen) document?.Close(true);
                    CoreUtils.ReleaseCom(ref flatPatterns);
                    CoreUtils.ReleaseCom(ref models);
                    CoreUtils.ReleaseCom(ref metalSheet);
                    CoreUtils.ReleaseCom(ref part);
                    CoreUtils.ReleaseCom(ref document);
                }
            }
            return exportedDxf;
        }

        public static void CopyDxfs(SolidEdgeAssembly.AssemblyDocument assembly, List<string> dxfs)
        {
            string mainFolder = Path.GetDirectoryName(assembly.FullName);

            foreach (string dxf in dxfs)
            {
                if (File.Exists(dxf))
                {
                    string name = Path.GetFileName(dxf);
                    string newPath = Path.Combine(mainFolder, name);
                    File.Copy(dxf, newPath, true);
                }
            }
        }
    }
}