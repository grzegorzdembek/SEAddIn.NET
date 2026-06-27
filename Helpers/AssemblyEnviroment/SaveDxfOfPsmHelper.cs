using SolidEdgeAdd_In.Utils;
using System;
using System.Collections.Generic;
using System.IO;

namespace SolidEdgeAdd_In.Helpers.AssemblyEnviroment
{
    public class SaveDxfOfPsmHelper
    {
        public static string GetLocation(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            string fileName = Path.GetFileNameWithoutExtension(assembly.FullName);
            string number = fileName.Length >= 4 ? fileName.Substring(0, 4) : fileName;
            string date = DateTime.Now.ToString("yyyy-MM-dd");

            string packagesFolder = Path.Combine(Path.GetDirectoryName(assembly.FullName), "Paczki");
            string dailySubFolder = Path.Combine(packagesFolder, $"{number}_{date}");
            if (!Directory.Exists(dailySubFolder)) Directory.CreateDirectory(dailySubFolder);

            return dailySubFolder;
        }

        public static Dictionary<string, int> GetMetalSheets(SolidEdgeAssembly.AssemblyDocument assembly)
        {
            Dictionary<string, int> sheetsMetal = new Dictionary<string, int>();
            AssemblyTreeWalker.MetalSheets(assembly.Occurrences, sheetsMetal);
            return sheetsMetal;
        }

        public static List<string> SaveAndGetDxfs(SolidEdgeAssembly.AssemblyDocument assembly, Dictionary<string, int> occurrences, string location)
        {
            SolidEdgeFramework.SolidEdgeDocument document = null;
            SolidEdgePart.SheetMetalDocument metalSheet = null;
            SolidEdgePart.Models models = null;
            SolidEdgePart.FlatPatternModels flatPatterns = null;

            bool isOpen = false;
            List<string> exportedDxf = new List<string>();
            foreach (var occurrence in occurrences)
            {
                bool isSaved = false;
                try
                {
                    if (!PropertyProvider.HasMaterial(occurrence.Key)) continue;
                    if (!PropertyProvider.HasThickness(occurrence.Key)) continue;
                    if (!PropertyProvider.IsStatusAvailable(occurrence.Key)) continue;

                    string dxfFilePath = RaportGenerationUtils.GetDxfPath(location, occurrence.Key, occurrences);
                    document = CoreUtils.GetOpenDocument(assembly.Application, occurrence.Key); isOpen = true;

                    metalSheet = (SolidEdgePart.SheetMetalDocument)document;
                    models = metalSheet.Models;
                    flatPatterns = metalSheet.FlatPatternModels;
                    if (flatPatterns.Count == 0 || models.Count == 0) continue;

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