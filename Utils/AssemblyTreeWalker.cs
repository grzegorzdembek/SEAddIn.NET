using SolidEdgeAssembly;
using System;
using System.Collections.Generic;
using System.IO;

namespace SolidEdgeAdd_In.Utils
{
    public class AssemblyTreeWalker
    {
        public static void PartsAndMetalSheets(SolidEdgeAssembly.Occurrences occurrences, List<string> partsAndMetalSheets, string mainDir)
        {
            mainDir = mainDir.TrimEnd('\\', '/');

            for (int i = 1; i <= occurrences.Count; i++)
            {
                SolidEdgeAssembly.Occurrence occurrence = null;
                SolidEdgeFramework.SolidEdgeDocument doc = null;
                SolidEdgePart.SheetMetalDocument sheetMetal = null;
                SolidEdgePart.PartDocument part = null;
                SolidEdgeAssembly.AssemblyDocument subAssembly = null;

                try
                {
                    occurrence = occurrences.Item(i);

                    doc = (SolidEdgeFramework.SolidEdgeDocument)occurrence.OccurrenceDocument;

                    if (doc is SolidEdgePart.PartDocument pDoc)
                    {
                        part = pDoc;
                        string partPath = part.FullName;
                        string partDir = System.IO.Path.GetDirectoryName(partPath);
                        if (string.IsNullOrEmpty(partDir)) continue;
                        partDir = partDir.TrimEnd('\\', '/');
                        if (string.Equals(mainDir, partDir, StringComparison.OrdinalIgnoreCase)) partsAndMetalSheets.Add(partPath);
                    }
                    else if (doc is SolidEdgePart.SheetMetalDocument smDoc)
                    {
                        sheetMetal = smDoc;
                        string metalSheetPath = sheetMetal.FullName;
                        string metalSheetDir = System.IO.Path.GetDirectoryName(metalSheetPath);
                        metalSheetDir = metalSheetDir.TrimEnd('\\', '/');
                        if (string.Equals(mainDir, metalSheetDir, StringComparison.OrdinalIgnoreCase)) partsAndMetalSheets.Add(metalSheetPath);
                    }
                    else if (doc is SolidEdgeAssembly.AssemblyDocument asmDoc)
                    {
                        subAssembly = asmDoc;
                        string subAsmPath = subAssembly.FullName;
                        if (!PropertyProvider.IsTypeA(subAsmPath)) continue;
                        PartsAndMetalSheets(subAssembly.Occurrences, partsAndMetalSheets, mainDir);
                    }
                }
                catch { continue; }
                finally
                {
                    CoreUtils.ReleaseCom(ref subAssembly);
                    CoreUtils.ReleaseCom(ref sheetMetal);
                    CoreUtils.ReleaseCom(ref part);
                    CoreUtils.ReleaseCom(ref doc);
                    CoreUtils.ReleaseCom(ref occurrence);
                }
            }
        }

        public static void MetalSheets(SolidEdgeAssembly.Occurrences occurrences, Dictionary<string, int> metalSheets)
        {
            for (int i = 1; i <= occurrences.Count; i++)
            {
                SolidEdgeAssembly.Occurrence occurrence = null;
                SolidEdgeFramework.SolidEdgeDocument doc = null;
                SolidEdgePart.SheetMetalDocument sheetMetal = null;
                SolidEdgeAssembly.AssemblyDocument subAssembly = null;              

                try
                {
                    occurrence = occurrences.Item(i);
                    if (occurrence.IncludeInBom == false) { continue; }

                    doc = (SolidEdgeFramework.SolidEdgeDocument)occurrence.OccurrenceDocument;
                    if (doc is SolidEdgePart.SheetMetalDocument smDoc)
                    {
                        sheetMetal = smDoc;
                        string metalSheetPath = sheetMetal.FullName;

                        if (!PropertyProvider.IsTypeB(metalSheetPath)) continue;

                        if (metalSheets.ContainsKey(metalSheetPath)) metalSheets[metalSheetPath]++;
                        else metalSheets[metalSheetPath] = 1;
                    }
                    else if (doc is SolidEdgeAssembly.AssemblyDocument asmDoc)
                    {
                        subAssembly = asmDoc;
                        string subAsmPath = subAssembly.FullName;

                        if (!PropertyProvider.IsTypeA(subAsmPath)) continue;
                        MetalSheets(subAssembly.Occurrences, metalSheets);
                    }
                }
                catch { continue; }
                finally
                {
                    CoreUtils.ReleaseCom(ref sheetMetal);
                    CoreUtils.ReleaseCom(ref subAssembly);
                    CoreUtils.ReleaseCom(ref doc);
                    CoreUtils.ReleaseCom(ref occurrence);
                }
            }
        }

        public static void PartsAndMetalSheets(SolidEdgeAssembly.Occurrences occurrences, Dictionary<string, int> partsAndMetalSheets)
        {
            for (int i = 1; i <= occurrences.Count; i++)
            {
                SolidEdgeAssembly.Occurrence occurrence = null;
                SolidEdgeFramework.SolidEdgeDocument doc = null;
                SolidEdgePart.SheetMetalDocument sheetMetal = null;
                SolidEdgePart.PartDocument part = null;
                SolidEdgeAssembly.AssemblyDocument subAssembly = null;

                try
                {
                    occurrence = occurrences.Item(i);
                    if (occurrence.IncludeInBom == false) { continue; }

                    doc = (SolidEdgeFramework.SolidEdgeDocument)occurrence.OccurrenceDocument;
                    if (doc is SolidEdgePart.PartDocument pDoc)
                    {
                        part = pDoc;
                        string partPath = part.FullName;

                        if (!PropertyProvider.IsTypeB(partPath)) continue;

                        if (partsAndMetalSheets.ContainsKey(partPath)) partsAndMetalSheets[partPath]++;
                        else partsAndMetalSheets[partPath] = 1;
                    }
                    else if (doc is SolidEdgePart.SheetMetalDocument smDoc)
                    {
                        sheetMetal = smDoc;
                        string metalSheetPath = sheetMetal.FullName;

                        if (!PropertyProvider.IsTypeB(metalSheetPath)) continue;

                        if (partsAndMetalSheets.ContainsKey(metalSheetPath)) partsAndMetalSheets[metalSheetPath]++;
                        else partsAndMetalSheets[metalSheetPath] = 1;
                    }
                    else if (doc is SolidEdgeAssembly.AssemblyDocument asmDoc)
                    {
                        subAssembly = asmDoc;
                        string subAsmPath = subAssembly.FullName;

                        if (!PropertyProvider.IsTypeA(subAsmPath)) continue;
                        PartsAndMetalSheets(subAssembly.Occurrences, partsAndMetalSheets);
                    }
                }
                catch { continue; }
                finally
                {
                    CoreUtils.ReleaseCom(ref subAssembly);
                    CoreUtils.ReleaseCom(ref sheetMetal);
                    CoreUtils.ReleaseCom(ref part);
                    CoreUtils.ReleaseCom(ref doc);
                    CoreUtils.ReleaseCom(ref occurrence);
                }
            }
        }

        public static void AllOccurrences(SolidEdgeAssembly.Occurrences occurrences, Dictionary<string, int> allOccurrences)
        {
            for (int i = 1; i <= occurrences.Count; i++)
            {
                SolidEdgeAssembly.Occurrence occurrence = null;
                SolidEdgeFramework.SolidEdgeDocument doc = null;
                SolidEdgePart.PartDocument part = null;
                SolidEdgePart.SheetMetalDocument sheetMetal = null;
                SolidEdgeAssembly.AssemblyDocument subAssembly = null;
                try
                {
                    occurrence = occurrences.Item(i);
                    //if (occurrence.IncludeInBom == false) { continue; }

                    doc = (SolidEdgeFramework.SolidEdgeDocument)occurrence.OccurrenceDocument;
                    if (doc is SolidEdgePart.PartDocument pDoc)
                    {
                        part = pDoc;
                        string partPath = part.FullName;

                        if (allOccurrences.ContainsKey(partPath)) allOccurrences[partPath]++;
                        else allOccurrences[partPath] = 1;
                    }
                    else if (doc is SolidEdgePart.SheetMetalDocument smDoc)
                    {
                        sheetMetal = smDoc;
                        string sheetMetalPath = sheetMetal.FullName;

                        if (allOccurrences.ContainsKey(sheetMetalPath)) allOccurrences[sheetMetalPath]++;
                        else allOccurrences[sheetMetalPath] = 1;
                    }
                    else if (doc is SolidEdgeAssembly.AssemblyDocument asmDoc)
                    {
                        subAssembly = asmDoc;
                        string subAsmPath = subAssembly.FullName;

                        if (allOccurrences.ContainsKey(subAsmPath)) allOccurrences[subAsmPath]++;
                        else allOccurrences[subAsmPath] = 1;

                        if (!PropertyProvider.IsTypeA(subAsmPath)) continue;

                        AllOccurrences(subAssembly.Occurrences, allOccurrences);
                    }
                }
                catch { continue; }
                finally
                {
                    CoreUtils.ReleaseCom(ref subAssembly);
                    CoreUtils.ReleaseCom(ref sheetMetal);
                    CoreUtils.ReleaseCom(ref part);
                    CoreUtils.ReleaseCom(ref doc);
                    CoreUtils.ReleaseCom(ref occurrence);
                }
            }
        }
    }
}

