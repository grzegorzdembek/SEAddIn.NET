using SolidEdgeAssembly;
using SolidEdgePart;

namespace SolidEdgeAdd_In.Utils
{
    public class AssemblyTreeWalker
    {
        public static void OccurrencesForExportDxfs(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> occurrences, DxfExportLogger logger)
        {
            int count = assemblyOccurrences.Count;

            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument doc = null;
                SeAssembly subAssembly = null;
                try
                {
                    occurrence = (SeOccurrence)assemblyOccurrences.Item(i);

                    if (occurrence.IncludeInBom == false)
                    {
                        logger.LogSkip("Nieznany Plik", "Plik wykluczony z zestawienia");
                        continue;
                    }

                    doc = (SeDocument)occurrence.OccurrenceDocument;

                    string path = null;

                    try
                    {
                        path = doc.FullName;
                    }
                    catch
                    {
                        logger.LogSkip("Nieznany Plik", "Brak dostępu do pliku");
                        continue;
                    }

                    string name = System.IO.Path.GetFileNameWithoutExtension(path);

                    if (string.IsNullOrEmpty(path))
                    {
                        logger.LogSkip(name, "Plik nie istnieje");
                        continue;
                    }

                    if (doc is SePart || doc is SeSheetMetal)
                    {            
                        using var properties = new PropertyProvider(doc);

                        if (!properties.IsTypeB)
                        {
                            logger.LogSkip(name, "Brak Typu B");
                            continue;
                        }

                        if (!properties.HasMaterial)
                        {
                            logger.LogSkip(name, "Brak zapisanego materiału");
                            continue;
                        }

                        if (!properties.HasThickness)
                        {
                            logger.LogSkip(name, "Brak grubości blachy");
                            continue;
                        }

                        if (!properties.IsStatusAvailable)
                        {
                            logger.LogSkip(name, "Status pliku jest niedostępny");
                            continue;
                        }

                        if (properties.HasDxfDate)
                        {
                            logger.LogSkip(name, "Plik ma właściwość Dxf");
                            continue;
                        }

                        if (!occurrences.ContainsKey(path))
                        {
                            occurrences[path] = new FileData
                            {
                                OccurrenceCount = 1,
                                Material = properties.Material,
                                Thickness = properties.Thickness,
                                Name = name,
                                SizeX = properties.SizeX,
                                SizeY = properties.SizeY
                            };
                        }
                        else
                        {
                            occurrences[path].OccurrenceCount++;
                        }
                    }
                    else if (doc is SeAssembly asmDoc)
                    {
                        subAssembly = asmDoc;
                        bool isTypeA = false;

                        using var properties = new PropertyProvider(doc);
                        isTypeA = properties.IsTypeA;
                     
                        if (!isTypeA)
                        {
                            logger.LogSkip(name, "Złożenie pominęto (Nie jest Typem A)");
                            continue;
                        }
                        OccurrencesForExportDxfs(subAssembly.Occurrences, occurrences, logger);
                    }
                }
                catch (Exception ex)
                { 
                    logger.LogError("Nieznany obiekt", $"Błąd podczas skanowania drzewa złożenia: {ex.Message}"); 
                    continue; 
                }
                finally
                {
                    CoreUtils.ReleaseCom(ref subAssembly);
                    CoreUtils.ReleaseCom(ref doc);
                    CoreUtils.ReleaseCom(ref occurrence);
                }
            }
        }

        public static void OccurrencesForSetCount(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> occurrences)
        {
            int count = assemblyOccurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument doc = null;
                SeAssembly subAssembly = null;

                try
                {
                    occurrence = (SeOccurrence)assemblyOccurrences.Item(i);
                    doc = (SeDocument)occurrence.OccurrenceDocument;

                    string path = null;

                    try
                    {
                        path = doc.FullName;
                    }
                    catch
                    {
                        continue;
                    }

                    string name = System.IO.Path.GetFileNameWithoutExtension(path);

                    if (string.IsNullOrEmpty(path))
                    {
                        continue;
                    }

                    if (doc is SePart || doc is SeSheetMetal)
                    {
                        string filePath = doc.FullName;
                        if (string.IsNullOrEmpty(filePath)) continue;

                        if (!occurrences.ContainsKey(filePath))
                        {
                            
                            using var properties = new PropertyProvider(doc);
                            occurrences[filePath] = new FileData
                            {
                                Name = name,
                                Type = properties.Type,
                                Count = properties.Count,
                                OccurrenceCount = 1
                            };
                        }
                        else
                        {
                            occurrences[filePath].OccurrenceCount++;
                        }
                    }
                    else if (doc is SeAssembly asmDoc)
                    {
                        subAssembly = asmDoc;
                        string subAsmPath = subAssembly.FullName;
                        if (string.IsNullOrEmpty(subAsmPath)) continue;

                        bool isTypeA = false;

                        if (!occurrences.ContainsKey(subAsmPath))
                        {
                            
                            using var properties = new PropertyProvider(doc);
                            occurrences[subAsmPath] = new FileData
                            {
                                Name = name,
                                Type = properties.Type,
                                Count = properties.Count,
                                OccurrenceCount = 1
                            };
                            isTypeA = properties.IsTypeA;
                        }
                        else
                        {
                            occurrences[subAsmPath].OccurrenceCount++;
                            isTypeA = occurrences[subAsmPath].Type == "A";
                        }

                        if (!isTypeA) continue;

                        OccurrencesForSetCount(subAssembly.Occurrences, occurrences);
                    }
                }
                catch { continue; }
                finally
                {
                    CoreUtils.ReleaseCom(ref subAssembly);
                    CoreUtils.ReleaseCom(ref doc);
                    CoreUtils.ReleaseCom(ref occurrence);
                }
            }
        }

        public static void OccurrencesForExportPartsList(SeOccurrences assemblyOccurrences, Dictionary<string, int> occurrences)
        {
            int count = assemblyOccurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument doc = null;
                SeAssembly subAssembly = null;

                try
                {
                    occurrence = (SeOccurrence)assemblyOccurrences.Item(i);
                    doc = (SeDocument)occurrence.OccurrenceDocument;

                    string path = doc.FullName;
                    if (string.IsNullOrEmpty(path)) continue;

                    if (doc is SePart || doc is SeSheetMetal)
                    {
                        if (!occurrences.ContainsKey(path)) occurrences[path] = 1;
                        else occurrences[path]++;
                    }
                    else if (doc is SeAssembly asmDoc)
                    {
                        subAssembly = asmDoc;
                        bool isTypeA = false;

                        using (var properties = new PropertyProvider(doc))
                        {
                            isTypeA = properties.IsTypeA;
                        }

                        if (!occurrences.ContainsKey(path)) occurrences[path] = 1;
                        else occurrences[path]++;

                        if (!isTypeA) continue;

                        OccurrencesForExportPartsList(subAssembly.Occurrences, occurrences);
                    }
                }
                catch { continue; }
                finally
                {
                    CoreUtils.ReleaseCom(ref subAssembly);
                    CoreUtils.ReleaseCom(ref doc);
                    CoreUtils.ReleaseCom(ref occurrence);
                }
            }
        }
    }
}