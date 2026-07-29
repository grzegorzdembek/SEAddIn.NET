namespace SolidEdgeAdd_In.Utils
{
    public class AssemblyTreeWalker
    {
        public static void BuildDataForExportDxfs(SeOccurrences occurrences, Dictionary<string, FileData> data, Logger logger)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrences subOccurrences = null; 
                SeOccurrence occurrence = null; 
                SeDocument document = null;
              
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i); 
                    if (occurrence.IncludeInBom == false) 
                    {
                        logger.LogSkip("Nieznany Plik", "Plik wykluczony z zestawienia"); 
                        continue; 
                    }

                    document = (SeDocument)occurrence.OccurrenceDocument;

                    string filePath = null; 
                    try 
                    { 
                        filePath = document.FullName; 
                    } 
                    catch 
                    { 
                        logger.LogSkip("Nieznany Plik", "Brak dostępu do pliku"); 
                        continue; 
                    }

                    if (string.IsNullOrEmpty(filePath)) 
                    { 
                        logger.LogSkip("Niezapisany Plik", "Plik nie istnieje na dysku"); 
                        continue; 
                    }

                    string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);

                    using var properties = new PropertyProvider(document);

                    if (document is SePart || document is SeSheetMetal)
                    {
                        /*
                         * these data are required for dxf path, avaibale status required by user
                         * if any of them is missing, we cannot generate dxf file
                        */ 

                        string type = properties.Type; 
                        if (!(type == "B"))
                        { 
                            logger.LogSkip(fileName, "Brak Typu"); 
                            continue; 
                        }

                        string material = properties.Material;
                        if (string.IsNullOrEmpty(material)) 
                        { 
                            logger.LogSkip(fileName, "Brak materiału blachy"); 
                            continue; 
                        }

                        string thickness = properties.Thickness;
                        if (string.IsNullOrEmpty(thickness)) 
                        { 
                            logger.LogSkip(fileName, "Brak grubości blachy"); 
                            continue; 
                        }

                        int status = properties.Status;
                        if (!(status == 0)) 
                        { 
                            logger.LogSkip(fileName, "Status inny niż dostępny"); 
                            continue; 
                        }

                        /*
                         * these data are only for excel summary, not for dxf path
                         * if any of them is missing, we can still generate dxf file, just the excel summary will have empty fields
                        */

                        string sizeX = properties.SizeX;
                        string sizeY = properties.SizeY;

                        string dxfDate = properties.DxfDate;
                        string title = properties.Title;

                        
                        if (!data.ContainsKey(filePath)) 
                        { 
                            data[filePath] = new FileData 
                            { 
                                OccurrenceCount = 1,
                                Material = material, 
                                Thickness = thickness, 
                                Name = fileName, 
                                SizeX = sizeX, 
                                SizeY = sizeY, 
                                DxfDate = dxfDate, 
                                Title = title
                            }; 
                        }
                        else 
                        { 
                            data[filePath].OccurrenceCount++; 
                        }
                    }
                    else if (document is SeAssembly asmDoc)
                    {
                        bool isTypeA = false; 
                        isTypeA = properties.IsTypeA;

                        if (!isTypeA) 
                        { 
                            logger.LogSkip(fileName, "Złożenie pominęto"); 
                            continue; 
                        }

                        try 
                        {
                            subOccurrences = asmDoc.Occurrences;
                            BuildDataForExportDxfs(subOccurrences, data, logger); 
                        } 
                        finally
                        { 
                            CoreUtils.ReleaseCom(ref subOccurrences); 
                        }
                    }
                }
                catch (Exception ex) 
                { 
                    logger.LogError("Nieznany obiekt", $"Błąd podczas skanowania drzewa złożenia: {ex.Message}"); 
                    continue; 
                }
                finally 
                { 
                    CoreUtils.ReleaseCom(ref document);
                    CoreUtils.ReleaseCom(ref occurrence); 
                }
            }
        }

        public static void BuildDataForExportPartsList(SeOccurrences assemblyOccurrences, Dictionary<string, int> data)
        {
            int count = assemblyOccurrences.Count;

            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null; SeDocument doc = null;

                try
                {
                    occurrence = (SeOccurrence)assemblyOccurrences.Item(i); doc = (SeDocument)occurrence.OccurrenceDocument;

                    string path = null; try { path = doc.FullName; } catch { continue; }

                    if (string.IsNullOrEmpty(path)) continue;

                    if (doc is SePart || doc is SeSheetMetal)
                    {
                        if (!data.ContainsKey(path)) data[path] = 1; else data[path]++;
                    }
                    else if (doc is SeAssembly asmDoc)
                    {
                        bool isTypeA = false;

                        using (var properties = new PropertyProvider(doc)) { isTypeA = properties.IsTypeA; }

                        if (!data.ContainsKey(path)) data[path] = 1; else data[path]++;

                        if (!isTypeA) continue;

                        SeOccurrences subOccurrences = null; try { subOccurrences = asmDoc.Occurrences; BuildDataForExportPartsList(subOccurrences, data); } finally { CoreUtils.ReleaseCom(ref subOccurrences); }
                    }
                }
                catch { continue; }
                finally { CoreUtils.ReleaseCom(ref doc); CoreUtils.ReleaseCom(ref occurrence); }
            }
        }

        public static void BuildDataForSetCount(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data)
        {
            int count = assemblyOccurrences.Count;

            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null; SeDocument doc = null;

                try
                {
                    occurrence = (SeOccurrence)assemblyOccurrences.Item(i); doc = (SeDocument)occurrence.OccurrenceDocument;

                    string path = null; try { path = doc.FullName; } catch { continue; }

                    if (string.IsNullOrEmpty(path)) { continue; }

                    string name = System.IO.Path.GetFileNameWithoutExtension(path);

                    using var properties = new PropertyProvider(doc);

                    if (doc is SePart || doc is SeSheetMetal)
                    {
                        if (!data.ContainsKey(path)) { data[path] = new FileData { Name = name, Type = properties.Type, Count = properties.Count, OccurrenceCount = 1 }; }
                        else { data[path].OccurrenceCount++; }
                    }
                    else if (doc is SeAssembly asmDoc)
                    {
                        string subAsmPath = asmDoc.FullName; if (string.IsNullOrEmpty(subAsmPath)) continue;

                        bool isTypeA = false;

                        if (!data.ContainsKey(subAsmPath)) { data[subAsmPath] = new FileData { Name = name, Type = properties.Type, Count = properties.Count, OccurrenceCount = 1 }; isTypeA = properties.IsTypeA; }
                        else { data[subAsmPath].OccurrenceCount++; isTypeA = data[subAsmPath].Type == Constants.PartTypes.Assembly; }

                        if (!isTypeA) continue;

                        SeOccurrences subOccurrences = null; try { subOccurrences = asmDoc.Occurrences; BuildDataForSetCount(subOccurrences, data); } finally { CoreUtils.ReleaseCom(ref subOccurrences); }
                    }
                }
                catch { continue; }
                finally { CoreUtils.ReleaseCom(ref doc); CoreUtils.ReleaseCom(ref occurrence); }
            }
        }

        public static void ApplyCounts(SeOccurrences occurrences, Dictionary<string, FileData> data, int multiplier, HashSet<string> processed)
        {
            int count = occurrences.Count;

            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null; SeDocument doc = null;

                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i); doc = (SeDocument)occurrence.OccurrenceDocument;

                    string path = null; try { path = doc.FullName; } catch { continue; }

                    if (string.IsNullOrEmpty(path)) continue;

                    if (data.ContainsKey(path) && !processed.Contains(path))
                    {
                        using var properties = new PropertyProvider(doc); properties.Count = data[path].OccurrenceCount * multiplier; processed.Add(path);
                    }

                    if (doc is SeAssembly asmDoc)
                    {
                        SeOccurrences subOccurrences = null; try { subOccurrences = asmDoc.Occurrences; ApplyCounts(subOccurrences, data, multiplier, processed); } finally { CoreUtils.ReleaseCom(ref subOccurrences); }
                    }
                }
                catch { continue; }
                finally { CoreUtils.ReleaseCom(ref doc); CoreUtils.ReleaseCom(ref occurrence); }
            }
        }
    }
}