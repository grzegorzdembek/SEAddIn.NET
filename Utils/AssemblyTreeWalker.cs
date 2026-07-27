namespace SolidEdgeAdd_In.Utils
{
    public class AssemblyTreeWalker
    {
        public static void BuildDataForExportDxfs(SeOccurrences occurrences, Dictionary<string, FileData> data, Logger logger)
        {
            int count = occurrences.Count;

            for (int i = 1; i <= count; i++)
            {
                SeOccurrences subOccurrences = null; SeOccurrence occurrence = null; SeDocument doc = null;

                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i); if (occurrence.IncludeInBom == false) { logger.LogSkip("Nieznany Plik", "Plik wykluczony z zestawienia"); continue; }

                    doc = (SeDocument)occurrence.OccurrenceDocument;

                    string path = null; try { path = doc.FullName; } catch { logger.LogSkip("Nieznany Plik", "Brak dostępu do pliku"); continue; }

                    if (string.IsNullOrEmpty(path)) { logger.LogSkip("Niezapisany Plik", "Plik nie istnieje na dysku"); continue; }

                    string name = System.IO.Path.GetFileNameWithoutExtension(path);

                    using var properties = new PropertyProvider(doc);

                    if (doc is SePart || doc is SeSheetMetal)
                    {
                       

                        if (!properties.IsTypeB) { logger.LogSkip(name, "Brak Typu"); continue; }
                        if (!properties.HasMaterial) { logger.LogSkip(name, "Brak materiału blachy"); continue; }
                        if (!properties.HasThickness) { logger.LogSkip(name, "Brak grubości blachy"); continue; }
                        if (!properties.IsStatusAvailable) { logger.LogSkip(name, "Status inny niż dostępny"); continue; }

                        if (!data.ContainsKey(path)) { data[path] = new FileData { OccurrenceCount = 1, Material = properties.Material, Thickness = properties.Thickness, Name = name, SizeX = properties.SizeX, SizeY = properties.SizeY, DxfDate = properties.DxfDate }; }
                        else { data[path].OccurrenceCount++; }
                    }
                    else if (doc is SeAssembly asmDoc)
                    {
                        bool isTypeA = false; isTypeA = properties.IsTypeA;

                        if (!isTypeA) { logger.LogSkip(name, "Złożenie pominęto"); continue; }

                        try { subOccurrences = asmDoc.Occurrences; BuildDataForExportDxfs(subOccurrences, data, logger); } finally { CoreUtils.ReleaseCom(ref subOccurrences); }
                    }
                }
                catch (Exception ex) { logger.LogError("Nieznany obiekt", $"Błąd podczas skanowania drzewa złożenia: {ex.Message}"); continue; }
                finally { CoreUtils.ReleaseCom(ref doc); CoreUtils.ReleaseCom(ref occurrence); }
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