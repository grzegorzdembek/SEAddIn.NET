namespace SolidEdgeAdd_In.Utils
{
    public class DataUtils
    {
        
        public static void BuildDataForExportOccurrenceList(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data, List<string> types)
        {
            ProcessExportOccurrenceList(assemblyOccurrences, data, types, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }

        private static void ProcessExportOccurrenceList(SeOccurrences occurrences, Dictionary<string, FileData> data, List<string> types, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    document = (SeDocument)occurrence.OccurrenceDocument;
              
                    string filePath = null;
                    try { filePath = document.FullName; } catch { continue; }
                    if (string.IsNullOrEmpty(filePath)) { continue; }

                    bool isAssembly = document is SeAssembly;
   
                    if (data.ContainsKey(filePath))
                    {
                        data[filePath].OccurrenceCount++;
                    }
                    else if (!isAssembly || !assemblyCache.ContainsKey(filePath))
                    {
                        using var properties = new PropertyUtils(document);
                        string type = properties.Type;

                        if (types.Contains(type))
                        {
                            data[filePath] = new FileData
                            {
                                FileName = Path.GetFileNameWithoutExtension(filePath),
                                Title = properties.TitleEng ?? properties.TitlePl,
                                Type = type,
                                OccurrenceCount = 1
                            };
                        }

                        if (isAssembly) { assemblyCache[filePath] = properties.IsTypeA; }
                    }

                    
                    if (isAssembly && assemblyCache.TryGetValue(filePath, out bool isTypeA) && isTypeA)
                    {
                        SeOccurrences subOccurrences = null;
                        try { subOccurrences = ((SeAssembly)document).Occurrences; ProcessExportOccurrenceList(subOccurrences, data, types, assemblyCache); }
                        finally { Helpers.ReleaseCom(ref subOccurrences); }
                    }
                }
                catch { continue; }
                finally { Helpers.ReleaseCom(ref document); Helpers.ReleaseCom(ref occurrence); }
            }
        }

       
        public static void BuildDataForSetCount(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data)
        {
            ProcessSetCount(assemblyOccurrences, data, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }

        private static void ProcessSetCount(SeOccurrences occurrences, Dictionary<string, FileData> data, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    document = (SeDocument)occurrence.OccurrenceDocument;

                    
                    string filePath = null;
                    try { filePath = document.FullName; } catch { continue; }
                    if (string.IsNullOrEmpty(filePath)) { continue; }

                    bool isAssembly = document is SeAssembly;

                    
                    if (data.ContainsKey(filePath))
                    {
                        data[filePath].OccurrenceCount++;
                    }
                    else if (!isAssembly || !assemblyCache.ContainsKey(filePath))
                    {
                        using var properties = new PropertyUtils(document);
                        string type = properties.Type;

                        if (type == Constants.PartTypes.Assembly || type == Constants.PartTypes.SheetMetal || type == Constants.PartTypes.Part || type == Constants.PartTypes.Steelmaking)
                        {
                            data[filePath] = new FileData
                            {
                                FileName = Path.GetFileNameWithoutExtension(filePath),
                                Type = type,
                                Count = properties.Count,
                                OccurrenceCount = 1
                            };
                        }

                        if (isAssembly) { assemblyCache[filePath] = properties.IsTypeA; }
                    }

                    
                    if (isAssembly && assemblyCache.TryGetValue(filePath, out bool isTypeA) && isTypeA)
                    {
                        SeOccurrences subOccurrences = null;
                        try { subOccurrences = ((SeAssembly)document).Occurrences; ProcessSetCount(subOccurrences, data, assemblyCache); }
                        finally { Helpers.ReleaseCom(ref subOccurrences); }
                    }
                }
                catch { continue; }
                finally { Helpers.ReleaseCom(ref document); Helpers.ReleaseCom(ref occurrence); }
            }
        }

       
        public static void BuildDataForExportDxfs(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data, Logger logger)
        {
            ProcessExportDxfs(assemblyOccurrences, data, logger, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }

        private static void ProcessExportDxfs(SeOccurrences occurrences, Dictionary<string, FileData> data, Logger logger, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    if (occurrence.IncludeInBom == false) { logger.LogSkip("Unknown File", "File excluded from BOM"); continue; }

                    document = (SeDocument)occurrence.OccurrenceDocument;

                    
                    string filePath = null;
                    try { filePath = document.FullName; } catch { logger.LogSkip("Unknown File", "File unavailable"); continue; }
                    if (string.IsNullOrEmpty(filePath)) { logger.LogSkip("Unsaved File", "File does not exist on disk"); continue; }

                    bool isAssembly = document is SeAssembly;

                    
                    if (data.ContainsKey(filePath))
                    {
                        data[filePath].OccurrenceCount++;
                    }
                    else if (!isAssembly || !assemblyCache.ContainsKey(filePath))
                    {
                        using var properties = new PropertyUtils(document);
                        string fileName = Path.GetFileNameWithoutExtension(filePath);

                        if (document is SePart || document is SeSheetMetal)
                        {
                            string type = properties.Type;
                            if (type != Constants.PartTypes.SheetMetal) { logger.LogSkip(fileName, $"Missing Property {Constants.PartTypes.SheetMetal}"); }
                            else
                            {
                                string material = properties.Material;
                                string thickness = properties.Thickness;
                                int status = properties.Status;

                                if (string.IsNullOrEmpty(material)) { logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Material}"); }
                                else if (string.IsNullOrEmpty(thickness)) { logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Thickness}"); }
                                else if (status != 0) { logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Status} set to available"); }
                                else
                                {
                                    data[filePath] = new FileData
                                    {
                                        OccurrenceCount = 1,
                                        Material = material,
                                        Thickness = thickness,
                                        FileName = fileName,
                                        SizeX = properties.SizeX,
                                        SizeY = properties.SizeY,
                                        DxfDate = properties.DxfDate,
                                        Title = properties.TitleEng ?? properties.TitlePl,
                                        Color = properties.Color,
                                        Finish = properties.Finish
                                    };
                                }
                            }
                        }
                        else if (isAssembly)
                        {
                            bool isTypeA = properties.IsTypeA;
                            assemblyCache[filePath] = isTypeA;
                            if (!isTypeA) { logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Type} set to Assembly"); }
                        }
                    }

                    
                    if (isAssembly && assemblyCache.TryGetValue(filePath, out bool typeA) && typeA)
                    {
                        SeOccurrences subOccurrences = null;
                        try { subOccurrences = ((SeAssembly)document).Occurrences; ProcessExportDxfs(subOccurrences, data, logger, assemblyCache); }
                        finally { Helpers.ReleaseCom(ref subOccurrences); }
                    }
                }
                catch (Exception ex) { logger.LogError("Unknown File", $"Error scanning assembly tree: {ex.Message}"); continue; }
                finally { Helpers.ReleaseCom(ref document); Helpers.ReleaseCom(ref occurrence); }
            }
        }

        public static void BuildDataForExportPartsList(SeOccurrences assemblyOccurrences, Dictionary<string, int> data)
        {
            ProcessExportPartsList(assemblyOccurrences, data, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }

        private static void ProcessExportPartsList(SeOccurrences occurrences, Dictionary<string, int> data, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    document = (SeDocument)occurrence.OccurrenceDocument;

                    string filePath = null;
                    try { filePath = document.FullName; } catch { continue; }
                    if (string.IsNullOrEmpty(filePath)) { continue; }

                    bool isAssembly = document is SeAssembly;

                    if (data.ContainsKey(filePath))
                    {
                        data[filePath]++;
                    }
                    else
                    {
                        data[filePath] = 1;
                        if (isAssembly && !assemblyCache.ContainsKey(filePath))
                        {
                            using var properties = new PropertyUtils(document);
                            assemblyCache[filePath] = properties.IsTypeA;
                        }
                    }

                    if (isAssembly && assemblyCache.TryGetValue(filePath, out bool isTypeA) && isTypeA)
                    {
                        SeOccurrences subOccurrences = null;
                        try { subOccurrences = ((SeAssembly)document).Occurrences; ProcessExportPartsList(subOccurrences, data, assemblyCache); }
                        finally { Helpers.ReleaseCom(ref subOccurrences); }
                    }
                }
                catch { continue; }
                finally { Helpers.ReleaseCom(ref document); Helpers.ReleaseCom(ref occurrence); }
            }
        }

        public static void BuildDataForOrganiseDrawings(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data)
        {
            ProcessOrganiseDrawings(assemblyOccurrences, data, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }

        private static void ProcessOrganiseDrawings(SeOccurrences occurrences, Dictionary<string, FileData> data, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    document = (SeDocument)occurrence.OccurrenceDocument;

                    string filePath = null;
                    try { filePath = document.FullName; } catch { continue; }
                    if (string.IsNullOrEmpty(filePath)) { continue; }

                    bool isAssembly = document is SeAssembly;

                    if (data.ContainsKey(filePath))
                    {
                        
                    }
                    else if (!isAssembly || !assemblyCache.ContainsKey(filePath))
                    {
                        using var properties = new PropertyUtils(document);
                        data[filePath] = new FileData
                        {
                            FileName = Path.GetFileNameWithoutExtension(filePath),
                            Type = properties.Type
                        };

                        if (isAssembly) { assemblyCache[filePath] = properties.IsTypeA; }
                    }

                   
                    if (isAssembly && assemblyCache.TryGetValue(filePath, out bool isTypeA) && isTypeA)
                    {
                        SeOccurrences subOccurrences = null;
                        try { subOccurrences = ((SeAssembly)document).Occurrences; ProcessOrganiseDrawings(subOccurrences, data, assemblyCache); }
                        finally { Helpers.ReleaseCom(ref subOccurrences); }
                    }
                }
                catch { continue; }
                finally { Helpers.ReleaseCom(ref document); Helpers.ReleaseCom(ref occurrence); }
            }
        }

        
        public static void ApplyCounts(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data, int multiplier, HashSet<string> processed)
        {
            ProcessApplyCounts(assemblyOccurrences, data, multiplier, processed, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }

        private static void ProcessApplyCounts(SeOccurrences occurrences, Dictionary<string, FileData> data, int multiplier, HashSet<string> processed, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    document = (SeDocument)occurrence.OccurrenceDocument;


                    string filePath = null;
                    try { filePath = document.FullName; } catch { continue; }
                    if (string.IsNullOrEmpty(filePath)) { continue; }

                    bool isAssembly = document is SeAssembly;
                    bool needsUpdate = data.ContainsKey(filePath) && !processed.Contains(filePath);

                    if (!needsUpdate && !isAssembly) { continue; }

                    bool isTypeA = false;


                    if (isAssembly && assemblyCache.TryGetValue(filePath, out bool cachedTypeA))
                    {
                        isTypeA = cachedTypeA;
                        if (needsUpdate)
                        {
                            using var properties = new PropertyUtils(document);
                            properties.Count = data[filePath].OccurrenceCount * multiplier;
                            processed.Add(filePath);
                        }
                    }
                    else
                    {
                        using var properties = new PropertyUtils(document);
                        if (needsUpdate)
                        {
                            properties.Count = data[filePath].OccurrenceCount * multiplier;
                            processed.Add(filePath);
                        }
                        if (isAssembly)
                        {
                            isTypeA = properties.IsTypeA;
                            assemblyCache[filePath] = isTypeA;
                        }
                    }

                    if (isAssembly && isTypeA)
                    {
                        SeOccurrences subOccurrences = null;
                        try { subOccurrences = ((SeAssembly)document).Occurrences; ProcessApplyCounts(subOccurrences, data, multiplier, processed, assemblyCache); }
                        finally { Helpers.ReleaseCom(ref subOccurrences); }
                    }
                }
                catch { continue; }
                finally { Helpers.ReleaseCom(ref document); Helpers.ReleaseCom(ref occurrence); }
            }
        }

        public static void ClearDxfDates(SeOccurrences assemblyOccurrences, HashSet<string> processed)
        {
            ProcessClearDxfDates(assemblyOccurrences, processed, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }

        private static void ProcessClearDxfDates(SeOccurrences occurrences, HashSet<string> processed, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    if (occurrence.IncludeInBom == false) { continue; }

                    document = (SeDocument)occurrence.OccurrenceDocument;

                   
                    string filePath = null;
                    try { filePath = document.FullName; } catch { continue; }
                    if (string.IsNullOrEmpty(filePath)) { continue; }

                    bool isAssembly = document is SeAssembly;

                    if (!processed.Contains(filePath))
                    {
                        using var properties = new PropertyUtils(document);

                        if (document is SePart || document is SeSheetMetal)
                        {
                            if (properties.IsTypeB && properties.HasMaterial && properties.HasThickness && properties.IsStatusAvailable)
                            {
                                if (properties.HasDxfDate) { properties.ClearDxfDate(); }
                            }
                        }

                        processed.Add(filePath);

                        if (isAssembly) { assemblyCache[filePath] = properties.IsTypeA; }
                    }

                    if (isAssembly && assemblyCache.TryGetValue(filePath, out bool isTypeA) && isTypeA)
                    {
                        SeOccurrences subOccurrences = null;
                        try { subOccurrences = ((SeAssembly)document).Occurrences; ProcessClearDxfDates(subOccurrences, processed, assemblyCache); }
                        finally { Helpers.ReleaseCom(ref subOccurrences); }
                    }
                }
                catch { continue; }
                finally { Helpers.ReleaseCom(ref document); Helpers.ReleaseCom(ref occurrence); }
            }
        }
    }
}