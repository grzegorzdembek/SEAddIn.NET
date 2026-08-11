namespace SolidEdgeAdd_In.Utils
{
    public class DataUtils
    {
        /**************************************************                  EXPORT DXFS                  ****************************************************/
        public static void BuildDataForExportDxfs(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data, Logger logger)
        {
            ProcessDataForExportDxfs(assemblyOccurrences, data, logger, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }
        private static void ProcessDataForExportDxfs(SeOccurrences occurrences, Dictionary<string, FileData> data, Logger logger, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    if (occurrence.IncludeInBom == false) 
                    { 
                        logger.LogSkip("Unknown File", "File excluded from BOM"); 
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
                        logger.LogSkip("Unknown File", "File unavailable"); 
                        continue;
                    }

                    if (string.IsNullOrEmpty(filePath)) 
                    { 
                        logger.LogSkip("Unsaved File", "File does not exist on disk"); 
                        continue; 
                    }

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
                            if (type != Constants.PartTypes.SheetMetal) 
                            { 
                                logger.LogSkip(fileName, $"Missing Property {Constants.PartTypes.SheetMetal}");
                                continue;
                            }
                            else
                            {
                                string material = properties.Material;
                                string thickness = properties.Thickness;
                                int status = properties.Status;

                                if (string.IsNullOrEmpty(material)) 
                                { 
                                    logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Material}"); 
                                    continue;
                                }

                                if (string.IsNullOrEmpty(thickness)) 
                                { 
                                    logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Thickness}");
                                    continue;
                                }

                                if (status != 0) 
                                {
                                    logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Status} set to available"); 
                                    continue;
                                }

                                data[filePath] = new FileData
                                {
                                    OccurrenceCount = 1,
                                    Material = material,
                                    Thickness = thickness,
                                    Name = fileName,
                                    SizeX = properties.SizeX,
                                    SizeY = properties.SizeY,
                                    DxfDate = properties.DxfDate,
                                    Title = properties.TitleEng ?? properties.TitlePl,
                                    Color = properties.Color,
                                    Finish = properties.Finish
                                };
                            }
                        }
                        else if (isAssembly)
                        {
                            bool isTypeA = properties.IsTypeA;
                            assemblyCache[filePath] = isTypeA;
                            if (!isTypeA) logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Type} set to Assembly"); 
                        }
                    }

                    if (isAssembly && assemblyCache.TryGetValue(filePath, out bool typeA) && typeA)
                    {
                        SeOccurrences subOccurrences = null;
                        try
                        { 
                            subOccurrences = ((SeAssembly)document).Occurrences; 
                            ProcessDataForExportDxfs(subOccurrences, data, logger, assemblyCache); 
                        }
                        finally
                        { 
                            Helpers.ReleaseCom(ref subOccurrences); 
                        }
                    }
                }
                catch (Exception ex)
                { 
                    logger.LogError("Unknown File", $"Error scanning assembly tree: {ex.Message}"); 
                    continue; 
                }
                finally 
                { 
                    Helpers.ReleaseCom(ref document); 
                    Helpers.ReleaseCom(ref occurrence); 
                }
            }
        }
        /*****************************************************************************************************************************************************/



        /***************************************             EXPORT PARTS LIST                     ***********************************************************/
        public static void BuildDataForExportPartsList(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data)
        {
            ProcessDataForExportPartsList(assemblyOccurrences, data, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }
        private static void ProcessDataForExportPartsList(SeOccurrences occurrences, Dictionary<string, FileData> data, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    if (occurrence.IncludeInBom == false) continue;
                    document = (SeDocument)occurrence.OccurrenceDocument;

                    string filePath = null;
                    try { filePath = document.FullName; } catch { continue; }
                    if (string.IsNullOrEmpty(filePath)) { continue; }

                    bool isAssembly = document is SeAssembly;

                    if (data.ContainsKey(filePath))
                    {
                        data[filePath].OccurrenceCount++;
                    }
                    else
                    {
                        using PropertyUtils properties = new(document);
                        data[filePath] = new FileData
                        {
                            Name = Path.GetFileNameWithoutExtension(filePath),
                            OccurrenceCount = 1,
                            DxfDate = properties.DxfDate
                        };
                        if (isAssembly && !assemblyCache.ContainsKey(filePath)) assemblyCache[filePath] = properties.IsTypeA;
                    }

                    if (isAssembly && assemblyCache.TryGetValue(filePath, out bool isTypeA) && isTypeA)
                    {
                        SeOccurrences subOccurrences = null;
                        try
                        { 
                            subOccurrences = ((SeAssembly)document).Occurrences;
                            ProcessDataForExportPartsList(subOccurrences, data, assemblyCache); }
                        finally 
                        {
                            Helpers.ReleaseCom(ref subOccurrences); 
                        }
                    }
                }
                catch
                {
                    continue; 
                }
                finally
                { 
                    Helpers.ReleaseCom(ref document); Helpers.ReleaseCom(ref occurrence);
                }
            }
        }
        /*************************************************************************************************************************************************/






        /************************************************            EXPORT OCCURRENCES LIST               ***********************************************/
        public static void BuildDataForExportOccurrencesList(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data, List<string> types)
        {
            ProcessDataForExportOccurrencesList(assemblyOccurrences, data, types, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }
        private static void ProcessDataForExportOccurrencesList(SeOccurrences occurrences, Dictionary<string, FileData> data, List<string> types, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    if (occurrence.IncludeInBom == false) continue;
                    document = (SeDocument)occurrence.OccurrenceDocument;

                    string filePath = null;
                    try { filePath = document.FullName; }
                    catch { continue; }
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
                                Name = Path.GetFileNameWithoutExtension(filePath),
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
                        try
                        {
                            subOccurrences = ((SeAssembly)document).Occurrences;
                            ProcessDataForExportOccurrencesList(subOccurrences, data, types, assemblyCache);
                        }
                        finally
                        {
                            Helpers.ReleaseCom(ref subOccurrences);
                        }
                    }
                }
                catch
                {
                    continue;
                }
                finally
                {
                    Helpers.ReleaseCom(ref document); Helpers.ReleaseCom(ref occurrence);
                }
            }
        }
        /************************************************************************************************************************************************************/








        /****************************************************         SET COUNT PROPERTY       *********************************************************************/
        public static void BuildDataForSetCountProperty(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data)
        {
            ProcessDataForSetCountProperty(assemblyOccurrences, data, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }
        private static void ProcessDataForSetCountProperty(SeOccurrences occurrences, Dictionary<string, FileData> data, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    if (occurrence.IncludeInBom == false) continue;
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
                                Name = Path.GetFileNameWithoutExtension(filePath),
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
                        try 
                        {
                            subOccurrences = ((SeAssembly)document).Occurrences; 
                            ProcessDataForSetCountProperty(subOccurrences, data, assemblyCache); 
                        }
                        finally
                        { 
                            Helpers.ReleaseCom(ref subOccurrences); 
                        }
                    }
                }
                catch 
                { 
                    continue; 
                }
                finally
                { 
                    Helpers.ReleaseCom(ref document); 
                    Helpers.ReleaseCom(ref occurrence); 
                }
            }
        }
        /*****************************************************************************************************************************************************************/



        /*****************************************************************************************************************************************************************/
        public static void ApplyCounts(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data, int multiplier, HashSet<string> processed)
        {
            ProcessDataForApplyCounts(assemblyOccurrences, data, multiplier, processed, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }
        private static void ProcessDataForApplyCounts(SeOccurrences occurrences, Dictionary<string, FileData> data, int multiplier, HashSet<string> processed, Dictionary<string, bool> assemblyCache)
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
                        try 
                        { 
                            subOccurrences = ((SeAssembly)document).Occurrences; 
                            ProcessDataForApplyCounts(subOccurrences, data, multiplier, processed, assemblyCache);
                        }
                        finally 
                        { 
                            Helpers.ReleaseCom(ref subOccurrences);
                        }
                    }
                }
                catch
                {
                    continue; 
                }
                finally 
                { 
                    Helpers.ReleaseCom(ref document); 
                    Helpers.ReleaseCom(ref occurrence);
                }
            }
        }
        /**************************************************************/



        /**************************************************************/
        public static void ClearDxfDates(SeOccurrences assemblyOccurrences, HashSet<string> processed)
        {
            ProcessDataForClearDxfDates(assemblyOccurrences, processed, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }
        private static void ProcessDataForClearDxfDates(SeOccurrences occurrences, HashSet<string> processed, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    if (occurrence.IncludeInBom == false) continue; 

                    document = (SeDocument)occurrence.OccurrenceDocument;

                    string filePath = null;
                    try { filePath = document.FullName; } catch { continue; }
                    if (string.IsNullOrEmpty(filePath)) { continue; }

                    bool isAssembly = document is SeAssembly;

                    if (!processed.Contains(filePath))
                    {
                        using PropertyUtils properties = new(document);

                        if (document is SePart || document is SeSheetMetal)
                        {
                            if (properties.IsTypeB)
                            {
                                if (properties.HasDxfDate) properties.ClearDxfDate(); 
                            }
                        }

                        processed.Add(filePath);

                        if (isAssembly) assemblyCache[filePath] = properties.IsTypeA; 
                    }

                    if (isAssembly && assemblyCache.TryGetValue(filePath, out bool isTypeA) && isTypeA)
                    {
                        SeOccurrences subOccurrences = null;
                        try 
                        { 
                            subOccurrences = ((SeAssembly)document).Occurrences;
                            ProcessDataForClearDxfDates(subOccurrences, processed, assemblyCache); 
                        }
                        finally
                        { 
                            Helpers.ReleaseCom(ref subOccurrences); 
                        }
                    }
                }
                catch
                { 
                    continue;
                }
                finally 
                { 
                    Helpers.ReleaseCom(ref document); Helpers.ReleaseCom(ref occurrence);
                }
            }
        }
        /************************************************************************************************************************************************/









        /****************************************************          ORGANISE DRAWINGS              ******************************************************/
        public static void BuildDataForOrganiseDrawings(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data)
        {
            ProcessDataForOrganiseDrawings(assemblyOccurrences, data, new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
        }
        private static void ProcessDataForOrganiseDrawings(SeOccurrences occurrences, Dictionary<string, FileData> data, Dictionary<string, bool> assemblyCache)
        {
            int count = occurrences.Count;
            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument document = null;
                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
                    if (occurrence.IncludeInBom == false) continue;
                    document = (SeDocument)occurrence.OccurrenceDocument;

                    string filePath = null;
                    try { filePath = document.FullName; } catch { continue; }
                    if (string.IsNullOrEmpty(filePath)) { continue; }

                    bool isAssembly = document is SeAssembly;

                    if (data.ContainsKey(filePath))
                    {
                        continue;
                    }
                    else if (!isAssembly || !assemblyCache.ContainsKey(filePath))
                    {
                        using var properties = new PropertyUtils(document);
                        data[filePath] = new FileData
                        {
                            Name = Path.GetFileNameWithoutExtension(filePath),
                            Type = properties.Type
                        };

                        if (isAssembly) assemblyCache[filePath] = properties.IsTypeA; 
                    }

                   
                    if (isAssembly && assemblyCache.TryGetValue(filePath, out bool isTypeA) && isTypeA)
                    {
                        SeOccurrences subOccurrences = null;
                        try 
                        {
                            subOccurrences = ((SeAssembly)document).Occurrences; 
                            ProcessDataForOrganiseDrawings(subOccurrences, data, assemblyCache); 
                        }
                        finally 
                        { 
                            Helpers.ReleaseCom(ref subOccurrences);
                        }
                    }
                }
                catch 
                { 
                    continue; 
                }
                finally 
                { 
                    Helpers.ReleaseCom(ref document); Helpers.ReleaseCom(ref occurrence); 
                }
            }
        }
        /**************************************************************************************************************************************************/
    }
}