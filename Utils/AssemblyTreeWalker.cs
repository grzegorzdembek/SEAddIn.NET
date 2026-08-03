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

                    string fileName = Path.GetFileNameWithoutExtension(filePath);

                    using var properties = new PropertyProvider(document);

                    if (document is SePart || document is SeSheetMetal)
                    {
                        string type = properties.Type;

                        if (!(type == "B"))
                        {
                            logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Type}");
                            continue;
                        }

                        string material = properties.Material;

                        if (string.IsNullOrEmpty(material))
                        {
                            logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Material}");
                            continue;
                        }

                        string thickness = properties.Thickness;

                        if (string.IsNullOrEmpty(thickness))
                        {
                            logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Thickness}");
                            continue;
                        }

                        int status = properties.Status;

                        if (!(status == 0))
                        {
                            logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Status} set to available");
                            continue;
                        }

                        string sizeX = properties.SizeX;
                        string sizeY = properties.SizeY;
                        string color = properties.Color;
                        string finish = properties.Finish;
                        string dxfDate = properties.DxfDate;
                        string title = properties.TitleEng;

                        if (string.IsNullOrEmpty(title))
                        {
                            title = properties.TitlePl;
                        }

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
                                Title = title,
                                Color = color,
                                Finish = finish
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
                            logger.LogSkip(fileName, $"Missing Property {Constants.Properties.Type} set to Assembly");
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
                    logger.LogError("Unknown File", $"Error scanning assembly tree: {ex.Message}");
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
                SeOccurrence occurrence = null;
                SeDocument doc = null;

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

                    if (string.IsNullOrEmpty(path))
                    {
                        continue;
                    }

                    if (doc is SePart || doc is SeSheetMetal)
                    {
                        if (!data.ContainsKey(path))
                        {
                            data[path] = 1;
                        }
                        else
                        {
                            data[path]++;
                        }
                    }
                    else if (doc is SeAssembly asmDoc)
                    {
                        bool isTypeA = false;

                        using (var properties = new PropertyProvider(doc))
                        {
                            isTypeA = properties.IsTypeA;
                        }

                        if (!data.ContainsKey(path))
                        {
                            data[path] = 1;
                        }
                        else
                        {
                            data[path]++;
                        }

                        if (!isTypeA)
                        {
                            continue;
                        }

                        SeOccurrences subOccurrences = null;

                        try
                        {
                            subOccurrences = asmDoc.Occurrences;
                            BuildDataForExportPartsList(subOccurrences, data);
                        }
                        finally
                        {
                            CoreUtils.ReleaseCom(ref subOccurrences);
                        }
                    }
                }
                catch
                {
                    continue;
                }
                finally
                {
                    CoreUtils.ReleaseCom(ref doc);
                    CoreUtils.ReleaseCom(ref occurrence);
                }
            }
        }

        public static void BuildDataForSetCount(SeOccurrences assemblyOccurrences, Dictionary<string, FileData> data)
        {
            int count = assemblyOccurrences.Count;

            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument doc = null;

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

                    if (string.IsNullOrEmpty(path))
                    {
                        continue;
                    }

                    string name = Path.GetFileNameWithoutExtension(path);

                    using var properties = new PropertyProvider(doc);

                    if (doc is SePart || doc is SeSheetMetal)
                    {
                        if (!data.ContainsKey(path))
                        {
                            data[path] = new FileData
                            {
                                Name = name,
                                Type = properties.Type,
                                Count = properties.Count,
                                OccurrenceCount = 1
                            };
                        }
                        else
                        {
                            data[path].OccurrenceCount++;
                        }
                    }
                    else if (doc is SeAssembly asmDoc)
                    {
                        string subAsmPath = asmDoc.FullName;

                        if (string.IsNullOrEmpty(subAsmPath))
                        {
                            continue;
                        }

                        bool isTypeA = false;

                        if (!data.ContainsKey(subAsmPath))
                        {
                            data[subAsmPath] = new FileData
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
                            data[subAsmPath].OccurrenceCount++;
                            isTypeA = data[subAsmPath].Type == Constants.PartTypes.Assembly;
                        }

                        if (!isTypeA)
                        {
                            continue;
                        }

                        SeOccurrences subOccurrences = null;

                        try
                        {
                            subOccurrences = asmDoc.Occurrences;
                            BuildDataForSetCount(subOccurrences, data);
                        }
                        finally
                        {
                            CoreUtils.ReleaseCom(ref subOccurrences);
                        }
                    }
                }
                catch
                {
                    continue;
                }
                finally
                {
                    CoreUtils.ReleaseCom(ref doc);
                    CoreUtils.ReleaseCom(ref occurrence);
                }
            }
        }

        public static void ApplyCounts(SeOccurrences occurrences, Dictionary<string, FileData> data, int multiplier, HashSet<string> processed)
        {
            int count = occurrences.Count;

            for (int i = 1; i <= count; i++)
            {
                SeOccurrence occurrence = null;
                SeDocument doc = null;

                try
                {
                    occurrence = (SeOccurrence)occurrences.Item(i);
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

                    if (string.IsNullOrEmpty(path))
                    {
                        continue;
                    }

                    if (data.ContainsKey(path) && !processed.Contains(path))
                    {
                        using var properties = new PropertyProvider(doc);
                        properties.Count = data[path].OccurrenceCount * multiplier;
                        processed.Add(path);
                    }

                    if (doc is SeAssembly asmDoc)
                    {
                        SeOccurrences subOccurrences = null;

                        try
                        {
                            subOccurrences = asmDoc.Occurrences;
                            ApplyCounts(subOccurrences, data, multiplier, processed);
                        }
                        finally
                        {
                            CoreUtils.ReleaseCom(ref subOccurrences);
                        }
                    }
                }
                catch
                {
                    continue;
                }
                finally
                {
                    CoreUtils.ReleaseCom(ref doc);
                    CoreUtils.ReleaseCom(ref occurrence);
                }
            }
        }
    }
}